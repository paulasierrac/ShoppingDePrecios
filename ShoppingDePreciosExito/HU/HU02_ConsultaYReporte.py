"""
================================================================================
HU02 - Consulta y Reporte
Nombre de la iniciativa: Shopping de Precios Exito
Autor: KPMG Advisory, Tax & Legal SAS
Descripcion: Consulta precios en la web de Exito por EAN usando JavaScript,
             guarda los resultados en BD y genera el reporte Excel.
             Equivale al bot HU02_ConsultaYReporte de Automation Anywhere.
Ultima modificacion: 30/06/2026
Propiedad de Colsubsidio
================================================================================

Estados en TablaExito:
  1   : Pendiente de consultar
  2   : Producto encontrado
  3   : Sin coincidencia (titulo no corresponde al EAN)
  99  : Sin informacion (producto no aparece en la busqueda)
  100 : Consultado y reportado (fue Estado=2)
  199 : Consultado y reportado (fue Estado=99)

Flujo principal:
  1. Inserta en TablaExito los IDs nuevos que esten en TicketInsumo
     pero aun no en TablaExito.
  2. Verifica que existan registros pendientes; si no, termina.
  3. Crea estructura de carpetas de screenshots (anio/mes/dia).
  4. Bucle de scraping: extrae producto/precios via JS en exito.com
     hasta que no queden registros en Estado=1.
  5. Limpieza de puntos de miles en columnas de precio (formato colombiano).
  6. Generacion de reporte: por cada FechaInicio con registros procesados
     exporta un Excel y envia el correo de resultado.

Modo debug (in_config["_debug"] = True):
  - PASO 1: INSERT omitido; los EANs se toman directo de TicketInsumo.
  - PASO 4: Sin UPDATE en BD; resultados en memoria. Chrome visible.
            Solo procesa los primeros N EANs (LoteDebug, defecto=3).
  - PASO 5: UPDATE de precios omitido.
  - PASO 6: Excel generado en ./debug/ local; sin correo ni cambio de estado.

URL de busqueda: {UrlExito}  (ej: https://www.exito.com/s?q=REEMPLAZAR&sort=score_desc&page=0)
"""

import os
import re
import sys
import time
import socket
import winreg
from datetime import datetime
from pathlib import Path

import pandas as pd

from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeout

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
from Funciones.utils import write_log, conectar_bd, conectar_bd_debug, csv_a_excel, enviar_correo


# ============================================================
# Tiempos de espera (milisegundos para Playwright)
# ============================================================
ESPERA_3S = 3000
ESPERA_5S = 5000

# Numero de productos a procesar por lote de DB (modo normal)
_LOTE_DEFAULT = 50


# ============================================================
# Helpers generales
# ============================================================

def _proxy_sistema_windows() -> dict:
    """Lee la configuracion de proxy del registro de Windows."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        )
        proxy_enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
        if not proxy_enable:
            return None
        proxy_server, _ = winreg.QueryValueEx(key, "ProxyServer")
        winreg.CloseKey(key)
        if proxy_server:
            if not proxy_server.startswith("http"):
                proxy_server = f"http://{proxy_server}"
            return {"server": proxy_server}
    except Exception:
        pass
    return None


def _asegurar_chromium(in_config: dict, task_name: str) -> None:
    """Descarga Playwright Chromium si no existe para el usuario actual."""
    import subprocess
    try:
        with sync_playwright() as _pw:
            exec_path = _pw.chromium.executable_path
        if os.path.isfile(exec_path):
            return
    except Exception:
        pass
    write_log("Info", "HU02: Playwright Chromium no encontrado — descargando...", task_name, in_config)
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    write_log("Info", "HU02: Playwright Chromium instalado correctamente", task_name, in_config)


def _js(page: Page, script: str, default=""):
    """Ejecuta JavaScript via Playwright; retorna `default` si falla o devuelve None."""
    try:
        result = page.evaluate(script)
        return result if result is not None else default
    except Exception:
        return default


def _entre(texto: str, antes: str, despues: str) -> str:
    """
    Extrae la subcadena que aparece entre 'antes' y 'despues' en texto.
    Equivale a la operacion Before/After de AA.
    """
    if not texto or antes not in texto:
        return ""
    start = texto.index(antes) + len(antes)
    rest  = texto[start:]
    if despues not in rest:
        return rest.strip()
    return rest[: rest.index(despues)].strip()


def _limpiar_precio(texto: str) -> str:
    """Elimina simbolos de moneda, espacios y puntos de miles. Ej: '$ 15.000' -> '15000'."""
    if not texto:
        return ""
    return re.sub(r"[^\d]", "", texto)


def _tomar_screenshot(page: Page, ruta: str) -> None:
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        page.screenshot(path=ruta)
    except Exception:
        pass


# ============================================================
# Logica de scraping por EAN en Exito
# ============================================================

def _consultar_ean_exito(page: Page, ean: str, palabra_clave: str,
                         url_template: str, ruta_screenshot: str,
                         in_config: dict, task_name: str) -> dict:
    """
    Navega a la URL de busqueda de Exito para el EAN, extrae los datos del
    primer resultado usando JavaScript sobre el DOM, y retorna un dict con:
      nombre_prd, marca, precio_fidelizacion, precio_con_desc,
      precio_sin_desc, porc_descuento, precio_unitario,
      url_producto, banner, estado, observaciones
    """
    resultado = {
        "nombre_prd":         "",
        "marca":              "",
        "precio_fidelizacion":"",
        "precio_con_desc":    "",
        "precio_sin_desc":    "",
        "porc_descuento":     "",
        "precio_unitario":    "",
        "url_producto":       "",
        "banner":             "",
        "estado":             "99",
        "observaciones":      "No existe el producto en la farmacia",
    }

    url_consulta = url_template.replace("REEMPLAZAR", ean)

    try:
        page.goto(url_consulta, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(ESPERA_5S)

        # ── Paso 1: Obtener HTML de todas las tarjetas de producto (3 intentos) ──
        html_cards = ""
        for intento in range(3):
            try:
                html_cards = page.evaluate(
                    "Array.from(document.getElementsByClassName("
                    "'productCard_productCard__M0677 productCard_column__Lp3OF'))"
                    ".map(el => el.innerHTML).join(' ');"
                ) or ""
                if html_cards:
                    break
            except Exception as e:
                err_str = str(e)
                if "Unable to run the JavaScript" in err_str:
                    break
                try:
                    page.reload(wait_until="domcontentloaded", timeout=30000)
                except Exception:
                    pass
                page.wait_for_timeout(ESPERA_5S)

        if not html_cards:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — No existe el producto en la farmacia",
                task_name, in_config
            )
            _tomar_screenshot(page, ruta_screenshot)
            return resultado

        # ── Paso 2: Separar en bloques individuales de producto ──────────────
        bloque = ""
        for delimitador in [
            "productCard_contentInfo__CBBA7 productCard_column__Lp3OF",
            '<div _ngcontent-ng-ftd-c94=""',
            'app-new-product-card _ngcontent-ng-ftd-c97=""',
        ]:
            if delimitador in html_cards:
                partes = html_cards.split(delimitador)
                if len(partes) > 1:
                    bloque = partes[1]
                break

        if not bloque:
            bloque = html_cards

        # ── Paso 3: Extraer URL del primer producto ───────────────────────────
        url_path = _entre(html_cards, 'a href="/', '"')
        url_base_site = url_template.split("/s?q=")[0]
        url_producto = f"{url_base_site}/{url_path}" if url_path else url_consulta
        resultado["url_producto"] = url_producto

        # ── Paso 4: Extraer campos del primer bloque ──────────────────────────
        nombre_prd = _entre(bloque, 'class="styles_name__qQJiK">', "<")
        marca      = _entre(bloque, 'class="styles_brand__IdJcB">', "<")

        precio_fid_raw  = _entre(bloque, 'class="price_fs-price__4GZ9F ">', "<")
        precio_con_raw  = _entre(
            bloque,
            'class="ProductPrice_container__price__XmMWA ProductPrice_text14___ZxlL">',
            "<",
        )
        precio_sin_raw  = _entre(
            bloque,
            'class="priceSection_container-promotion_price-dashed__FJ7nI">',
            "<",
        )
        porc_desc_raw   = _entre(bloque, '<span data-percentage="true">', "<")
        precio_unit_raw = _entre(
            bloque,
            '<span class="product-unit_price-unit__text__qeheS">',
            "<",
        )

        if not nombre_prd:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — No existe el producto en la farmacia",
                task_name, in_config
            )
            _tomar_screenshot(page, ruta_screenshot)
            return resultado

        # ── Paso 5: Validar que el nombre corresponda a la palabra clave ──────
        nombre_upper = nombre_prd.upper()
        kw_upper     = (palabra_clave or "").upper().strip()

        if kw_upper and kw_upper not in nombre_upper:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — Sin coincidencia: nombre='{nombre_prd}', "
                f"palabra_clave='{palabra_clave}'",
                task_name, in_config
            )
            _tomar_screenshot(page, ruta_screenshot)
            resultado.update({
                "nombre_prd":    nombre_prd,
                "marca":         marca,
                "url_producto":  url_producto,
                "estado":        "3",
                "observaciones": (
                    "No existe coincidencia entre la informacion encontrada "
                    "y el producto consultado"
                ),
            })
            return resultado

        write_log(
            "Info",
            f"HU02: EAN ({ean}) — Producto encontrado: '{nombre_prd}'",
            task_name, in_config
        )
        _tomar_screenshot(page, ruta_screenshot)

        resultado.update({
            "nombre_prd":          nombre_prd,
            "marca":               marca,
            "precio_fidelizacion": _limpiar_precio(precio_fid_raw),
            "precio_con_desc":     _limpiar_precio(precio_con_raw),
            "precio_sin_desc":     _limpiar_precio(precio_sin_raw),
            "porc_descuento":      porc_desc_raw.replace("%", "").strip(),
            "precio_unitario":     precio_unit_raw,
            "url_producto":        url_producto,
            "banner":              "",
            "estado":              "2",
            "observaciones":       "",
        })

    except PlaywrightTimeout:
        write_log("Warning", f"HU02: Timeout consultando EAN ({ean})", task_name, in_config)
        resultado["estado"]        = "99"
        resultado["observaciones"] = "Timeout al cargar la pagina"
    except Exception as e:
        write_log(
            "Warning",
            f"HU02: Error consultando EAN ({ean}): {e}",
            task_name, in_config
        )
        resultado["estado"]        = "99"
        resultado["observaciones"] = f"Error: {e}"

    return resultado


# ============================================================
# Funcion principal
# ============================================================

def hu02_consulta_y_reporte(in_config: dict) -> str:
    """
    Ejecuta la consulta web y generacion de reporte para Exito.
    Retorna '' si exitoso, mensaje de error si fallo.
    """
    out_system_exception = ""
    task_name = "HU02_ConsultaYReporte"
    debug     = in_config.get("_debug", False)

    write_log("Info", "Inicia HU02", task_name, in_config)
    if debug:
        write_log("Info", "[DEBUG] Modo debug activo: sin escrituras en BD ni correos", task_name, in_config)

    pw_instance = None
    browser     = None

    try:
        esquema      = in_config.get("Scheme", "[ShoppingDePrecios]")
        tabla_ex     = in_config.get("TablaExito",        "Exito")
        tabla_ins    = in_config.get("TablaTicketInsumo", "TicketInsumo")
        url_template = in_config.get("UrlExito", "")
        maquina      = socket.gethostname()

        # ----------------------------------------------------------------
        # PASO 1 + PASO 2: Verificar registros (debug → SQLite / normal → SQL Server)
        # ----------------------------------------------------------------
        if debug:
            conn_sq = conectar_bd_debug(in_config)
            cur_sq  = conn_sq.cursor()
            cur_sq.execute(f"SELECT COUNT(*) FROM {esquema}.{tabla_ins} WHERE Estado=1")
            cnt = cur_sq.fetchone()[0] or 0
            conn_sq.close()
            hay_pendientes = cnt > 0
            if hay_pendientes:
                write_log("Info",
                          f"[DEBUG] {cnt} registros en BD Dev ({tabla_ins}) con Estado=1",
                          task_name, in_config)
            else:
                write_log("Info",
                          f"[DEBUG] No hay registros en BD Dev ({tabla_ins}) con Estado=1",
                          task_name, in_config)
        else:
            conn   = conectar_bd(in_config)
            cursor = conn.cursor()

            # PASO 1: Insertar en TablaExito los IDs nuevos de TicketInsumo
            cursor.execute(f"""
                SELECT a.Id
                FROM {esquema}.{tabla_ins} a
                LEFT JOIN {esquema}.{tabla_ex} b ON a.Id = b.Id
                WHERE b.Id IS NULL AND a.Estado='1'
            """)
            if cursor.fetchone() is not None:
                cursor.execute(f"SET IDENTITY_INSERT {esquema}.{tabla_ex} ON")
                cursor.execute(f"""
                    INSERT INTO {esquema}.{tabla_ex}
                        ([Id],[FechaInicio],[FechaModificacion],[FechaFin],
                         [Estado],[Maquina],[PLU],[EAN],[Descripcion],
                         [MarcaProducto],[NombrePrd],[RegistroInvima],
                         [PrecioUnitario],[PrecioConDescuento],[PrecioSinDescuento],
                         [Porc.Descuento],[PrecioFidelizacion],[UrlProducto],[BannerProducto],[RutaImagen])
                    SELECT
                        a.[Id], a.[FechaInicio], GETDATE(), '',
                        '1', '{maquina}', a.[PLU], a.[EAN], a.[Descripcion],
                        '','','','','','','','','','',''
                    FROM {esquema}.{tabla_ins} a
                    LEFT JOIN {esquema}.{tabla_ex} b ON a.Id = b.Id
                    WHERE b.Id IS NULL AND a.Estado='1'
                """)
                cursor.execute(f"SET IDENTITY_INSERT {esquema}.{tabla_ex} OFF")
                write_log("Info",
                          f"HU02: Existen nuevos registros para cargar a la tabla ({tabla_ex})",
                          task_name, in_config)

            # PASO 2: Verificar si hay registros pendientes
            cursor.execute(f"""
                SELECT TOP(1) 1 FROM {esquema}.{tabla_ex}
                WHERE Estado='1' OR Estado='2' OR Estado='99'
            """)
            hay_pendientes = cursor.fetchone() is not None
            conn.commit()
            conn.close()

        if not hay_pendientes:
            write_log("Info", "HU02: No existen registros para consultar en pagina", task_name, in_config)
            write_log("Info", "Finaliza HU02", task_name, in_config)
            return ""

        write_log("Info", "HU02: Existen registros que requieren consulta en la pagina", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 3: Estructura de carpetas de screenshots
        # ----------------------------------------------------------------
        now = datetime.now()
        if debug:
            ruta_screenshots = str(
                _PROJECT_ROOT / "debug" / "screenshots" / "Exito"
                / str(now.year) / f"{now.month:02d}" / f"{now.day:02d}"
            )
        else:
            ruta_screenshots = os.path.join(
                in_config.get("RutaScreenshots", ""),
                in_config.get("CarpetaExito", "Exito\\"),
                str(now.year),
                f"{now.month:02d}",
                f"{now.day:02d}",
            )
        os.makedirs(ruta_screenshots, exist_ok=True)

        # ----------------------------------------------------------------
        # PASO 4: Bucle principal de scraping
        # ----------------------------------------------------------------
        headless  = False if debug else str(in_config.get("HeadlessChrome", "true")).lower() == "true"
        proxy_cfg = _proxy_sistema_windows()

        write_log("Info", "HU02: Inicia consulta de productos por EAN", task_name, in_config)

        _asegurar_chromium(in_config, task_name)
        pw_instance = sync_playwright().start()
        browser = pw_instance.chromium.launch(
            headless=headless,
            proxy=proxy_cfg,
            args=[
                "--lang=es-CO",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--disable-software-rasterizer",
            ]
        )

        if debug:
            _ejecutar_scraping_debug(
                browser, in_config, esquema, tabla_ins, url_template,
                ruta_screenshots, task_name
            )
        else:
            _ejecutar_scraping_normal(
                browser, in_config, esquema, tabla_ex, tabla_ins, url_template,
                ruta_screenshots, maquina, task_name
            )

        write_log("Info", "HU02: Termina consulta de productos por EAN", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 5: Limpieza de puntos de miles en columnas de precio
        # ----------------------------------------------------------------
        if not debug:
            conn   = conectar_bd(in_config)
            cursor = conn.cursor()
            cursor.execute(f"""
                UPDATE {esquema}.{tabla_ex}
                SET [PrecioConDescuento] = REPLACE([PrecioConDescuento], '.', '')
                WHERE Estado='2' OR Estado='100'
            """)
            cursor.execute(f"""
                UPDATE {esquema}.{tabla_ex}
                SET [PrecioSinDescuento] = REPLACE([PrecioSinDescuento], '.', '')
                WHERE Estado='2' OR Estado='100'
            """)
            conn.commit()
            conn.close()

        # ----------------------------------------------------------------
        # PASO 6: Generacion de reportes
        # ----------------------------------------------------------------
        if not debug:
            _generar_reportes(in_config, esquema, tabla_ex, task_name)

        write_log("Info", "Finaliza HU02", task_name, in_config)

    except Exception as e:
        out_system_exception = str(e)
        write_log("Error", f"HU02: {e}", task_name, in_config)
        write_log("Info", "Finaliza HU02", task_name, in_config)

    finally:
        if browser:
            try:
                browser.close()
            except Exception:
                pass
        if pw_instance:
            try:
                pw_instance.stop()
            except Exception:
                pass

    return out_system_exception


# ============================================================
# Scraping modo normal (con escritura en BD)
# ============================================================

def _ejecutar_scraping_normal(browser, in_config, esquema, tabla_ex, tabla_ins,
                               url_template, ruta_screenshots, maquina, task_name):
    lote  = int(in_config.get("CantExito", str(_LOTE_DEFAULT)))
    delay = int(in_config.get("SegExito",  "5"))

    hay_mas = True
    while hay_mas:
        conn   = conectar_bd(in_config)
        cursor = conn.cursor()

        cursor.execute(f"""
            SELECT TOP({lote})
                [Id], [EAN],
                LEFT(
                    LTRIM(SUBSTRING(Descripcion,
                        PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100)),
                    CHARINDEX(' ',
                        LTRIM(SUBSTRING(Descripcion,
                            PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100))
                        + ' ') - 1
                )
            FROM {esquema}.{tabla_ex}
            WHERE Estado='1'
        """)
        registros = cursor.fetchall()
        conn.close()

        if not registros:
            hay_mas = False
            break

        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            locale="es-CO",
            viewport={"width": 1920, "height": 1080},
            ignore_https_errors=True,
        )
        page = context.new_page()

        try:
            for row in registros:
                id_ticket     = str(row[0])
                ean           = str(row[1])
                palabra_clave = str(row[2] or "")

                conn   = conectar_bd(in_config)
                cursor = conn.cursor()
                cursor.execute(f"""
                    UPDATE {esquema}.{tabla_ex}
                    SET FechaModificacion=GETDATE()
                    WHERE Id='{id_ticket}'
                """)
                conn.commit()
                conn.close()

                ruta_ss = os.path.join(ruta_screenshots, f"{ean}_{id_ticket}.jpg")

                write_log(
                    "Info",
                    f"HU02: Se consultara el EAN ({ean}) en la ruta "
                    f"({url_template.replace('REEMPLAZAR', ean)})",
                    task_name, in_config
                )

                res = _consultar_ean_exito(
                    page=page,
                    ean=ean,
                    palabra_clave=palabra_clave,
                    url_template=url_template,
                    ruta_screenshot=ruta_ss,
                    in_config=in_config,
                    task_name=task_name,
                )

                conn   = conectar_bd(in_config)
                cursor = conn.cursor()

                estado      = res["estado"]
                nombre_prd  = res["nombre_prd"].replace(";", "").replace("'", "''")
                marca       = res["marca"].replace("'", "''")
                precio_fid  = res["precio_fidelizacion"]
                precio_con  = res["precio_con_desc"]
                precio_sin  = res["precio_sin_desc"]
                porc_desc   = res["porc_descuento"]
                precio_unit = res["precio_unitario"].replace("'", "''")
                url_prd     = res["url_producto"].replace("'", "''")
                ruta_img    = ruta_ss.replace("'", "''")

                if estado == "99":
                    cursor.execute(f"""
                        UPDATE {esquema}.{tabla_ex}
                        SET [FechaFin]=GETDATE(),
                            [Estado]='99',
                            [UrlProducto]='{url_prd}',
                            [RutaImagen]='{ruta_img}'
                        WHERE Id='{id_ticket}'
                    """)
                elif estado == "3":
                    cursor.execute(f"""
                        UPDATE {esquema}.{tabla_ex}
                        SET [FechaFin]=GETDATE(),
                            [Estado]='3',
                            [NombrePrd]='{nombre_prd}',
                            [MarcaProducto]='{marca}',
                            [UrlProducto]='{url_prd}',
                            [RutaImagen]='{ruta_img}'
                        WHERE Id='{id_ticket}'
                    """)
                else:
                    cursor.execute(f"""
                        UPDATE {esquema}.{tabla_ex}
                        SET [FechaFin]=GETDATE(),
                            [Estado]='2',
                            [NombrePrd]='{nombre_prd}',
                            [MarcaProducto]='{marca}',
                            [PrecioFidelizacion]='{precio_fid}',
                            [PrecioConDescuento]='{precio_con}',
                            [PrecioSinDescuento]='{precio_sin}',
                            [Porc.Descuento]='{porc_desc}',
                            [PrecioUnitario]='{precio_unit}',
                            [BannerProducto]='{res["banner"]}',
                            [UrlProducto]='{url_prd}',
                            [RutaImagen]='{ruta_img}'
                        WHERE Id='{id_ticket}'
                    """)

                conn.commit()
                conn.close()

        finally:
            try:
                context.close()
            except Exception:
                pass

        conn   = conectar_bd(in_config)
        cursor = conn.cursor()
        cursor.execute(f"SELECT TOP(1) 1 FROM {esquema}.{tabla_ex} WHERE Estado='1'")
        hay_mas = cursor.fetchone() is not None
        conn.close()

        if hay_mas and delay > 0:
            write_log("Info", f"HU02: Esperando {delay}s antes del siguiente lote", task_name, in_config)
            time.sleep(delay)


# ============================================================
# Scraping modo debug (sin escritura en BD)
# ============================================================

def _ejecutar_scraping_debug(browser, in_config, esquema, tabla_ins,
                              url_template, ruta_screenshots, task_name):
    """
    Lee EANs de pruebas.db (TicketInsumo), hace el scraping, escribe
    resultados en pruebas.db (Exito) y genera Excel en ./debug/.
    No toca SQL Server.
    """
    lote_debug = int(in_config.get("LoteDebug", "3"))

    conn_sq = conectar_bd_debug(in_config)
    cur_sq  = conn_sq.cursor()
    cur_sq.execute(
        f"SELECT Id, EAN, Descripcion FROM {esquema}.TicketInsumo WHERE Estado=1 LIMIT ?",
        (lote_debug,)
    )
    registros = cur_sq.fetchall()

    if not registros:
        write_log("Info", "[DEBUG] No hay registros en pruebas.db TicketInsumo con Estado=1", task_name, in_config)
        conn_sq.close()
        return

    write_log("Info", f"[DEBUG] Procesando {len(registros)} EAN(s) (LoteDebug={lote_debug})", task_name, in_config)
    print(f"\n{'─'*70}")
    print(f"  [DEBUG] Procesando {len(registros)} EAN(s)")
    print(f"{'─'*70}")

    resultados = []

    context = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        locale="es-CO",
        viewport={"width": 1920, "height": 1080},
        ignore_https_errors=True,
    )
    page = context.new_page()

    try:
        for row in registros:
            id_ticket   = str(row[0])
            ean         = str(row[1])
            descripcion = str(row[2] or "")
            m = re.search(r'[a-zA-Z]{3,}', descripcion)
            palabra_clave = m.group(0) if m else ""
            ruta_ss       = os.path.join(ruta_screenshots, f"{ean}_{id_ticket}.jpg")

            write_log(
                "Info",
                f"[DEBUG] Consultando EAN ({ean}) — {descripcion[:60]}",
                task_name, in_config
            )
            print(f"\n  EAN: {ean}  |  Descripcion: {descripcion[:50]}")

            res = _consultar_ean_exito(
                page=page,
                ean=ean,
                palabra_clave=palabra_clave,
                url_template=url_template,
                ruta_screenshot=ruta_ss,
                in_config=in_config,
                task_name=task_name,
            )

            estado_label = {
                "2":  "ENCONTRADO",
                "3":  "SIN COINCIDENCIA",
                "99": "SIN INFORMACION",
            }.get(res["estado"], res["estado"])

            print(f"  Estado    : {res['estado']} ({estado_label})")
            if res["nombre_prd"]:
                print(f"  Nombre    : {res['nombre_prd']}")
            if res["marca"]:
                print(f"  Marca     : {res['marca']}")
            if res["precio_con_desc"]:
                print(f"  Precio    : {res['precio_con_desc']}")
            if res["precio_sin_desc"]:
                print(f"  Precio s/d: {res['precio_sin_desc']}")
            if res["url_producto"]:
                print(f"  URL       : {res['url_producto']}")
            if res["observaciones"]:
                print(f"  Obs       : {res['observaciones']}")

            resultados.append({
                "Id":                 id_ticket,
                "EAN":                ean,
                "Descripcion":        descripcion,
                "Estado":             res["estado"],
                "NombrePrd":          res["nombre_prd"],
                "MarcaProducto":      res["marca"],
                "PrecioConDescuento": res["precio_con_desc"],
                "PrecioSinDescuento": res["precio_sin_desc"],
                "PrecioFidelizacion": res["precio_fidelizacion"],
                "Porc.Descuento":     res["porc_descuento"],
                "PrecioUnitario":     res["precio_unitario"],
                "UrlProducto":        res["url_producto"],
                "RutaImagen":         ruta_ss,
                "Observaciones":      res["observaciones"],
            })

    finally:
        try:
            context.close()
        except Exception:
            pass

    # Guardar resultados en pruebas.db (tabla Exito)
    if resultados:
        ahora   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        maquina = socket.gethostname()
        cur_sq.execute(f"DELETE FROM {esquema}.Exito")
        for r in resultados:
            cur_sq.execute(
                f"INSERT INTO {esquema}.Exito "
                "(FechaInicio, FechaModificacion, FechaFin, Estado, Reintentos, Maquina, "
                " PLU, EAN, Descripcion, Categoria, HoraConsulta, MarcaProducto, NombrePrd, RegistroInvima, "
                " PrecioUnitario, PrecioConDescuento, PrecioSinDescuento, [Porc.Descuento], PrecioFidelizacion, "
                " BannerProducto, UrlProducto, RutaImagen) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (ahora, ahora, ahora,
                 r["Estado"], 0, maquina,
                 "", r["EAN"], r["Descripcion"], "", ahora,
                 r["MarcaProducto"], r["NombrePrd"], "",
                 r["PrecioUnitario"], r["PrecioConDescuento"], r["PrecioSinDescuento"],
                 r["Porc.Descuento"], r["PrecioFidelizacion"],
                 "", r["UrlProducto"], r["RutaImagen"])
            )
        conn_sq.commit()
        write_log("Info",
                  f"[DEBUG] {len(resultados)} registros guardados en BD Dev ({esquema}.Exito)",
                  task_name, in_config)
    conn_sq.close()

    # Generar Excel de debug
    _generar_reporte_debug(in_config, resultados, task_name)

    print(f"\n{'─'*70}")
    print(f"  [DEBUG] Resumen: {sum(1 for r in resultados if r['Estado']=='2')} encontrados, "
          f"{sum(1 for r in resultados if r['Estado']=='99')} sin info, "
          f"{sum(1 for r in resultados if r['Estado']=='3')} sin coincidencia")
    print(f"{'─'*70}\n")


def _generar_reporte_debug(in_config: dict, resultados: list, task_name: str) -> None:
    """Genera un Excel local en ./debug/ con los resultados del modo debug."""
    if not resultados:
        write_log("Info", "[DEBUG] Sin resultados para reporte", task_name, in_config)
        return

    ruta_debug = _PROJECT_ROOT / "debug"
    ruta_debug.mkdir(exist_ok=True)

    sello        = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_excel = f"DEBUG_ReportePricingExito_{sello}.xlsx"
    ruta_excel   = str(ruta_debug / nombre_excel)

    df = pd.DataFrame(resultados)
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Debug", index=False)

    write_log("Info", f"[DEBUG] Reporte generado en ({ruta_excel})", task_name, in_config)
    print(f"\n  Reporte debug: {ruta_excel}")


# ============================================================
# Generacion de reportes Excel (modo normal)
# ============================================================

def _generar_reportes(in_config: dict, esquema: str,
                      tabla_ex: str, task_name: str) -> None:
    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    cursor.execute(f"""
        SELECT DISTINCT(FechaInicio),
            REPLACE(CONCAT(
                REPLACE(CAST(FechaInicio AS DATE),'-','_'),
                REPLACE(REPLACE(SUBSTRING(CAST(FechaInicio AS varchar),12,6),' ','_'),':','_')
            ),'__','_0')
        FROM {esquema}.{tabla_ex}
        WHERE Estado='2' OR Estado='99'
    """)
    fechas = cursor.fetchall()
    conn.close()

    for fecha_row in fechas:
        fecha_inicio = str(fecha_row[0])
        fecha_sello  = str(fecha_row[1])
        _generar_reporte_fecha(in_config, esquema, tabla_ex,
                               fecha_inicio, fecha_sello, task_name)


def _generar_reporte_fecha(in_config: dict, esquema: str, tabla_ex: str,
                           fecha_inicio: str, fecha_sello: str,
                           task_name: str) -> None:
    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    # Restaurar estados intermedios si los hubiera
    cursor.execute(
        f"UPDATE {esquema}.{tabla_ex} SET [Estado]='2' "
        f"WHERE [Estado]='100' AND FechaInicio='{fecha_inicio}'"
    )
    cursor.execute(
        f"UPDATE {esquema}.{tabla_ex} SET [Estado]='99' "
        f"WHERE [Estado]='199' AND FechaInicio='{fecha_inicio}'"
    )

    # Estadisticas
    cursor.execute(f"""
        SELECT
            COUNT(*)                                                             AS TotalRegistros,
            SUM(CASE WHEN Estado IN ('2','100') THEN 1 ELSE 0 END)              AS CantidadExtraidos,
            SUM(CASE WHEN Estado IN ('99','199') THEN 1 ELSE 0 END)             AS CantidadEstado99
        FROM {esquema}.{tabla_ex}
        WHERE FechaInicio='{fecha_inicio}'
    """)
    stats          = cursor.fetchone()
    total          = stats[0] if stats else 0
    extraidos      = stats[1] if stats else 0
    estado99_count = stats[2] if stats else 0

    write_log(
        "Info",
        f"HU02: Reporte FechaInicio={fecha_inicio} — "
        f"Total={total}, Extraidos={extraidos}, Estado99={estado99_count}",
        task_name, in_config
    )

    # Marcar como reportados
    cursor.execute(
        f"UPDATE {esquema}.{tabla_ex} SET [Estado]='100' "
        f"WHERE [Estado]='2' AND FechaInicio='{fecha_inicio}'"
    )
    cursor.execute(
        f"UPDATE {esquema}.{tabla_ex} SET [Estado]='199' "
        f"WHERE [Estado]='99' AND FechaInicio='{fecha_inicio}'"
    )

    # Calcular porcentaje de descuento
    cursor.execute(f"""
        UPDATE {esquema}.{tabla_ex}
        SET [Porc.Descuento] =
            ((TRY_CAST([PrecioSinDescuento] AS INT) - TRY_CAST([PrecioConDescuento] AS INT)) * 100)
            / TRY_CAST([PrecioSinDescuento] AS INT)
        WHERE [Estado]='100'
          AND FechaInicio='{fecha_inicio}'
          AND TRY_CAST([PrecioSinDescuento] AS INT) > 0
          AND TRY_CAST([PrecioConDescuento] AS INT) > 0
          AND TRY_CAST([PrecioSinDescuento] AS INT) != TRY_CAST([PrecioConDescuento] AS INT)
    """)
    conn.commit()

    # Consultar datos para el reporte
    cursor.execute(f"""
        SELECT
            [FechaInicio]        AS fechainsumo,
            [PLU]                AS plu,
            [Descripcion]        AS descripcion,
            [FechaModificacion]  AS horaconsulta,
            [EAN]                AS ean,
            [Estado]             AS estado,
            [MarcaProducto]      AS marcaproducto,
            [NombrePrd]          AS nombreprd,
            [RegistroInvima]     AS registroinvima,
            [PrecioUnitario]     AS preciounitario,
            [PrecioConDescuento] AS preciocondescuento,
            [PrecioSinDescuento] AS preciosindescuento,
            [Porc.Descuento]     AS [porc.descuento],
            [PrecioFidelizacion] AS preciofidelizacion,
            [BannerProducto]     AS bannerproducto,
            [UrlProducto]        AS urlproducto,
            [RutaImagen]         AS rutaimagen
        FROM {esquema}.{tabla_ex}
        WHERE FechaInicio='{fecha_inicio}'
    """)
    cols  = [col[0] for col in cursor.description]
    filas = cursor.fetchall()
    conn.close()

    if not filas:
        return

    df = pd.DataFrame(filas, columns=cols)

    ruta_reporte     = in_config.get("RutaReporte", "")
    os.makedirs(ruta_reporte, exist_ok=True)
    nombre_resultado = in_config.get("NombreResultado",     "ReportePricingExito_")
    nombre_hoja      = in_config.get("NombreHojaResultado", "ReportePricingExito")
    nombre_excel     = f"{nombre_resultado}{fecha_sello}.xlsx"
    ruta_excel       = os.path.join(ruta_reporte, nombre_excel)

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)

    write_log(
        "Info",
        f"HU02: Reporte generado en ({ruta_excel})",
        task_name, in_config
    )

    from_address = in_config.get("_correo", {}).get("usuario", "")
    reemplazo    = {"$NombrePagina$": in_config.get("DrogueriaExito", "Exito")}
    err = enviar_correo(
        in_config=in_config,
        i_cod_email=100,
        i_from_address=from_address,
        i_replace_in_message=reemplazo,
        i_replace_in_subject=reemplazo,
        i_html_format=False,
        i_attachment=[ruta_excel],
    )
    if err:
        write_log(
            "Info",
            f"HU02: No fue posible enviar el correo de notificacion: {err}",
            task_name, in_config
        )


if __name__ == "__main__":
    from Funciones.utils import obtener_config
    config = obtener_config()
    exc = hu02_consulta_y_reporte(config)
    if exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    print("HU02 completada exitosamente.")
