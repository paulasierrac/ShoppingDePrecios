"""
================================================================================
HU02 - Consulta y Reporte
Nombre de la iniciativa: Shopping de Precios Locatel
Autor: Paula Sierra — Net Applications
Descripcion: Consulta precios en la web de Locatel por EAN, guarda los
             resultados en BD y genera el reporte Excel.
             Equivale al bot HU02_ConsultaYReporte de Automation Anywhere.
Ultima modificacion: 27/05/2025
Propiedad de Colsubsidio
================================================================================

Estados en TablaLocatel:
  1   : Pendiente de consultar
  2   : Producto encontrado
  3   : Sin coincidencia (titulo no corresponde al EAN)
  99  : Sin informacion (producto no aparece en la busqueda)
  100 : Consultado y reportado (fue Estado=2)
  199 : Consultado y reportado (fue Estado=99)

Flujo principal:
  1. Reprocesa registros "Sin stock" que aun tienen reintentos disponibles.
  2. Inserta en TablaLocatel los IDs nuevos que esten en TicketInsumo
     pero aun no en TablaLocatel.
  3. Verifica que existan registros pendientes; si no, termina.
  4. Crea estructura de carpetas de screenshots (anio/mes/dia).
  5. Bucle de scraping: procesa lotes de LoteLocatel productos con Playwright
     hasta que no queden registros en Estado=1.
  6. Generacion de reporte: por cada FechaInicio con registros procesados
     exporta un Excel y envia el correo de resultado.
"""

import os
import re
import sys
import time
import socket
import shutil
import traceback
import winreg
from datetime import datetime
from pathlib import Path

import pandas as pd

from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeout

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
from Funciones.utils import write_log, conectar_bd, csv_a_excel, enviar_correo


# ============================================================
# Tiempos de espera (milisegundos para Playwright)
# ============================================================
ESPERA_3S = 3000
ESPERA_5S = 5000
ESPERA_7S = 7000


# ============================================================
# Helpers
# ============================================================

def _proxy_sistema_windows() -> str:
    """Retorna el proxy configurado en Windows Internet Settings, o '' si no hay."""
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        ) as key:
            if winreg.QueryValueEx(key, "ProxyEnable")[0]:
                proxy = winreg.QueryValueEx(key, "ProxyServer")[0]
                if proxy and "://" not in proxy:
                    proxy = f"http://{proxy}"
                return proxy or ""
    except Exception:
        pass
    return ""


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
    """Ejecuta JavaScript en la pagina y retorna el resultado."""
    try:
        result = page.evaluate(script)
        return result if result is not None else default
    except Exception:
        return default


def _tomar_screenshot(page: Page, ruta: str) -> None:
    """Toma screenshot de la pagina actual."""
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        page.screenshot(path=ruta)
    except Exception:
        pass


def _extraer_precio_entero(texto: str) -> str:
    """
    Limpia un texto de precio: elimina espacios, puntos de miles y signo $.
    Ejemplo: "$ 15.000" -> "15000"
    """
    if not texto:
        return ""
    return re.sub(r"[^\d]", "", texto)


def _cerrar_modales_locatel(page: Page, task_name: str, in_config: dict) -> None:
    """Cierra popups y modales de Locatel si están visibles:
    - Modal 'Elige tu ubicación' (geolocalización)
    - Popup promocional (wpn-modal-img-container: Green Days, etc.)
    """
    # Modal de geolocalización "Elige tu ubicación"
    try:
        btn = page.query_selector("button.locatelcolombia-regionalizador-0-x-btnClose")
        if btn and btn.is_visible():
            btn.click()
            page.wait_for_timeout(800)
            write_log("Info", "HU02: Modal de geolocalización cerrado", task_name, in_config)
    except Exception:
        pass

    # Popup promocional (imágenes, ofertas, etc.)
    # Intenta botón "Saltar" por texto o selectores comunes; Escape como fallback.
    try:
        popup = page.query_selector("div.wpn-modal-img-container")
        if popup and popup.is_visible():
            cerrado = False
            for sel in [
                "button:has-text('Saltar')",
                ".wpn-modal-img-container button",
                "button.wpn-modal-close",
                "button.wpn-close",
                "button.wpn-btn-close",
            ]:
                try:
                    btn = page.query_selector(sel)
                    if btn and btn.is_visible():
                        btn.click()
                        page.wait_for_timeout(500)
                        cerrado = True
                        break
                except Exception:
                    pass
            if not cerrado:
                page.keyboard.press("Escape")
                page.wait_for_timeout(500)
            write_log("Info", "HU02: Popup promocional cerrado", task_name, in_config)
    except Exception:
        pass


# ============================================================
# Logica de scraping por EAN en Locatel
# ============================================================

def _consultar_ean(page: Page, ean: str, palabra_clave: str, url_template: str,
                   ruta_screenshot: str, in_config: dict, task_name: str,
                   primer_ean: bool = False) -> dict:
    """
    Navega a la URL de busqueda de Locatel para el EAN dado y extrae:
      titulo, precio_con_desc, precio_sin_desc, disponibilidad, marca, url_producto

    Retorna un dict con esas claves y 'estado' (2, 3 o 99).
    """
    resultado = {
        "titulo":           "",
        "precio_con_desc":  "0",
        "precio_sin_desc":  "",
        "marca":            "",
        "url_producto":     "",
        "disponibilidad":   "",
        "estado":           "99",
        "observaciones":    "No existe el producto en la farmacia",
    }

    url_consulta = url_template.replace("REEMPLAZAR", ean)

    try:
        # Capturar errores HTTP (4xx/5xx) por separado para no dejar la pagina
        # en estado chrome-error:// y arrastrar el fallo a los EANs siguientes.
        try:
            page.goto(url_consulta, wait_until="domcontentloaded", timeout=60000)
        except Exception as nav_err:
            write_log("Warning",
                      f"HU02: Error de navegacion EAN ({ean}): {str(nav_err)[:150]}",
                      task_name, in_config)
            page.wait_for_timeout(1500)
            return resultado

        # Primer EAN: esperar 12 s para que el banner inicial desaparezca (~10 s)
        page.wait_for_timeout(12000 if primer_ean else ESPERA_5S)
        _cerrar_modales_locatel(page, task_name, in_config)

        # ── URL y Titulo del primer producto ──────────────────────────
        # Locatel puede estar en página de búsqueda (cards) o de detalle
        # (redirect cuando hay un único resultado por EAN).
        # Se prueba primero el h1 de detalle y se usa productBrand como fallback.
        titulo = ""
        url_producto = ""
        for _ in range(5):
            # Página de detalle (redirect de EAN único)
            titulo = _js(
                page,
                "document.querySelector("
                "'h1.locatelcolombia-components-5-x-name-products')"
                "?.innerText.trim()"
            )
            if not titulo:
                # Página de resultados de búsqueda (cards)
                titulo = _js(
                    page,
                    "document.querySelector("
                    "'.vtex-product-summary-2-x-productBrand')"
                    "?.innerText.trim()"
                )
            if not url_producto:
                url_producto = _js(
                    page,
                    "document.querySelector("
                    "'.vtex-product-summary-2-x-clearLink--itemList')?.href"
                ) or page.url
            if titulo:
                break
            page.wait_for_timeout(ESPERA_3S)

        resultado["titulo"]      = titulo or ""
        resultado["url_producto"] = url_producto or page.url

        if not titulo:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — No existe el producto en la farmacia",
                task_name, in_config
            )
            _tomar_screenshot(page, ruta_screenshot)
            resultado["estado"]        = "99"
            resultado["observaciones"] = "No existe el producto en la farmacia"
            return resultado

        # ── Precios ───────────────────────────────────────────────────
        # El precio se parte en múltiples spans currencyInteger--summary
        # ("26" + "850" → "26850"). querySelector() acota al primer producto
        # de la página (evita mezclar spans de varios cards en búsqueda).
        precio_con_desc_raw = _js(
            page,
            "(()=>{"
            "  const c=document.querySelector("
            "    '.vtex-product-price-1-x-sellingPriceValue--summary');"
            "  return c ? [...c.querySelectorAll("
            "    '.vtex-product-price-1-x-currencyInteger--summary')]"
            "    .map(e=>e.innerText.trim()).join('') : '';"
            "})()"
        )
        precio_sin_desc_raw = _js(
            page,
            "(()=>{"
            "  const c=document.querySelector("
            "    '.vtex-product-price-1-x-listPriceValue--summary');"
            "  return c ? [...c.querySelectorAll("
            "    '.vtex-product-price-1-x-currencyInteger--summary')]"
            "    .map(e=>e.innerText.trim()).join('') : '';"
            "})()"
        )

        precio_con_desc = _extraer_precio_entero(str(precio_con_desc_raw or ""))
        precio_sin_desc = _extraer_precio_entero(str(precio_sin_desc_raw or ""))

        if not precio_sin_desc:
            precio_sin_desc = precio_con_desc
            precio_con_desc = "0"

        resultado["precio_con_desc"] = precio_con_desc
        resultado["precio_sin_desc"] = precio_sin_desc

        # ── Disponibilidad / Stock ─────────────────────────────────────
        # Si aparece el botón "sin stock" (buttonNoPdp) → sin stock.
        # Si solo existe "buttonPdp" (COMPRAR) → disponible.
        disponibilidad_raw = _js(
            page,
            "(document.querySelector("
            "'.locatelcolombia-delivery-modal-0-x-buttonNoPdp')"
            "?.innerText.trim()) || 'Texto no encontrado'"
        )
        resultado["disponibilidad"] = str(disponibilidad_raw or "Texto no encontrado")

        # ── Marca ─────────────────────────────────────────────────────
        # En paginas de detalle, brandName coincide con el h1 del nombre.
        # Se prueba primero el brand link de store-components (detalle)
        # y se excluyen elementos H1 para no confundir marca con nombre.
        marca_raw = _js(
            page,
            "(()=>{"
            "  const sels=['vtex-store-components-3-x-productBrandLink',"
            "               'vtex-store-components-3-x-productBrand a',"
            "               'vtex-product-summary-2-x-brandName'];"
            "  for(const c of sels){"
            "    const el=document.querySelector('.'+c);"
            "    if(el && el.tagName!=='H1') return el.innerText.trim();"
            "  }"
            "  return '';"
            "})()"
        )
        resultado["marca"] = str(marca_raw or "")

        # ── Determinar estado del registro ────────────────────────────
        titulo_upper = titulo.upper()
        kw_upper     = (palabra_clave or "").upper().strip()

        if kw_upper and kw_upper not in titulo_upper:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — Sin coincidencia: titulo='{titulo}', "
                f"palabra_clave='{palabra_clave}'",
                task_name, in_config
            )
            _tomar_screenshot(page, ruta_screenshot)
            resultado["estado"]        = "3"
            resultado["observaciones"] = (
                "No existe coincidencia entre la informacion encontrada "
                "y el producto consultado"
            )
            return resultado

        sin_stock = (
            "no encontrado" not in disponibilidad_raw.lower()
            and disponibilidad_raw.strip() != ""
        )

        if sin_stock:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — Encontrado pero sin stock: '{titulo}'",
                task_name, in_config
            )
            resultado["estado"]        = "2"
            resultado["observaciones"] = "Sin stock"
        else:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — Producto encontrado: '{titulo}' "
                f"precio={precio_sin_desc}",
                task_name, in_config
            )
            resultado["estado"]        = "2"
            resultado["observaciones"] = ""

        _tomar_screenshot(page, ruta_screenshot)

    except PlaywrightTimeout:
        write_log(
            "Warning",
            f"HU02: Timeout consultando EAN ({ean})",
            task_name, in_config
        )
        resultado["estado"]        = "99"
        resultado["observaciones"] = "Timeout al cargar la pagina"
    except Exception as e:
        write_log(
            "Warning",
            f"HU02: Error consultando EAN ({ean}): {str(e)[:200]}",
            task_name, in_config
        )
        resultado["estado"]        = "99"
        resultado["observaciones"] = f"Error: {str(e)[:200]}"

    return resultado


# ============================================================
# Funcion principal
# ============================================================

def hu02_consulta_y_reporte(in_config: dict) -> str:
    """
    Ejecuta la consulta web y generacion de reporte.

    Parametros:
        in_config: Diccionario de configuracion (ioConfig de HU00).

    Retorna:
        '' si exitoso, mensaje de error si fallo.
    """
    out_system_exception = ""
    task_name = "HU02_ConsultaYReporte"
    write_log("Info", "Inicia HU02", task_name, in_config)

    pw_instance = None
    browser     = None

    try:
        esquema      = in_config["Scheme"]
        tabla_loc    = in_config["TablaLocatel"]
        tabla_ins    = in_config["TablaTicketInsumo"]
        url_template = in_config.get("UrlLocatel") or ""
        debug        = in_config.get("_debug", False)
        lote         = int(in_config["LoteDebug"]) if debug else int(in_config["LoteLocatel"])
        reintentos_r = in_config["ReintentosReprocesamiento"]
        maquina      = socket.gethostname()

        # ----------------------------------------------------------------
        # PASO 1: Reprocesar registros "Sin stock" con reintentos disponibles
        # ----------------------------------------------------------------
        conn   = conectar_bd(in_config)
        cursor = conn.cursor()

        cursor.execute(f"""
            UPDATE {esquema}.{tabla_loc}
            SET Estado='1', FechaModificacion=GETDATE(), Reintentos=Reintentos+1
            WHERE (Estado='2' OR Estado='100')
              AND [Observaciones]='Sin stock'
              AND Reintentos<={reintentos_r}
        """)
        write_log("Info", f"HU02: Registros reactivados (sin stock): {cursor.rowcount}", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 2: Insertar en TablaLocatel los IDs nuevos de TicketInsumo
        # ----------------------------------------------------------------
        # TRUNCATE en TicketInsumo resetea el IDENTITY a 1, por lo que los
        # nuevos IDs colisionan con los de corridas anteriores en Locatel.
        # Eliminamos los registros viejos cuyos IDs coinciden con el lote
        # actual pero tienen una fecha anterior (ya reportados).
        cursor.execute(f"""
            DELETE b FROM {esquema}.{tabla_loc} b
            JOIN {esquema}.{tabla_ins} a ON a.Id = b.Id
            WHERE b.FechaInicio < a.FechaInicio
               OR b.Estado IN ('100', '199', '3')
        """)
        write_log("Info", f"HU02: Registros anteriores eliminados de {tabla_loc}: {cursor.rowcount}", task_name, in_config)

        cursor.execute(f"""
            SELECT COUNT(*) FROM {esquema}.{tabla_ins} a
            LEFT JOIN {esquema}.{tabla_loc} b ON a.Id = b.Id
            WHERE b.Id IS NULL AND a.Estado='1'
        """)
        cnt_nuevos = cursor.fetchone()[0]
        write_log("Info", f"HU02: Nuevos registros para insertar en {tabla_loc}: {cnt_nuevos}", task_name, in_config)

        if cnt_nuevos > 0:
            cursor.execute(f"SET IDENTITY_INSERT {esquema}.{tabla_loc} ON")
            cursor.execute(f"""
                INSERT INTO {esquema}.{tabla_loc}
                    ([Id],[FechaInicio],[FechaModificacion],[FechaFin],
                     [Estado],[Observaciones],[Reintentos],[Maquina],[FechaInsumo],
                     [PLU],[EAN],[Descripcion],[Categoria],
                     [HoraConsulta],[MarcaProducto],[NombrePrd],[RegistroInvima],
                     [PrecioUnitario],[PrecioConDescuento],[PrecioSinDescuento],
                     [Porc.Descuento],[PrecioFidelizacion],[BannerProducto],
                     [UrlProducto],[RutaImagen])
                SELECT
                    a.[Id], a.[FechaInicio], GETDATE(), NULL,
                    '1', '', '0', '{maquina}', a.[FechaInicio],
                    a.[PLU], a.[EAN], a.[Descripcion], a.[Categoria],
                    NULL,'','','','','','','','','','',''
                FROM {esquema}.{tabla_ins} a
                LEFT JOIN {esquema}.{tabla_loc} b ON a.Id = b.Id
                WHERE b.Id IS NULL AND a.Estado='1'
            """)
            cursor.execute(f"SET IDENTITY_INSERT {esquema}.{tabla_loc} OFF")
            write_log("Info", f"HU02: Insertados {cursor.rowcount} registros en {tabla_loc}", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 3: Verificar si hay registros pendientes
        # ----------------------------------------------------------------
        cursor.execute(f"""
            SELECT COUNT(*) FROM {esquema}.{tabla_loc}
            WHERE Estado='1' OR Estado='2' OR Estado='99'
        """)
        cnt_pendientes = cursor.fetchone()[0]
        hay_pendientes = cnt_pendientes > 0
        conn.commit()
        conn.close()

        write_log("Info", f"HU02: Registros pendientes en {tabla_loc}: {cnt_pendientes}", task_name, in_config)

        if not hay_pendientes:
            write_log("Info", "HU02: No existen registros para consultar en pagina", task_name, in_config)
            write_log("Info", "Finaliza HU02", task_name, in_config)
            return ""

        write_log("Info", "HU02: Existen registros que requieren consulta en la pagina", task_name, in_config)
        if not url_template or "REEMPLAZAR" not in url_template:
            write_log("Warning", f"HU02: URL template no contiene 'REEMPLAZAR': '{url_template}'", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 4: Estructura de carpetas de screenshots
        # ----------------------------------------------------------------
        debug      = in_config.get("_debug", False)
        now = datetime.now()
        if debug:
            ruta_screenshots = str(
                _PROJECT_ROOT / "debug" / "screenshots" / "Locatel"
                / str(now.year) / f"{now.month:02d}" / f"{now.day:02d}"
            )
        else:
            ruta_screenshots = os.path.join(
                in_config["RutaScreenshots"],
                in_config["CarpetaLocatel"],
                str(now.year),
                f"{now.month:02d}",
                f"{now.day:02d}"
            )
        os.makedirs(ruta_screenshots, exist_ok=True)

        # ----------------------------------------------------------------
        # PASO 5: Bucle principal de scraping con Playwright
        # ----------------------------------------------------------------
        headless   = False if debug else str(in_config["HeadlessChrome"]).lower() == "true"
        proxy_url  = _proxy_sistema_windows()
        proxy_cfg  = {"server": proxy_url} if proxy_url else None

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

        # Contexto y página únicos para toda la sesión de scraping.
        # Reutilizar la misma pestaña evita que el popup promocional de Locatel
        # se dispare en cada EAN (solo aparece en la primera carga) y reduce
        # las peticiones redundantes que pueden activar bloqueos anti-bot.
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            locale="es-CO",
            viewport={"width": 1920, "height": 1080},
            ignore_https_errors=True
        )
        page = context.new_page()
        primer_ean = True

        hay_mas = True
        while hay_mas:
            conn   = conectar_bd(in_config)
            cursor = conn.cursor()

            cursor.execute(f"""
                SELECT TOP({lote}) [Id], [EAN],
                    ISNULL(
                        NULLIF(
                            LEFT(
                                LTRIM(SUBSTRING(Descripcion,
                                    PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100)),
                                CASE
                                    WHEN CHARINDEX(' ',
                                        LTRIM(SUBSTRING(Descripcion,
                                            PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100))
                                        + ' ') > 1
                                    THEN CHARINDEX(' ',
                                        LTRIM(SUBSTRING(Descripcion,
                                            PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100))
                                        + ' ') - 1
                                    ELSE LEN(Descripcion)
                                END
                            ), ''
                        ), ''
                    ),
                    [PLU]
                FROM {esquema}.{tabla_loc}
                WHERE Estado='1'
            """)
            registros = cursor.fetchall()
            conn.close()

            write_log("Info", f"HU02: Lote scraping obtenido: {len(registros)} registros con Estado=1", task_name, in_config)

            if not registros:
                hay_mas = False
                break

            for i, row in enumerate(registros):
                id_ticket     = str(row[0])
                ean           = str(row[1])
                palabra_clave = str(row[2] or "")

                try:
                    conn   = conectar_bd(in_config)
                    cursor = conn.cursor()
                    cursor.execute(f"""
                        UPDATE {esquema}.{tabla_loc}
                        SET FechaModificacion=GETDATE()
                        WHERE Id='{id_ticket}'
                    """)
                    conn.commit()
                    conn.close()

                    ruta_ss = os.path.join(
                        ruta_screenshots,
                        f"{ean}_{id_ticket}.jpg"
                    )

                    write_log(
                        "Info",
                        f"HU02: [{i+1}/{len(registros)}] Consultando EAN ({ean}) "
                        f"— {url_template.replace('REEMPLAZAR', ean)}",
                        task_name, in_config
                    )

                    resultado = _consultar_ean(
                        page=page,
                        ean=ean,
                        palabra_clave=palabra_clave,
                        url_template=url_template,
                        ruta_screenshot=ruta_ss,
                        in_config=in_config,
                        task_name=task_name,
                        primer_ean=primer_ean,
                    )
                    primer_ean = False

                    conn   = conectar_bd(in_config)
                    cursor = conn.cursor()
                    estado        = resultado["estado"]
                    observaciones = resultado["observaciones"][:250].replace("'", "''")
                    titulo        = resultado["titulo"].replace(";", "").replace("'", "''")
                    marca         = resultado["marca"].replace("'", "''")
                    precio_sin    = resultado["precio_sin_desc"]
                    precio_con    = resultado["precio_con_desc"]
                    url_prd       = resultado["url_producto"].replace("'", "''")
                    ruta_img      = ruta_ss.replace("'", "''")

                    if estado == "99":
                        cursor.execute(f"""
                            UPDATE {esquema}.{tabla_loc}
                            SET [FechaFin]=GETDATE(),
                                [Estado]='99',
                                [Observaciones]='{observaciones}',
                                [RutaImagen]='{ruta_img}',
                                [UrlProducto]='{url_prd}'
                            WHERE Id='{id_ticket}'
                        """)
                    elif estado == "3":
                        cursor.execute(f"""
                            UPDATE {esquema}.{tabla_loc}
                            SET [FechaFin]=GETDATE(),
                                [Estado]='3',
                                [Observaciones]='{observaciones}',
                                [RutaImagen]='{ruta_img}',
                                [UrlProducto]='{url_prd}'
                            WHERE Id='{id_ticket}'
                        """)
                    else:
                        banner = "No disponible" if observaciones == "Sin stock" else ""
                        cursor.execute(f"""
                            UPDATE {esquema}.{tabla_loc}
                            SET [FechaFin]=GETDATE(),
                                [Estado]='2',
                                [Observaciones]='{observaciones}',
                                [BannerProducto]='{banner}',
                                [PrecioSinDescuento]=REPLACE(REPLACE(REPLACE('{precio_sin}',' ',''),'.',''),'$',''),
                                [PrecioConDescuento]=REPLACE(REPLACE(REPLACE('{precio_con}',' ',''),'.',''),'$',''),
                                [PrecioUnitario]='',
                                [UrlProducto]='{url_prd}',
                                [NombrePrd]=REPLACE('{titulo}',';',''),
                                [MarcaProducto]='{marca}',
                                [RutaImagen]='{ruta_img}'
                            WHERE Id='{id_ticket}'
                        """)

                    conn.commit()
                    conn.close()

                except Exception as ean_err:
                    write_log("Warning", f"HU02: Error procesando EAN ({ean}): {ean_err}", task_name, in_config)
                    try:
                        conn.close()
                    except Exception:
                        pass
                    # Recuperar la página si quedó en mal estado
                    try:
                        page.goto("about:blank", timeout=5000)
                    except Exception:
                        try:
                            page.close()
                        except Exception:
                            pass
                        try:
                            page = context.new_page()
                            primer_ean = True
                        except Exception:
                            pass

            conn   = conectar_bd(in_config)
            cursor = conn.cursor()
            cursor.execute(f"""
                SELECT COUNT(*) FROM {esquema}.{tabla_loc} WHERE Estado='1'
            """)
            cnt_restantes = cursor.fetchone()[0]
            hay_mas = cnt_restantes > 0
            conn.close()
            write_log("Info", f"HU02: Registros restantes con Estado=1: {cnt_restantes}", task_name, in_config)

        try:
            page.close()
        except Exception:
            pass
        try:
            context.close()
        except Exception:
            pass

        write_log("Info", "HU02: Termina consulta de productos por EAN", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 6: Generacion de reportes
        # ----------------------------------------------------------------
        _generar_reportes(in_config, esquema, tabla_loc, task_name)

        out_system_exception = ""
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
# Generacion de reportes Excel
# ============================================================

def _generar_reportes(in_config: dict, esquema: str,
                      tabla_loc: str, task_name: str) -> None:
    """
    Por cada FechaInicio con registros procesados (Estado 2 o 99):
      - Obtiene estadisticas.
      - Marca como reportados (2->100, 99->199).
      - Calcula Porc.Descuento.
      - Exporta Excel y envia correo.
    """
    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    cursor.execute(f"""
        SELECT DISTINCT(FechaInicio),
            CONVERT(varchar(8), FechaInicio, 112)
            + '_'
            + REPLACE(LEFT(CONVERT(varchar(8), FechaInicio, 108), 5), ':', '_')
        FROM {esquema}.{tabla_loc}
        WHERE Estado='2' OR Estado='99'
    """)
    fechas = cursor.fetchall()
    conn.close()

    for fecha_row in fechas:
        fecha_inicio = str(fecha_row[0])[:23]   # '2026-07-02 12:32:35.123' — 3 ms digits que acepta SQL Server
        fecha_sello  = str(fecha_row[1])

        _generar_reporte_fecha(
            in_config=in_config,
            esquema=esquema,
            tabla_loc=tabla_loc,
            fecha_inicio=fecha_inicio,
            fecha_sello=fecha_sello,
            task_name=task_name
        )


def _generar_reporte_fecha(in_config: dict, esquema: str, tabla_loc: str,
                           fecha_inicio: str, fecha_sello: str,
                           task_name: str) -> None:
    """Genera el reporte Excel para una FechaInicio especifica."""
    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='2' "
        f"WHERE [Estado]='100' AND FechaInicio='{fecha_inicio}'"
    )
    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='99' "
        f"WHERE [Estado]='199' AND FechaInicio='{fecha_inicio}'"
    )

    cursor.execute(f"""
        SELECT
            COUNT(*) AS TotalRegistros,
            SUM(CASE WHEN (Estado='2' OR Estado='100') THEN 1 ELSE 0 END) AS CantidadExtraidos,
            SUM(CASE WHEN ((Estado='2' OR Estado='100') AND Observaciones!='Sin stock') THEN 1 ELSE 0 END) AS CantidadEstado2,
            SUM(CASE WHEN ((Estado='2' OR Estado='100') AND Observaciones='Sin stock')  THEN 1 ELSE 0 END) AS CantidadSinStock,
            SUM(CASE WHEN (Estado='99' OR Estado='199') THEN 1 ELSE 0 END) AS CantidadEstado99
        FROM {esquema}.{tabla_loc}
        WHERE FechaInicio='{fecha_inicio}'
    """)
    stats = cursor.fetchone()
    total          = stats[0] if stats else 0
    extraidos      = stats[1] if stats else 0
    estado2_count  = stats[2] if stats else 0
    sin_stock      = stats[3] if stats else 0
    estado99_count = stats[4] if stats else 0

    write_log(
        "Info",
        f"HU02: Reporte FechaInicio={fecha_inicio} — "
        f"Total={total}, Extraidos={extraidos}, Estado2={estado2_count}, "
        f"SinStock={sin_stock}, Estado99={estado99_count}",
        task_name, in_config
    )

    # Calcular Porc.Descuento y limpiar PrecioConDescuento ANTES de cambiar
    # el estado a 100/199, porque las clausulas WHERE filtan por Estado='2'.
    cursor.execute(f"""
        UPDATE {esquema}.{tabla_loc}
        SET [PrecioConDescuento]='0'
        WHERE [Estado]='2'
          AND FechaInicio='{fecha_inicio}'
          AND TRY_CAST(PrecioSinDescuento AS INT) = TRY_CAST(PrecioConDescuento AS INT)
    """)

    cursor.execute(f"""
        UPDATE {esquema}.{tabla_loc}
        SET [Porc.Descuento] =
            ((TRY_CAST(PrecioSinDescuento AS INT) - TRY_CAST(PrecioConDescuento AS INT)) * 100)
            / TRY_CAST(PrecioSinDescuento AS INT)
        WHERE [Estado]='2'
          AND FechaInicio='{fecha_inicio}'
          AND TRY_CAST(PrecioSinDescuento AS INT) != TRY_CAST(PrecioConDescuento AS INT)
          AND TRY_CAST(PrecioConDescuento AS INT) > 0
    """)

    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='100' "
        f"WHERE [Estado]='2' AND FechaInicio='{fecha_inicio}'"
    )
    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='199' "
        f"WHERE [Estado]='99' AND FechaInicio='{fecha_inicio}'"
    )

    conn.commit()
    conn.close()

    # Conexion fresca para el SELECT final: evita que el estado residual
    # de los cursores de UPDATE corrompa cursor.description en pyodbc.
    conn   = conectar_bd(in_config)
    cursor = conn.cursor()
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
        FROM {esquema}.{tabla_loc}
        WHERE FechaInicio='{fecha_inicio}'
    """)
    cols  = [col[0] for col in cursor.description]
    filas = cursor.fetchall()
    conn.close()

    if not filas:
        return

    df = pd.DataFrame([list(row) for row in filas], columns=cols)

    nombre_resultado = in_config["NombreResultado"]
    nombre_hoja      = in_config["NombreHojaResultado"]

    if in_config.get("_debug"):
        ruta_reporte = str(_PROJECT_ROOT / "debug")
        nombre_excel = f"DEBUG_{nombre_resultado}{fecha_sello}.xlsx"
    else:
        ruta_reporte = in_config.get("RutaReporte") or ""
        nombre_excel = f"{nombre_resultado}{fecha_sello}.xlsx"

    os.makedirs(ruta_reporte, exist_ok=True)
    ruta_excel = os.path.join(ruta_reporte, nombre_excel)

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)

    write_log(
        "Info",
        f"HU02: Reporte generado en ({ruta_excel})",
        task_name, in_config
    )

    from_address = in_config.get("_correo", {}).get("usuario", "")
    reemplazo = {"$NombrePagina$": in_config["DrogueriaLocatel"]}
    err = enviar_correo(
        in_config=in_config,
        i_cod_email=100,
        i_from_address=from_address,
        i_replace_in_message=reemplazo,
        i_replace_in_subject=reemplazo,
        i_html_format=False,
        i_attachment=[ruta_excel]
    )
    if err:
        write_log("Info", f"HU02: No fue posible enviar el correo: {err}", task_name, in_config)


if __name__ == "__main__":
    from Funciones.utils import obtener_config
    config = obtener_config()
    exc = hu02_consulta_y_reporte(config)
    if exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    print("HU02 completada exitosamente.")
