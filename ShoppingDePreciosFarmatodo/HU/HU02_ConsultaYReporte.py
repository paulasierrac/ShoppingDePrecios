"""
================================================================================
HU02 - Consulta y Reporte
Nombre de la iniciativa: Shopping de Precios Farmatodo
Autor: Paula Sierra — Net Applications
Descripcion: Consulta precios en farmatodo.com.co por EAN.
             Los selectores CSS se cargan desde [ShoppingDePrecios].[Selectores]
             en SQL Server (Competencia='FARMATODO') en lugar de estar
             hardcodeados en el codigo.
Ultima modificacion: 01/07/2026
Propiedad de Colsubsidio
================================================================================

Flujo principal:
  1. Carga selectores CSS desde la tabla [Selectores] en BD.
  2. Inserta en [Farmatodo] los IDs nuevos desde [TicketInsumo].
  3. Verifica registros pendientes (Estado='1').
  4. Crea carpeta de screenshots.
  5. Bucle de scraping por lotes (CantFarmatodo).
     Para cada EAN:
       a. Navega a URL de busqueda y espera carga.
       b. Extrae campos via JS usando selectores de BD.
       c. Reintenta si precio no encontrado (hasta 3 veces).
       d. Actualiza BD.
     Espera SegFarmatodo segundos entre lotes.
  6. Generacion de reporte Excel y envio de correo.

Estados:
  1  : Pendiente
  2  : Producto encontrado
  3  : Sin coincidencia
  99 : Sin informacion / no encontrado
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
from Funciones.utils import write_log, conectar_bd, conectar_bd_debug, enviar_correo


ESPERA_CARGA  = 5000    # ms
ESPERA_REINT  = 3000    # ms entre reintentos
_LOTE_DEFAULT = 50


# ============================================================
# Helpers
# ============================================================

def _proxy_sistema_windows() -> dict:
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


def _js_selector(page: Page, selector: str, default: str = "") -> str:
    """Obtiene textContent del primer elemento que coincida con el selector CSS."""
    try:
        result = page.evaluate(
            f"document.querySelector({repr(selector)})?.textContent?.trim() || ''"
        )
        return (result or "").strip()
    except Exception:
        return default


def _limpiar_precio(texto: str) -> str:
    """Elimina simbolo de moneda y separadores de miles."""
    if not texto:
        return ""
    texto = texto.replace("$", "").replace("\xa0", "").strip()
    texto = texto.replace(".", "").replace(",", "")
    return re.sub(r"[^\d]", "", texto)


def _tomar_screenshot(page: Page, ruta: str) -> None:
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        page.screenshot(path=ruta)
    except Exception:
        pass


# ============================================================
# Carga de selectores CSS desde BD
# ============================================================

def _cargar_selectores(in_config: dict, task_name: str) -> dict:
    """Lee los selectores CSS desde [ShoppingDePrecios].[Selectores] WHERE Competencia='FARMATODO'."""
    selectores = {}
    try:
        conn   = conectar_bd(in_config)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT [Clave], [Selector] FROM [ShoppingDePrecios].[Selectores] "
            "WHERE Competencia='FARMATODO'"
        )
        for row in cursor.fetchall():
            selectores[str(row[0])] = str(row[1])
        conn.close()
        write_log("Info", f"HU02: {len(selectores)} selectores cargados desde BD", task_name, in_config)
    except Exception as e:
        write_log("Warning", f"HU02: Error cargando selectores: {e}", task_name, in_config)
    return selectores


# ============================================================
# Logica de scraping por EAN en Farmatodo
# ============================================================

def _consultar_ean_farmatodo(page: Page, ean: str, url_template: str,
                              selectores: dict, ruta_screenshot: str,
                              in_config: dict, task_name: str) -> dict:
    resultado = {
        "nombre_prd":      "",
        "marca":           "",
        "precio_con_desc": "",
        "precio_sin_desc": "",
        "registro_invima": "",
        "url_producto":    "",
        "banner":          "",
        "estado":          "99",
    }

    url_busqueda = url_template.replace("REEMPLAZAR", ean)

    try:
        page.goto(url_busqueda, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(ESPERA_CARGA)

        nombre_prd = ""
        precio_con = ""
        precio_sin = ""
        marca      = ""
        reg_invima = ""
        url_prod   = ""
        banner     = ""

        for intento in range(3):
            nombre_prd = _js_selector(page, selectores.get("NombrePrd", ".product-name"))
            precio_con = _js_selector(page, selectores.get("PrecioConDescuento", ".price-box .price"))
            precio_sin = _js_selector(page, selectores.get("PrecioSinDescuento", ".price-box .old-price .price"))
            marca      = _js_selector(page, selectores.get("Marca", ".product-brand"))
            reg_invima = _js_selector(page, selectores.get("RegistroInvima", ".invima"))
            banner     = _js_selector(page, selectores.get("Banner", ".badge-label"))
            try:
                url_prod = page.evaluate(
                    f"document.querySelector({repr(selectores.get('UrlProducto', 'link[rel=canonical]'))})?.href || ''"
                ) or url_busqueda
            except Exception:
                url_prod = url_busqueda

            if nombre_prd or precio_con:
                break
            if intento < 2:
                write_log("Info", f"HU02: EAN ({ean}) — Reintento {intento+1} sin datos, recargando pagina",
                          task_name, in_config)
                try:
                    page.reload(wait_until="domcontentloaded", timeout=60000)
                except Exception:
                    pass
                page.wait_for_timeout(ESPERA_REINT)

        _tomar_screenshot(page, ruta_screenshot)

        if not nombre_prd and not precio_con:
            write_log("Info", f"HU02: EAN ({ean}) — Sin informacion en Farmatodo", task_name, in_config)
            resultado["url_producto"] = url_prod
            return resultado

        write_log("Info", f"HU02: EAN ({ean}) — Producto encontrado: '{nombre_prd}'",
                  task_name, in_config)

        resultado.update({
            "nombre_prd":      nombre_prd,
            "marca":           marca,
            "precio_con_desc": _limpiar_precio(precio_con),
            "precio_sin_desc": _limpiar_precio(precio_sin),
            "registro_invima": reg_invima,
            "url_producto":    url_prod,
            "banner":          banner,
            "estado":          "2",
        })

    except PlaywrightTimeout:
        write_log("Warning", f"HU02: Timeout consultando EAN ({ean})", task_name, in_config)
        resultado["estado"] = "99"
    except Exception as e:
        write_log("Warning", f"HU02: Error consultando EAN ({ean}): {e}", task_name, in_config)
        resultado["estado"] = "99"

    return resultado


# ============================================================
# Funcion principal
# ============================================================

def hu02_consulta_y_reporte(in_config: dict) -> str:
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
        tabla_ex     = in_config.get("TablaFarmatodo",     "[Farmatodo]")
        tabla_ins    = in_config.get("TablaTicketInsumo",  "[TicketInsumo]")
        url_template = in_config.get("UrlFarmatodo", "")
        maquina      = socket.gethostname()
        lote         = int(in_config.get("CantFarmatodo", str(_LOTE_DEFAULT)))
        delay        = int(in_config.get("SegFarmatodo",  "300"))

        # ── Carga de selectores ───────────────────────────────────────────
        selectores = _cargar_selectores(in_config, task_name)

        # ── PASO 1 + PASO 2 ──────────────────────────────────────────────
        if debug:
            conn_sq = conectar_bd_debug(in_config)
            cur_sq  = conn_sq.cursor()
            cur_sq.execute(f"SELECT COUNT(*) FROM {esquema}.{tabla_ins} WHERE Estado=1")
            cnt = cur_sq.fetchone()[0] or 0
            conn_sq.close()
            hay_pendientes = cnt > 0
            if hay_pendientes:
                write_log("Info", f"[DEBUG] {cnt} registros en BD Dev ({tabla_ins}) con Estado=1",
                          task_name, in_config)
            else:
                write_log("Info", f"[DEBUG] No hay registros en BD Dev ({tabla_ins}) con Estado=1",
                          task_name, in_config)
        else:
            conn   = conectar_bd(in_config)
            cursor = conn.cursor()
            cursor.execute(f"""
                SELECT a.Id FROM {esquema}.{tabla_ins} a
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
                         [PrecioConDescuento],[PrecioSinDescuento],[Porc.Descuento],
                         [PrecioFidelizacion],[UrlProducto],[BannerProducto],[RutaImagen],[HoraConsulta])
                    SELECT a.[Id], a.[FechaInicio], GETDATE(), '',
                           '1', '{maquina}', a.[PLU], a.[EAN], a.[Descripcion],
                           '','','','','','','','','','',GETDATE()
                    FROM {esquema}.{tabla_ins} a
                    LEFT JOIN {esquema}.{tabla_ex} b ON a.Id = b.Id
                    WHERE b.Id IS NULL AND a.Estado='1'
                """)
                cursor.execute(f"SET IDENTITY_INSERT {esquema}.{tabla_ex} OFF")
                write_log("Info", f"HU02: Nuevos registros insertados en {tabla_ex}", task_name, in_config)
            cursor.execute(f"SELECT TOP(1) 1 FROM {esquema}.{tabla_ex} WHERE Estado='1'")
            hay_pendientes = cursor.fetchone() is not None
            conn.commit()
            conn.close()

        if not hay_pendientes:
            write_log("Info", "HU02: No existen registros pendientes", task_name, in_config)
            write_log("Info", "Finaliza HU02", task_name, in_config)
            return ""

        write_log("Info", "HU02: Existen registros pendientes para consultar", task_name, in_config)

        # ── PASO 3: Carpeta de screenshots ────────────────────────────────
        now = datetime.now()
        if debug:
            ruta_ss_base = str(_PROJECT_ROOT / "debug" / "screenshots" / "Farmatodo"
                               / str(now.year) / f"{now.month:02d}" / f"{now.day:02d}")
        else:
            ruta_ss_base = os.path.join(
                in_config.get("RutaScreenshots", ""),
                in_config.get("CarpetaFarmatodo", "Farmatodo\\"),
                str(now.year), f"{now.month:02d}", f"{now.day:02d}",
            )
        os.makedirs(ruta_ss_base, exist_ok=True)

        # ── PASO 4: Bucle de scraping ─────────────────────────────────────
        headless  = False if debug else str(in_config.get("HeadlessChrome", "true")).lower() == "true"
        proxy_cfg = _proxy_sistema_windows()

        write_log("Info", "HU02: Inicia consulta de productos por EAN", task_name, in_config)

        _asegurar_chromium(in_config, task_name)
        pw_instance = sync_playwright().start()
        browser = pw_instance.chromium.launch(
            headless=headless,
            proxy=proxy_cfg,
            args=["--lang=es-CO", "--no-sandbox", "--disable-dev-shm-usage",
                  "--disable-gpu", "--disable-software-rasterizer"]
        )

        if debug:
            _scraping_debug(browser, in_config, esquema, tabla_ins, url_template,
                            selectores, ruta_ss_base, task_name)
        else:
            _scraping_normal(browser, in_config, esquema, tabla_ex, url_template,
                             selectores, ruta_ss_base, maquina, lote, delay, task_name)

        write_log("Info", "HU02: Termina consulta de productos por EAN", task_name, in_config)

        # ── PASO 5: Reporte ───────────────────────────────────────────────
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
# Scraping modo normal
# ============================================================

def _scraping_normal(browser, in_config, esquema, tabla_ex, url_template,
                     selectores, ruta_ss_base, maquina, lote, delay, task_name):
    hay_mas = True
    while hay_mas:
        conn   = conectar_bd(in_config)
        cursor = conn.cursor()
        cursor.execute(f"""
            SELECT TOP({lote}) [Id], [EAN], [Descripcion]
            FROM {esquema}.{tabla_ex} WHERE Estado='1'
        """)
        registros = cursor.fetchall()
        conn.close()

        if not registros:
            break

        context = browser.new_context(
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
            locale="es-CO",
            viewport={"width": 1920, "height": 1080},
            ignore_https_errors=True,
        )
        page = context.new_page()

        try:
            for row in registros:
                id_t, ean = str(row[0]), str(row[1])
                ruta_ss   = os.path.join(ruta_ss_base, f"{ean}_{id_t}.jpg")

                conn   = conectar_bd(in_config)
                cursor = conn.cursor()
                cursor.execute(
                    f"UPDATE {esquema}.{tabla_ex} "
                    f"SET FechaModificacion=GETDATE(), HoraConsulta=GETDATE() WHERE Id='{id_t}'"
                )
                conn.commit()
                conn.close()

                write_log("Info", f"HU02: Consultando EAN ({ean})", task_name, in_config)
                res = _consultar_ean_farmatodo(page, ean, url_template, selectores,
                                               ruta_ss, in_config, task_name)
                _persistir(in_config, esquema, tabla_ex, id_t, ruta_ss, res, task_name)
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
# Scraping modo debug
# ============================================================

def _scraping_debug(browser, in_config, esquema, tabla_ins, url_template,
                    selectores, ruta_ss_base, task_name):
    lote_debug = int(in_config.get("LoteDebug", "3"))
    conn_sq = conectar_bd_debug(in_config)
    cur_sq  = conn_sq.cursor()
    cur_sq.execute(
        f"SELECT TOP (?) Id, EAN, Descripcion FROM {esquema}.TicketInsumo WHERE Estado=1",
        (lote_debug,)
    )
    registros = cur_sq.fetchall()

    if not registros:
        write_log("Info", "[DEBUG] No hay registros en TicketInsumo con Estado=1", task_name, in_config)
        conn_sq.close()
        return

    write_log("Info", f"[DEBUG] Procesando {len(registros)} EAN(s) (LoteDebug={lote_debug})",
              task_name, in_config)

    resultados = []
    context = browser.new_context(
        user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
        locale="es-CO",
        viewport={"width": 1920, "height": 1080},
        ignore_https_errors=True,
    )
    page = context.new_page()

    try:
        for row in registros:
            id_t = str(row[0])
            ean  = str(row[1])
            desc = str(row[2] or "")
            ruta_ss = os.path.join(ruta_ss_base, f"{ean}_{id_t}.jpg")
            print(f"\n  EAN: {ean}  |  {desc[:50]}")
            res = _consultar_ean_farmatodo(page, ean, url_template, selectores,
                                           ruta_ss, in_config, task_name)
            print(f"  Estado: {res['estado']} | Nombre: {res['nombre_prd']} | Precio: {res['precio_con_desc']}")
            resultados.append({"Id": id_t, "EAN": ean, "Descripcion": desc, "RutaImagen": ruta_ss, **res})
    finally:
        try:
            context.close()
        except Exception:
            pass

    if resultados:
        ahora   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        maquina = socket.gethostname()
        cur_sq.execute(f"DELETE FROM {esquema}.Farmatodo")
        for r in resultados:
            cur_sq.execute(
                f"INSERT INTO {esquema}.Farmatodo "
                "(FechaInicio, FechaModificacion, FechaFin, Estado, Maquina, "
                " PLU, EAN, Descripcion, HoraConsulta, MarcaProducto, NombrePrd, RegistroInvima, "
                " PrecioConDescuento, PrecioSinDescuento, [Porc.Descuento], "
                " PrecioFidelizacion, UrlProducto, BannerProducto, RutaImagen) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (ahora, ahora, ahora,
                 r.get("estado", "99"), maquina,
                 "", r["EAN"], r["Descripcion"], ahora,
                 r.get("marca", ""), r.get("nombre_prd", ""), r.get("registro_invima", ""),
                 r.get("precio_con_desc", ""), r.get("precio_sin_desc", ""), "",
                 "", r.get("url_producto", ""), r.get("banner", ""), r.get("RutaImagen", ""))
            )
        conn_sq.commit()
        write_log("Info", f"[DEBUG] {len(resultados)} registros guardados en ({esquema}.Farmatodo)",
                  task_name, in_config)

        ruta_debug = _PROJECT_ROOT / "debug"
        ruta_debug.mkdir(exist_ok=True)
        sello      = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_excel = str(ruta_debug / f"DEBUG_ReportePricingFarmatodo_{sello}.xlsx")
        pd.DataFrame(resultados).to_excel(ruta_excel, index=False)
        write_log("Info", f"[DEBUG] Reporte en ({ruta_excel})", task_name, in_config)
        print(f"\n  Reporte debug: {ruta_excel}")
    conn_sq.close()


# ============================================================
# Persistencia BD (modo normal)
# ============================================================

def _persistir(in_config, esquema, tabla_ex, id_t, ruta_ss, res, task_name):
    estado     = res["estado"]
    nombre_prd = res["nombre_prd"].replace(";", "").replace("'", "''")
    marca      = res["marca"].replace("'", "''")
    precio_con = res["precio_con_desc"]
    precio_sin = res["precio_sin_desc"]
    reg_inv    = res["registro_invima"].replace("'", "''")
    url_prd    = res["url_producto"].replace("'", "''")
    banner     = res["banner"].replace("'", "''")
    ruta_img   = ruta_ss.replace("'", "''")

    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    if estado == "99":
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='99',
                [UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)
    else:
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='2',
                [NombrePrd]='{nombre_prd}',[MarcaProducto]='{marca}',
                [RegistroInvima]='{reg_inv}',
                [PrecioConDescuento]='{precio_con}',[PrecioSinDescuento]='{precio_sin}',
                [BannerProducto]='{banner}',
                [UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)

    conn.commit()
    conn.close()


# ============================================================
# Generacion de reportes Excel (modo normal)
# ============================================================

def _generar_reportes(in_config, esquema, tabla_ex, task_name):
    conn   = conectar_bd(in_config)
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT DISTINCT FechaInicio,
            REPLACE(CONCAT(
                REPLACE(CAST(FechaInicio AS DATE),'-','_'),
                REPLACE(REPLACE(SUBSTRING(CAST(FechaInicio AS varchar),12,6),' ','_'),':','_')
            ),'__','_0')
        FROM {esquema}.{tabla_ex} WHERE Estado='2' OR Estado='99'
    """)
    fechas = cursor.fetchall()
    conn.close()
    for row in fechas:
        _generar_reporte_fecha(in_config, esquema, tabla_ex, str(row[0]), str(row[1]), task_name)


def _generar_reporte_fecha(in_config, esquema, tabla_ex, fecha_inicio, fecha_sello, task_name):
    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    cursor.execute(f"UPDATE {esquema}.{tabla_ex} SET [Estado]='2'  WHERE [Estado]='100' AND FechaInicio='{fecha_inicio}'")
    cursor.execute(f"UPDATE {esquema}.{tabla_ex} SET [Estado]='99' WHERE [Estado]='199' AND FechaInicio='{fecha_inicio}'")

    cursor.execute(f"""
        SELECT COUNT(*), SUM(CASE WHEN Estado IN ('2','100') THEN 1 ELSE 0 END),
               SUM(CASE WHEN Estado IN ('99','199') THEN 1 ELSE 0 END)
        FROM {esquema}.{tabla_ex} WHERE FechaInicio='{fecha_inicio}'
    """)
    stats = cursor.fetchone() or (0, 0, 0)
    write_log("Info", f"HU02: {fecha_inicio} — Total={stats[0]} Extraidos={stats[1]} Estado99={stats[2]}",
              task_name, in_config)

    cursor.execute(f"UPDATE {esquema}.{tabla_ex} SET [Estado]='100' WHERE [Estado]='2'  AND FechaInicio='{fecha_inicio}'")
    cursor.execute(f"UPDATE {esquema}.{tabla_ex} SET [Estado]='199' WHERE [Estado]='99' AND FechaInicio='{fecha_inicio}'")
    cursor.execute(f"""
        UPDATE {esquema}.{tabla_ex}
        SET [Porc.Descuento] =
            ((TRY_CAST(PrecioSinDescuento AS FLOAT) - TRY_CAST(PrecioConDescuento AS FLOAT)) * 100)
            / TRY_CAST(PrecioSinDescuento AS FLOAT)
        WHERE Estado='100' AND FechaInicio='{fecha_inicio}'
          AND TRY_CAST(PrecioSinDescuento AS FLOAT) > 0
          AND TRY_CAST(PrecioConDescuento AS FLOAT) > 0
          AND TRY_CAST(PrecioSinDescuento AS FLOAT) != TRY_CAST(PrecioConDescuento AS FLOAT)
    """)
    conn.commit()

    cursor.execute(f"""
        SELECT [FechaInicio],[PLU],[Descripcion],[HoraConsulta],[EAN],[Estado],
               [MarcaProducto],[NombrePrd],[RegistroInvima],
               [PrecioConDescuento],[PrecioSinDescuento],[Porc.Descuento],
               [PrecioFidelizacion],[BannerProducto],[UrlProducto],[RutaImagen]
        FROM {esquema}.{tabla_ex} WHERE FechaInicio='{fecha_inicio}'
    """)
    cols  = [c[0] for c in cursor.description]
    filas = cursor.fetchall()
    conn.close()

    if not filas:
        return

    df = pd.DataFrame(filas, columns=cols)
    ruta_rep    = in_config.get("RutaReporte", "")
    os.makedirs(ruta_rep, exist_ok=True)
    nombre_res  = in_config.get("NombreResultado", "ReportePricing")
    nombre_hoja = in_config.get("NombreHojaResultado", "ReportePricingFarmatodo")
    ruta_excel  = os.path.join(ruta_rep, f"{nombre_res}{fecha_sello}.xlsx")

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)

    write_log("Info", f"HU02: Reporte generado en ({ruta_excel})", task_name, in_config)
    from_addr = in_config.get("_correo", {}).get("usuario", "")
    reemplazo = {"$NombrePagina$": in_config.get("DrogueriaFarmatodo", "Farmatodo")}
    err = enviar_correo(in_config=in_config, i_cod_email=100, i_from_address=from_addr,
                        i_replace_in_message=reemplazo, i_replace_in_subject=reemplazo,
                        i_html_format=False, i_attachment=[ruta_excel])
    if err:
        write_log("Info", f"HU02: No fue posible enviar correo: {err}", task_name, in_config)


if __name__ == "__main__":
    from Funciones.utils import obtener_config
    config = obtener_config()
    exc = hu02_consulta_y_reporte(config)
    if exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    print("HU02 completada exitosamente.")
