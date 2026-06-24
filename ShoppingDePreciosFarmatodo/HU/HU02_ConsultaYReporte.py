"""
================================================================================
HU02 - Consulta y Reporte
Nombre de la iniciativa: Shopping de Precios Farmatodo
Autor: KPMG Advisory, Tax & Legal SAS
Descripcion: Consulta precios en farmatodo.com.co por EAN.
             Los selectores CSS de cada campo se cargan desde la tabla
             [ShoppingDePrecios].[Selectores] (Competencia='FARMATODO'),
             igual que el bot AA original. El bot navega a la pagina de detalle
             del primer resultado que contenga el EAN en su URL y extrae los
             campos via document.querySelector.
Ultima modificacion: 22/06/2026
Propiedad de Colsubsidio
================================================================================

Flujo principal:
  1. Carga selectores CSS desde [Selectores] WHERE Competencia='FARMATODO'.
  2. Inserta en [Farmatodo] los IDs nuevos desde [TicketInsumo].
  3. Verifica registros pendientes (Estado='1').
  4. Crea carpeta de screenshots.
  5. Bucle de scraping por lotes (CantFarmatodo, default 100).
     Para cada EAN:
       a. Navega a URL de busqueda (base + buscar?product=EAN&).
       b. Obtiene URL de detalle via JS: content-product item(0) y item(1).
       c. Valida EAN en URL; si no coincide prueba posicion 1.
       d. Navega a detalle, extrae campos usando los selectores de BD via JS.
       e. Valida nombre vs palabra clave.
       f. Actualiza BD.
     Espera SegFarmatodo segundos entre lotes.
  6. Limpieza de puntos de miles en precios.
  7. Generacion de reporte Excel y envio de correo.

Estados:
  1  : Pendiente
  2  : Producto encontrado
  3  : URL encontrada pero nombre no coincide
  99 : Sin informacion / no encontrado

Nota sobre selectores:
  Si la tabla [Selectores] no existe o no tiene filas para FARMATODO,
  se utilizan los defaults definidos en _SELECTORES_DEFAULT.
"""

import os
import re
import sys
import time
import socket
from datetime import datetime
from pathlib import Path

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
from Funciones.utils import write_log, conectar_bd, conectar_bd_debug, enviar_correo


ESPERA_CARGA  = 5    # segundos tras navegacion
ESPERA_REINT  = 3    # entre reintentos
_LOTE_DEFAULT = 100

# Selectores por defecto si la tabla Selectores no esta disponible
_SELECTORES_DEFAULT = {
    "NombreProducto":     ".product-name h1, h1.product-detail-info__header-name",
    "Marca":              ".product-manufacturer .manufacturer-name, [class*='brand']",
    "RegistroInvima":     ".product-reference span, [class*='invima']",
    "PrecioDescuento":    ".current-price span[itemprop='price'], .price",
    "PrecioNormal":       ".regular-price, .product-price-and-shipping .regular-price",
    "PrecioUnitario":     "[class*='price-per-unit'], [class*='unit-price']",
    "PorcentajeDescuento":"[class*='discount-percentage'], [class*='discount'] .discount-amount",
    "PrecioFidelizacion": "[class*='fidelizacion'], [class*='loyalty']",
    "BannerComentario":   "[class*='banner'], [class*='label']",
    "Divisor":            ".product-miniature, .js-product",
}
_COMPETENCIA = "FARMATODO"


def _crear_driver(headless: bool = True) -> webdriver.Chrome:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--lang=es-CO")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    return webdriver.Chrome(options=opts)


def _js_selector(driver, selector: str, default: str = "") -> str:
    """Extrae textContent del primer elemento que coincida con el selector CSS via JS."""
    if not selector:
        return default
    try:
        result = driver.execute_script(
            f"return document.querySelector({repr(selector)})?.textContent?.trim() || ''"
        )
        return (result or "").strip()
    except Exception:
        return default


def _limpiar_precio(texto: str) -> str:
    """Elimina simbolos de moneda y puntos de miles."""
    if not texto:
        return ""
    return re.sub(r"[^\d]", "", texto)


def _tomar_screenshot(driver, ruta: str) -> None:
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        driver.save_screenshot(ruta)
    except Exception:
        pass


def _cargar_selectores(in_config: dict, task_name: str) -> dict:
    """Carga selectores CSS desde [Selectores] WHERE Competencia='FARMATODO'."""
    try:
        esquema = in_config.get("Scheme", "[ShoppingDePrecios]")
        conn   = conectar_bd(in_config)
        cursor = conn.cursor()
        cursor.execute(f"""
            SELECT [Uso], [Selector]
            FROM {esquema}.[Selectores]
            WHERE Competencia = '{_COMPETENCIA}'
        """)
        filas = cursor.fetchall()
        conn.close()
        if filas:
            selectores = {str(r[0]): str(r[1]) for r in filas if r[0] and r[1]}
            write_log("Info", f"HU02: {len(selectores)} selectores cargados desde BD para {_COMPETENCIA}",
                      task_name, in_config)
            return selectores
    except Exception as e:
        write_log("Warning", f"HU02: No se pudo leer tabla Selectores ({e}), usando defaults",
                  task_name, in_config)
    return dict(_SELECTORES_DEFAULT)


def _consultar_ean_farmatodo(driver, ean: str, palabra_clave: str,
                              url_template: str, selectores: dict,
                              ruta_screenshot: str, in_config: dict,
                              task_name: str) -> dict:
    resultado = {
        "nombre_prd":          "",
        "marca":               "",
        "registro_invima":     "",
        "precio_con_desc":     "",
        "precio_sin_desc":     "",
        "precio_unitario":     "",
        "precio_fidelizacion": "",
        "porc_descuento":      "",
        "url_producto":        "",
        "banner":              "",
        "estado":              "99",
        "observaciones":       "No existe el producto en la farmacia",
    }

    url_busqueda = url_template.replace("REEMPLAZAR", ean)

    try:
        driver.get(url_busqueda)
        time.sleep(ESPERA_CARGA)

        # ── Paso 1: Obtener URL de producto desde resultados ────────────
        url_producto = ""
        for pos in range(2):
            for intento in range(3):
                try:
                    url_pos = driver.execute_script(
                        f"return document.getElementsByClassName('content-product').item({pos})?.href || ''"
                    ) or ""
                    if not url_pos:
                        break
                    if ean in url_pos:
                        url_producto = url_pos
                        break
                    if pos == 1:
                        # Si posicion 1 tampoco tiene EAN, podria ser producto unico
                        url_producto = url_pos
                    break
                except Exception as e:
                    err = str(e)
                    if "Cannot read properties of null" in err:
                        break
                    if intento < 2:
                        time.sleep(ESPERA_REINT)
                    else:
                        write_log("Warning", f"HU02: JS error pos {pos}: {e}", task_name, in_config)
            if url_producto and ean in url_producto:
                break

        if not url_producto:
            write_log("Info", f"HU02: EAN ({ean}) — No encontrado en resultados",
                      task_name, in_config)
            _tomar_screenshot(driver, ruta_screenshot)
            return resultado

        resultado["url_producto"] = url_producto

        # ── Paso 2: Navegar a detalle ──────────────────────────────────
        driver.get(url_producto)
        time.sleep(ESPERA_CARGA)

        # ── Paso 3: Extraer campos usando selectores de BD (3 reintentos) ─
        nombre_prd = marca = invima = precio_con = precio_sin = ""
        precio_unit = porc_desc = precio_fid = banner = ""

        for intento in range(3):
            nombre_prd  = _js_selector(driver, selectores.get("NombreProducto", ""))
            marca       = _js_selector(driver, selectores.get("Marca", ""))
            invima      = _js_selector(driver, selectores.get("RegistroInvima", ""))
            precio_con  = _js_selector(driver, selectores.get("PrecioDescuento", ""))
            precio_sin  = _js_selector(driver, selectores.get("PrecioNormal", ""))
            precio_unit = _js_selector(driver, selectores.get("PrecioUnitario", ""))
            porc_desc   = _js_selector(driver, selectores.get("PorcentajeDescuento", ""))
            precio_fid  = _js_selector(driver, selectores.get("PrecioFidelizacion", ""))
            banner      = _js_selector(driver, selectores.get("BannerComentario", ""))
            if nombre_prd or precio_con:
                break
            if intento < 2:
                try:
                    driver.find_element("tag name", "body").send_keys(Keys.F5)
                except Exception:
                    pass
                time.sleep(ESPERA_REINT)

        if not nombre_prd:
            write_log("Info", f"HU02: EAN ({ean}) — No se extrajo nombre del producto",
                      task_name, in_config)
            _tomar_screenshot(driver, ruta_screenshot)
            return resultado

        # ── Paso 4: Validar nombre vs palabra clave ────────────────────
        kw = (palabra_clave or "").upper().strip()
        if kw and kw not in nombre_prd.upper():
            write_log("Info",
                      f"HU02: EAN ({ean}) — Sin coincidencia: nombre='{nombre_prd}', kw='{kw}'",
                      task_name, in_config)
            _tomar_screenshot(driver, ruta_screenshot)
            resultado.update({
                "nombre_prd":   nombre_prd,
                "marca":        marca,
                "url_producto": url_producto,
                "estado":       "3",
                "observaciones": "No existe coincidencia entre la informacion encontrada y el producto consultado",
            })
            return resultado

        write_log("Info", f"HU02: EAN ({ean}) — Producto encontrado: '{nombre_prd}'",
                  task_name, in_config)
        _tomar_screenshot(driver, ruta_screenshot)

        resultado.update({
            "nombre_prd":          nombre_prd,
            "marca":               marca,
            "registro_invima":     invima,
            "precio_con_desc":     _limpiar_precio(precio_con),
            "precio_sin_desc":     _limpiar_precio(precio_sin),
            "precio_unitario":     precio_unit,
            "precio_fidelizacion": _limpiar_precio(precio_fid),
            "porc_descuento":      porc_desc.replace("%", "").strip(),
            "url_producto":        url_producto,
            "banner":              banner,
            "estado":              "2",
            "observaciones":       "",
        })

    except Exception as e:
        write_log("Warning", f"HU02: Error consultando EAN ({ean}): {e}", task_name, in_config)
        resultado["estado"]        = "99"
        resultado["observaciones"] = f"Error: {e}"

    return resultado


def hu02_consulta_y_reporte(in_config: dict) -> str:
    out_system_exception = ""
    task_name = "HU02_ConsultaYReporte"
    debug     = in_config.get("_debug", False)

    write_log("Info", "Inicia HU02", task_name, in_config)
    if debug:
        write_log("Info", "[DEBUG] Modo debug activo: sin escrituras en BD ni correos", task_name, in_config)

    driver = None
    try:
        esquema      = in_config.get("Scheme", "[ShoppingDePrecios]")
        tabla_ex     = in_config.get("TablaFarmatodo",     "[Farmatodo]")
        tabla_ins    = in_config.get("TablaTicketInsumo",  "[TicketInsumo]")
        url_template = in_config.get("UrlFarmatodo", "")
        maquina      = socket.gethostname()
        lote         = int(in_config.get("CantFarmatodo", str(_LOTE_DEFAULT)))
        delay        = int(in_config.get("SegFarmatodo",  "300"))

        # ── Cargar selectores desde BD ────────────────────────────────────
        selectores = _cargar_selectores(in_config, task_name)

        # ── PASO 1 + PASO 2: Verificar registros (debug → BD Dev / normal → SQL Server) ──
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
            cursor.execute(f"""
                SELECT a.Id FROM {esquema}.{tabla_ins} a
                LEFT JOIN {esquema}.{tabla_ex} b ON a.Id = b.Id
                WHERE b.Id IS NULL AND a.Estado='1'
            """)
            if cursor.fetchone() is not None:
                cursor.execute(f"""
                    INSERT INTO {esquema}.{tabla_ex}
                        ([Id],[FechaInicio],[FechaModificacion],[FechaFin],
                         [Estado],[Maquina],[PLU],[EAN],[Descripcion],
                         [MarcaProducto],[NombrePrd],[RegistroInvima],
                         [PrecioUnitario],[PrecioConDescuento],[PrecioSinDescuento],
                         [Porc.Descuento],[PrecioFidelizacion],
                         [UrlProducto],[BannerProducto],[RutaImagen])
                    SELECT a.[Id], a.[FechaInicio], GETDATE(), '',
                           '1', '{maquina}', a.[PLU], a.[EAN], a.[Descripcion],
                           '','','','','','','','','','',''
                    FROM {esquema}.{tabla_ins} a
                    LEFT JOIN {esquema}.{tabla_ex} b ON a.Id = b.Id
                    WHERE b.Id IS NULL AND a.Estado='1'
                """)
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
        headless = False if debug else str(in_config.get("HeadlessChrome", "true")).lower() == "true"
        write_log("Info", "HU02: Inicia consulta de productos por EAN", task_name, in_config)

        if debug:
            _scraping_debug(in_config, esquema, tabla_ins, url_template, selectores,
                            ruta_ss_base, task_name)
        else:
            _scraping_normal(in_config, esquema, tabla_ex, url_template, selectores,
                             ruta_ss_base, maquina, headless, lote, delay, task_name)

        write_log("Info", "HU02: Termina consulta de productos por EAN", task_name, in_config)

        # ── PASO 5: Limpieza de puntos de miles ───────────────────────────
        if not debug:
            conn   = conectar_bd(in_config)
            cursor = conn.cursor()
            for col in ("PrecioConDescuento", "PrecioSinDescuento"):
                cursor.execute(f"""
                    UPDATE {esquema}.{tabla_ex}
                    SET [{col}] = REPLACE([{col}], '.', '')
                    WHERE Estado='2' OR Estado='100'
                """)
            conn.commit()
            conn.close()

        # ── PASO 6: Reporte ───────────────────────────────────────────────
        if not debug:
            _generar_reportes(in_config, esquema, tabla_ex, task_name)

        write_log("Info", "Finaliza HU02", task_name, in_config)

    except Exception as e:
        out_system_exception = str(e)
        write_log("Error", f"HU02: {e}", task_name, in_config)
        write_log("Info", "Finaliza HU02", task_name, in_config)

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    return out_system_exception


def _scraping_normal(in_config, esquema, tabla_ex, url_template, selectores,
                     ruta_ss_base, maquina, headless, lote, delay, task_name):
    hay_mas = True
    while hay_mas:
        conn   = conectar_bd(in_config)
        cursor = conn.cursor()
        cursor.execute(f"""
            SELECT TOP({lote}) [Id], [EAN],
                LEFT(LTRIM(SUBSTRING(Descripcion,
                    PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100)),
                    CHARINDEX(' ', LTRIM(SUBSTRING(Descripcion,
                        PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100)) + ' ') - 1)
            FROM {esquema}.{tabla_ex} WHERE Estado='1'
        """)
        registros = cursor.fetchall()
        conn.close()

        if not registros:
            break

        driver = _crear_driver(headless=headless)
        try:
            for row in registros:
                id_t, ean, kw = str(row[0]), str(row[1]), str(row[2] or "")
                ruta_ss = os.path.join(ruta_ss_base, f"{ean}_{id_t}.jpg")

                conn   = conectar_bd(in_config)
                cursor = conn.cursor()
                cursor.execute(f"UPDATE {esquema}.{tabla_ex} SET FechaModificacion=GETDATE() WHERE Id='{id_t}'")
                conn.commit()
                conn.close()

                write_log("Info", f"HU02: Consultando EAN ({ean})", task_name, in_config)
                res = _consultar_ean_farmatodo(driver, ean, kw, url_template, selectores,
                                               ruta_ss, in_config, task_name)
                _persistir(in_config, esquema, tabla_ex, id_t, ruta_ss, res, task_name)
        finally:
            try:
                driver.quit()
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


def _persistir(in_config, esquema, tabla_ex, id_t, ruta_ss, res, task_name):
    estado       = res["estado"]
    nombre_prd   = res["nombre_prd"].replace(";", "").replace("'", "''")
    marca        = res["marca"].replace("'", "''")
    invima       = res["registro_invima"].replace("'", "''")
    precio_con   = res["precio_con_desc"]
    precio_sin   = res["precio_sin_desc"]
    precio_unit  = res["precio_unitario"].replace("'", "''")
    precio_fid   = res["precio_fidelizacion"]
    porc_desc    = res["porc_descuento"]
    url_prd      = res["url_producto"].replace("'", "''")
    banner       = res["banner"].replace("'", "''")
    ruta_img     = ruta_ss.replace("'", "''")

    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    if estado == "99":
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='99',
                [UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)
    elif estado == "3":
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='3',
                [NombrePrd]='{nombre_prd}',[MarcaProducto]='{marca}',
                [UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)
    else:
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='2',
                [NombrePrd]='{nombre_prd}',[MarcaProducto]='{marca}',
                [RegistroInvima]='{invima}',
                [PrecioConDescuento]='{precio_con}',[PrecioSinDescuento]='{precio_sin}',
                [PrecioUnitario]='{precio_unit}',[PrecioFidelizacion]='{precio_fid}',
                [Porc.Descuento]='{porc_desc}',[BannerProducto]='{banner}',
                [UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)

    conn.commit()
    conn.close()


def _scraping_debug(in_config, esquema, tabla_ins, url_template, selectores,
                    ruta_ss_base, task_name):
    """Lee de pruebas.db, hace scraping, escribe en pruebas.db (Farmatodo) y genera Excel."""
    lote_debug = int(in_config.get("LoteDebug", "3"))
    conn_sq = conectar_bd_debug(in_config)
    cur_sq  = conn_sq.cursor()
    cur_sq.execute(
        f"SELECT TOP (?) Id, EAN, Descripcion FROM {esquema}.TicketInsumo WHERE Estado=1",
        (lote_debug,)
    )
    registros = cur_sq.fetchall()

    if not registros:
        write_log("Info", "[DEBUG] No hay registros en pruebas.db TicketInsumo con Estado=1",
                  task_name, in_config)
        conn_sq.close()
        return

    resultados = []
    driver = _crear_driver(headless=False)
    try:
        for row in registros:
            id_t = str(row[0])
            ean  = str(row[1])
            desc = str(row[2] or "")
            m    = re.search(r'[a-zA-Z]{3,}', desc)
            kw   = m.group(0) if m else ""
            ruta_ss = os.path.join(ruta_ss_base, f"{ean}_{id_t}.jpg")
            print(f"\n  EAN: {ean}  |  {desc[:50]}")
            res = _consultar_ean_farmatodo(driver, ean, kw, url_template, selectores,
                                           ruta_ss, in_config, task_name)
            print(f"  Estado: {res['estado']} | Nombre: {res['nombre_prd']} | Precio: {res['precio_con_desc']}")
            resultados.append({"Id": id_t, "EAN": ean, "Descripcion": desc, "RutaImagen": ruta_ss, **res})
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    if resultados:
        ahora   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        maquina = socket.gethostname()
        cur_sq.execute(f"DELETE FROM {esquema}.Farmatodo")
        for r in resultados:
            cur_sq.execute(
                f"INSERT INTO {esquema}.Farmatodo "
                "(FechaInicio, FechaModificacion, FechaFin, Estado, Reintentos, Maquina, "
                " PLU, EAN, Descripcion, Categoria, HoraConsulta, MarcaProducto, NombrePrd, RegistroInvima, "
                " PrecioUnitario, PrecioConDescuento, PrecioSinDescuento, [Porc.Descuento], PrecioFidelizacion, "
                " BannerProducto, UrlProducto, RutaImagen) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (ahora, ahora, ahora,
                 r.get("estado", "99"), 0, maquina,
                 "", r["EAN"], r["Descripcion"], "", ahora,
                 r.get("marca", ""), r.get("nombre_prd", ""),
                 r.get("registro_invima", ""),
                 r.get("precio_unitario", ""), r.get("precio_con_desc", ""),
                 r.get("precio_sin_desc", ""), r.get("porc_descuento", ""),
                 r.get("precio_fidelizacion", ""), r.get("banner", ""),
                 r.get("url_producto", ""), r.get("RutaImagen", ""))
            )
        conn_sq.commit()
        write_log("Info",
                  f"[DEBUG] {len(resultados)} registros guardados en BD Dev ({esquema}.Farmatodo)",
                  task_name, in_config)

        ruta_debug = _PROJECT_ROOT / "debug"
        ruta_debug.mkdir(exist_ok=True)
        sello      = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_excel = str(ruta_debug / f"DEBUG_ReportePricingFarmatodo_{sello}.xlsx")
        pd.DataFrame(resultados).to_excel(ruta_excel, index=False)
        write_log("Info", f"[DEBUG] Reporte en ({ruta_excel})", task_name, in_config)
        print(f"\n  Reporte debug: {ruta_excel}")
    conn_sq.close()


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
        SELECT [FechaInicio],[PLU],[Descripcion],[FechaModificacion],[EAN],[Estado],
               [MarcaProducto],[NombrePrd],[RegistroInvima],[PrecioUnitario],
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
