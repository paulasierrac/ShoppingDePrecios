"""
================================================================================
HU02 - Consulta y Reporte
Nombre de la iniciativa: Shopping de Precios Cruz Verde
Autor: KPMG Advisory, Tax & Legal SAS
Descripcion: Consulta precios en cruzverde.com.co por EAN.
             A diferencia de otras farmacias, Cruz Verde extrae todos los datos
             directamente desde la pagina de RESULTADOS de busqueda, sin navegar
             a la pagina de detalle del producto.
             El sitio es Angular (ng-star-inserted / ml-card-product).
Ultima modificacion: 22/06/2026
Propiedad de Colsubsidio
================================================================================

Flujo principal:
  1. Inserta en [CruzVerde] los IDs nuevos desde [TicketInsumo].
  2. Verifica registros pendientes (Estado='1').
  3. Crea carpeta de screenshots.
  4. Bucle de scraping por lotes (LoteCruzVerde, default 1000).
     Para cada EAN:
       a. Navega a URL de busqueda y espera carga Angular.
       b. Descarta modal de ubicacion/cookies si aparece.
       c. Obtiene HTML de la primera tarjeta 'ml-card-product' via JS.
       d. Extrae URL, nombre, precios y disponibilidad del HTML.
       e. Valida nombre vs palabra clave.
       f. Actualiza BD.
     Espera DelayCruzVerde segundos entre lotes.
  5. Limpieza de puntos de miles en precios.
  6. Generacion de reporte Excel y envio de correo.

Estados:
  1  : Pendiente
  2  : Producto encontrado (incluye sin stock con Observaciones='Sin stock')
  3  : URL encontrada pero nombre no coincide
  99 : Sin informacion / no encontrado
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
from selenium.webdriver.common.by import By

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
from Funciones.utils import write_log, conectar_bd, conectar_bd_debug, enviar_correo


ESPERA_CARGA  = 6    # segundos para que cargue Angular
ESPERA_REINT  = 3
_LOTE_DEFAULT = 1000


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


def _entre(texto: str, antes: str, despues: str) -> str:
    if not texto or antes not in texto:
        return ""
    start = texto.index(antes) + len(antes)
    rest  = texto[start:]
    if despues not in rest:
        return rest.strip()
    return rest[: rest.index(despues)].strip()


def _limpiar_precio(texto: str) -> str:
    """Elimina simbolo de moneda, espacios y puntos de miles."""
    if not texto:
        return ""
    return re.sub(r"[^\d]", "", texto)


def _tomar_screenshot(driver, ruta: str) -> None:
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        driver.save_screenshot(ruta)
    except Exception:
        pass


def _descartar_modal(driver) -> None:
    """Cierra modal de ubicacion o cookies si esta visible."""
    try:
        btns = driver.find_elements(
            By.CSS_SELECTOR,
            "button.btn-secondary, button[class*='bg-prices'], button[class*='aceptar']"
        )
        for btn in btns:
            if btn.is_displayed() and btn.text.strip().lower() in ("aceptar", "ok", "continuar"):
                btn.click()
                time.sleep(1)
                break
    except Exception:
        pass


def _consultar_ean_cruzverde(driver, ean: str, palabra_clave: str,
                              url_template: str, ruta_screenshot: str,
                              in_config: dict, task_name: str) -> dict:
    resultado = {
        "nombre_prd":          "",
        "marca":               "",
        "precio_con_desc":     "",
        "precio_sin_desc":     "",
        "precio_unitario":     "",
        "precio_fidelizacion": "",
        "url_producto":        "",
        "banner":              "",
        "estado":              "99",
        "observaciones":       "No existe el producto en la farmacia",
    }

    url_busqueda = url_template.replace("REEMPLAZAR", ean)

    try:
        driver.get(url_busqueda)
        time.sleep(ESPERA_CARGA)
        _descartar_modal(driver)
        time.sleep(ESPERA_REINT)

        # ── Paso 1: Obtener HTML de la primera tarjeta ml-card-product ────
        card_html = ""
        url_producto = ""
        for intento in range(3):
            try:
                data = driver.execute_script("""
                    try {
                        var cards = document.querySelectorAll('ml-card-product');
                        if (!cards || cards.length === 0) return null;
                        var c = cards[0];
                        var a = c.querySelector('a[href]');
                        return {
                            html: c.innerHTML,
                            url:  a ? a.href : '',
                        };
                    } catch(e) { return null; }
                """)
                if data:
                    card_html    = data.get("html", "") or ""
                    url_producto = data.get("url",  "") or ""
                    if card_html:
                        break
            except Exception:
                pass
            time.sleep(ESPERA_REINT)

        if not card_html:
            write_log("Info", f"HU02: EAN ({ean}) — No se encontraron tarjetas de producto",
                      task_name, in_config)
            _tomar_screenshot(driver, ruta_screenshot)
            return resultado

        resultado["url_producto"] = url_producto

        # ── Paso 2: Extraer campos del HTML de la tarjeta ─────────────────
        # Precio con descuento: <span class="font-bold text-prices">$ 9.785</span>
        precio_con_raw = _entre(card_html, 'class="font-bold text-prices">', "<")
        if not precio_con_raw:
            precio_con_raw = _entre(card_html, "font-bold text-prices\">", "<")

        # Precio sin descuento (tachado): class que contiene "line-through"
        precio_sin_raw = ""
        if "line-through" in card_html:
            idx = card_html.index("line-through")
            bloque_sin = card_html[idx:]
            precio_sin_raw = _entre(bloque_sin, ">", "<").replace("(Normal)", "").strip()

        # Precio unitario (PUM): texto que empieza con "PUM:"
        precio_unit_raw = ""
        if "PUM:" in card_html:
            precio_unit_raw = _entre(card_html, "PUM:", "<").strip()
            precio_unit_raw = "PUM:" + precio_unit_raw

        # Nombre del producto
        nombre_prd = ""
        for marca_ini in ('class="text-sm font-bold', 'class="font-bold text-sm',
                          'class="truncate font-bold', 'ng-star-inserted">'):
            nombre_prd = _entre(card_html, marca_ini + '">', "<")
            if not nombre_prd:
                nombre_prd = _entre(card_html, marca_ini, "<")
            if nombre_prd and len(nombre_prd) > 3:
                break

        # Fallback nombre desde alt de imagen
        if not nombre_prd:
            nombre_prd = _entre(card_html, 'alt="', '"')

        # Disponibilidad: buscar "No disponible" en el HTML de la tarjeta
        sin_stock = "No disponible" in card_html or "no disponible" in card_html.lower()

        if not nombre_prd and not precio_con_raw:
            write_log("Info", f"HU02: EAN ({ean}) — No se pudo extraer datos de la tarjeta",
                      task_name, in_config)
            _tomar_screenshot(driver, ruta_screenshot)
            return resultado

        # ── Paso 3: Validar nombre vs palabra clave ────────────────────────
        kw = (palabra_clave or "").upper().strip()
        if kw and nombre_prd and kw not in nombre_prd.upper():
            write_log("Info",
                      f"HU02: EAN ({ean}) — Sin coincidencia: nombre='{nombre_prd}', kw='{kw}'",
                      task_name, in_config)
            _tomar_screenshot(driver, ruta_screenshot)
            resultado.update({
                "nombre_prd":    nombre_prd,
                "url_producto":  url_producto,
                "estado":        "3",
                "observaciones": "No existe coincidencia entre la informacion encontrada y el producto consultado",
            })
            return resultado

        write_log("Info", f"HU02: EAN ({ean}) — Producto encontrado: '{nombre_prd}'",
                  task_name, in_config)
        _tomar_screenshot(driver, ruta_screenshot)

        resultado.update({
            "nombre_prd":      nombre_prd,
            "precio_con_desc": _limpiar_precio(precio_con_raw),
            "precio_sin_desc": _limpiar_precio(precio_sin_raw),
            "precio_unitario": precio_unit_raw,
            "url_producto":    url_producto,
            "banner":          "No disponible" if sin_stock else "",
            "estado":          "2",
            "observaciones":   "Sin stock" if sin_stock else "",
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
        tabla_ex     = in_config.get("TablaCruzVerde",    "[CruzVerde]")
        tabla_ins    = in_config.get("TablaTicketInsumo", "[TicketInsumo]")
        url_template = in_config.get("UrlCruzVerde", "")
        maquina      = socket.gethostname()
        lote         = int(in_config.get("LoteCruzVerde",  str(_LOTE_DEFAULT)))
        delay        = int(in_config.get("DelayCruzVerde", "300"))

        # ── PASO 1 + PASO 2: Verificar registros (debug → SQLite / normal → SQL Server) ──
        if debug:
            conn_sq = conectar_bd_debug()
            cur_sq  = conn_sq.cursor()
            cur_sq.execute("SELECT COUNT(*) FROM TicketInsumo WHERE Estado=1")
            cnt = cur_sq.fetchone()[0] or 0
            conn_sq.close()
            hay_pendientes = cnt > 0
            if hay_pendientes:
                write_log("Info",
                          f"[DEBUG] {cnt} registros en pruebas.db TicketInsumo con Estado=1",
                          task_name, in_config)
            else:
                write_log("Info",
                          "[DEBUG] No hay registros en pruebas.db TicketInsumo con Estado=1",
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
                         [UrlProducto],[BannerProducto],[RutaImagen],
                         [Observaciones],[Reintentos])
                    SELECT a.[Id], a.[FechaInicio], GETDATE(), '',
                           '1', '{maquina}', a.[PLU], a.[EAN], a.[Descripcion],
                           '','','','','','','','','','','','',0
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
            ruta_ss_base = str(_PROJECT_ROOT / "debug" / "screenshots" / "CruzVerde"
                               / str(now.year) / f"{now.month:02d}" / f"{now.day:02d}")
        else:
            ruta_ss_base = os.path.join(
                in_config.get("RutaScreenshots", ""),
                in_config.get("CarpetaCruzVerde", "CruzVerde\\"),
                str(now.year), f"{now.month:02d}", f"{now.day:02d}",
            )
        os.makedirs(ruta_ss_base, exist_ok=True)

        # ── PASO 4: Bucle de scraping ─────────────────────────────────────
        headless = False if debug else str(in_config.get("HeadlessChrome", "true")).lower() == "true"
        write_log("Info", "HU02: Inicia consulta de productos por EAN", task_name, in_config)

        if debug:
            _scraping_debug(in_config, esquema, tabla_ins, url_template,
                            ruta_ss_base, task_name)
        else:
            _scraping_normal(in_config, esquema, tabla_ex, url_template,
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


def _scraping_normal(in_config, esquema, tabla_ex, url_template,
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
                res = _consultar_ean_cruzverde(driver, ean, kw, url_template,
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
    estado      = res["estado"]
    nombre_prd  = res["nombre_prd"].replace(";", "").replace("'", "''")
    marca       = res["marca"].replace("'", "''")
    precio_con  = res["precio_con_desc"]
    precio_sin  = res["precio_sin_desc"]
    precio_unit = res["precio_unitario"].replace("'", "''")
    precio_fid  = res["precio_fidelizacion"]
    url_prd     = res["url_producto"].replace("'", "''")
    obs         = res["observaciones"].replace("'", "''")
    ruta_img    = ruta_ss.replace("'", "''")

    conn   = conectar_bd(in_config)
    cursor = conn.cursor()

    if estado == "99":
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='99',
                [Observaciones]='{obs}',[UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)
    elif estado == "3":
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='3',
                [NombrePrd]='{nombre_prd}',[MarcaProducto]='{marca}',
                [Observaciones]='{obs}',[UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)
    else:
        cursor.execute(f"""
            UPDATE {esquema}.{tabla_ex}
            SET [FechaFin]=GETDATE(),[Estado]='2',
                [NombrePrd]='{nombre_prd}',[MarcaProducto]='{marca}',
                [PrecioConDescuento]='{precio_con}',[PrecioSinDescuento]='{precio_sin}',
                [PrecioUnitario]='{precio_unit}',[PrecioFidelizacion]='{precio_fid}',
                [BannerProducto]='{res["banner"]}',
                [Observaciones]='{obs}',[UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
            WHERE Id='{id_t}'
        """)

    conn.commit()
    conn.close()


def _scraping_debug(in_config, esquema, tabla_ins, url_template,
                    ruta_ss_base, task_name):
    """Lee de pruebas.db, hace scraping, escribe en pruebas.db (CruzVerde) y genera Excel."""
    lote_debug = int(in_config.get("LoteDebug", "3"))
    conn_sq = conectar_bd_debug()
    cur_sq  = conn_sq.cursor()
    cur_sq.execute(
        "SELECT Id, EAN, Descripcion FROM TicketInsumo WHERE Estado=1 LIMIT ?",
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
            res = _consultar_ean_cruzverde(driver, ean, kw, url_template,
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
        cur_sq.execute("DELETE FROM CruzVerde")
        for r in resultados:
            cur_sq.execute(
                "INSERT INTO CruzVerde "
                "(FechaInicio, FechaModificacion, FechaFin, Estado, Observaciones, Reintentos, Maquina, "
                " PLU, EAN, Descripcion, Categoria, HoraConsulta, MarcaProducto, NombrePrd, RegistroInvima, "
                " PrecioUnitario, PrecioConDescuento, PrecioSinDescuento, PorcDescuento, PrecioFidelizacion, "
                " BannerProducto, UrlProducto, RutaImagen) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (ahora, ahora, ahora,
                 r.get("estado", "99"), r.get("observaciones", ""), 0, maquina,
                 "", r["EAN"], r["Descripcion"], "", ahora,
                 r.get("marca", ""), r.get("nombre_prd", ""), "",
                 r.get("precio_unitario", ""), r.get("precio_con_desc", ""),
                 r.get("precio_sin_desc", ""), r.get("porc_descuento", ""),
                 r.get("precio_fidelizacion", ""), r.get("banner", ""),
                 r.get("url_producto", ""), r.get("RutaImagen", ""))
            )
        conn_sq.commit()
        write_log("Info",
                  f"[DEBUG] {len(resultados)} registros guardados en pruebas.db (CruzVerde)",
                  task_name, in_config)

        ruta_debug = _PROJECT_ROOT / "debug"
        ruta_debug.mkdir(exist_ok=True)
        sello      = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_excel = str(ruta_debug / f"DEBUG_ReportePricingCruzVerde_{sello}.xlsx")
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
               [PrecioFidelizacion],[BannerProducto],[UrlProducto],[RutaImagen],[Observaciones]
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
    nombre_hoja = in_config.get("NombreHojaResultado", "ReportePricingCruzVerde")
    ruta_excel  = os.path.join(ruta_rep, f"{nombre_res}{fecha_sello}.xlsx")

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)

    write_log("Info", f"HU02: Reporte generado en ({ruta_excel})", task_name, in_config)
    from_addr = in_config.get("_correo", {}).get("usuario", "")
    reemplazo = {"$NombrePagina$": in_config.get("DrogueriaCruzVerde", "Cruz Verde")}
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
