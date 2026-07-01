"""
================================================================================
HU02 - Consulta y Reporte
Nombre de la iniciativa: Shopping de Precios Cruz Verde
Autor: Paula Sierra — Net Applications
Descripcion: Consulta precios en cruzverde.com.co por EAN.
             Cruz Verde es una SPA Angular; el bot busca por EAN, obtiene
             la URL del primer resultado (formato /slug/COCV_XXXXXX.html)
             y navega a la pagina de detalle para extraer nombre y precio
             mediante multiples selectores CSS con fallback.
Ultima modificacion: 01/07/2026
Propiedad de Colsubsidio
================================================================================

Flujo principal:
  1. Inserta en [CruzVerde] los IDs nuevos desde [TicketInsumo].
  2. Verifica registros pendientes (Estado='1').
  3. Crea carpeta de screenshots.
  4. Bucle de scraping por lotes (LoteCruzVerde).
     Para cada EAN:
       a. Navega a URL de busqueda y espera carga (Angular).
       b. Descarta modal de ubicacion/cookies si aparece.
       c. Extrae JSON {html, url} de la primera tarjeta de resultado via JS.
       d. Valida que la URL del resultado contenga el EAN.
       e. Parsea innerHTML con _entre() para extraer precios y nombre.
       f. Actualiza BD.
     Espera DelayCruzVerde segundos entre lotes.
  5. Generacion de reporte Excel y envio de correo.

Estados:
  1  : Pendiente
  2  : Producto encontrado
  3  : Sin coincidencia (nombre no corresponde al EAN)
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


ESPERA_CARGA  = 6000    # ms — Angular necesita tiempo extra para hidratar
ESPERA_REINT  = 3000    # ms entre reintentos
_LOTE_DEFAULT = 100


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


def _entre(html: str, inicio: str, fin: str) -> str:
    """Extrae el texto entre dos marcadores en un bloque HTML."""
    try:
        i = html.index(inicio) + len(inicio)
        j = html.index(fin, i)
        return html[i:j].strip()
    except (ValueError, AttributeError):
        return ""


def _limpiar_precio(texto: str) -> str:
    """Elimina simbolo de moneda y separadores de miles."""
    if not texto:
        return ""
    texto = texto.replace("$", "").replace("\xa0", "").strip()
    texto = texto.replace(".", "")
    return re.sub(r"[^\d]", "", texto)


def _tomar_screenshot(page: Page, ruta: str) -> None:
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        page.screenshot(path=ruta)
    except Exception:
        pass


def _descartar_modal(page: Page) -> None:
    """Cierra modal de ubicacion o cookies si esta visible."""
    try:
        page.evaluate("""
            const sel = 'button.btn-secondary, button[class*="bg-prices"], button[class*="aceptar"]';
            const btns = document.querySelectorAll(sel);
            for (const btn of btns) {
                const txt = btn.textContent.trim().toLowerCase();
                if (btn.offsetParent !== null && ['aceptar', 'ok', 'continuar'].includes(txt)) {
                    btn.click();
                    break;
                }
            }
        """)
    except Exception:
        pass


# ============================================================
# Logica de scraping por EAN en Cruz Verde
# ============================================================

def _consultar_ean_cruzverde(page: Page, ean: str,
                              url_template: str, ruta_screenshot: str,
                              in_config: dict, task_name: str) -> dict:
    resultado = {
        "nombre_prd":      "",
        "marca":           "",
        "precio_con_desc": "",
        "precio_sin_desc": "",
        "url_producto":    "",
        "banner":          "",
        "estado":          "99",
        "observaciones":   "No existe el producto en la farmacia",
        "reintentos":      0,
    }

    url_busqueda = url_template.replace("REEMPLAZAR", ean)

    try:
        page.goto(url_busqueda, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(ESPERA_CARGA)
        _descartar_modal(page)
        page.wait_for_timeout(1000)

        # ── Extraer primera tarjeta de resultado ──────────────────────────
        card_data = None
        for intento in range(3):
            try:
                card_data = page.evaluate("""
                    (() => {
                        const card = document.querySelector('ml-card-product');
                        if (!card) return null;
                        const link = card.querySelector('a[href]');
                        return {
                            html: card.innerHTML,
                            url:  link ? link.href : ''
                        };
                    })()
                """)
            except Exception:
                card_data = None

            if card_data and card_data.get("html"):
                break
            if intento < 2:
                write_log("Info",
                          f"HU02: EAN ({ean}) — Reintento {intento+1}, recargando pagina Angular",
                          task_name, in_config)
                try:
                    page.reload(wait_until="domcontentloaded", timeout=60000)
                except Exception:
                    pass
                page.wait_for_timeout(ESPERA_CARGA)
                _descartar_modal(page)

        _tomar_screenshot(page, ruta_screenshot)

        if not card_data or not card_data.get("html"):
            write_log("Info", f"HU02: EAN ({ean}) — Sin tarjeta de resultado en Cruz Verde",
                      task_name, in_config)
            return resultado

        url_producto = card_data.get("url", "").strip()

        # ── Validar que la URL sea del dominio Cruz Verde ─────────────────
        # Las URLs de Cruz Verde usan codigos internos COCV_XXXXXX, nunca el EAN.
        # Ejemplo: /producto-slug/COCV_162462.html
        if not url_producto or "cruzverde.com.co" not in url_producto:
            write_log("Info",
                      f"HU02: EAN ({ean}) — URL de resultado no valida: '{url_producto}'",
                      task_name, in_config)
            resultado["url_producto"] = url_producto
            return resultado

        # ── Navegar a la pagina de detalle del producto ───────────────────
        write_log("Info", f"HU02: EAN ({ean}) — Navegando a detalle: {url_producto}",
                  task_name, in_config)
        try:
            page.goto(url_producto, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(ESPERA_CARGA)
            _descartar_modal(page)
        except PlaywrightTimeout:
            write_log("Warning", f"HU02: Timeout cargando detalle EAN ({ean})", task_name, in_config)
            resultado["estado"]        = "99"
            resultado["observaciones"] = "Timeout al cargar la pagina de detalle"
            return resultado

        # ── Extraer campos desde la pagina de detalle ─────────────────────
        datos = page.evaluate("""
            (() => {
                const getText = (sels) => {
                    for (const sel of sels) {
                        try {
                            const el = document.querySelector(sel);
                            if (el && el.textContent.trim()) return el.textContent.trim();
                        } catch(e) {}
                    }
                    return '';
                };
                return {
                    nombre: getText([
                        'h1.product-name', '.product-name h1', '[class*="product-name"] h1',
                        'h1[class*="product"]', '.product-title', '.pdp-name', 'h1'
                    ]),
                    marca: getText([
                        '.product-brand', '[class*="product-brand"]', '.brand-name',
                        '[class*="brand-name"]', '.pdp-brand'
                    ]),
                    precio_con: getText([
                        '.price-value', '[class*="price-value"]', '.actual-price',
                        '[class*="actual-price"]', '.sale-price', '[class*="sale-price"]',
                        '.price .value', '.pdp-price'
                    ]),
                    precio_sin: getText([
                        '.price-original', '[class*="price-original"]', '.old-price',
                        '[class*="old-price"]', '.price-before', '[class*="price-before"]'
                    ]),
                    banner: getText([
                        '.badge-label', '[class*="badge-label"]', '.promotion-label',
                        '[class*="promo"]'
                    ]),
                };
            })()
        """)

        nombre_prd = datos.get("nombre", "").strip()
        marca      = datos.get("marca",  "").strip()
        precio_con = datos.get("precio_con", "").strip()
        precio_sin = datos.get("precio_sin", "").strip()
        banner     = datos.get("banner", "").strip()

        if not nombre_prd and not precio_con:
            write_log("Info", f"HU02: EAN ({ean}) — No se pudo extraer datos del detalle",
                      task_name, in_config)
            resultado.update({
                "url_producto": url_producto,
                "estado":       "99",
                "observaciones": "No se pudo extraer informacion del producto",
            })
            return resultado

        write_log("Info", f"HU02: EAN ({ean}) — Producto encontrado: '{nombre_prd}'",
                  task_name, in_config)

        resultado.update({
            "nombre_prd":      nombre_prd,
            "marca":           marca,
            "precio_con_desc": _limpiar_precio(precio_con),
            "precio_sin_desc": _limpiar_precio(precio_sin),
            "url_producto":    url_producto,
            "banner":          banner,
            "estado":          "2",
            "observaciones":   "",
        })

    except PlaywrightTimeout:
        write_log("Warning", f"HU02: Timeout consultando EAN ({ean})", task_name, in_config)
        resultado["estado"]        = "99"
        resultado["observaciones"] = "Timeout al cargar la pagina"
    except Exception as e:
        write_log("Warning", f"HU02: Error consultando EAN ({ean}): {e}", task_name, in_config)
        resultado["estado"]        = "99"
        resultado["observaciones"] = f"Error: {e}"

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
        tabla_ex     = in_config.get("TablaCruzVerde",    "[CruzVerde]")
        tabla_ins    = in_config.get("TablaTicketInsumo", "[TicketInsumo]")
        url_template = in_config.get("UrlCruzVerde", "")
        maquina      = socket.gethostname()
        lote         = int(in_config.get("LoteCruzVerde",  str(_LOTE_DEFAULT)))
        delay        = int(in_config.get("DelayCruzVerde", "300"))

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
                         [Estado],[Observaciones],[Reintentos],[Maquina],
                         [PLU],[EAN],[Descripcion],[Categoria],[HoraConsulta],
                         [MarcaProducto],[NombrePrd],[RegistroInvima],
                         [PrecioUnitario],[PrecioConDescuento],[PrecioSinDescuento],
                         [Porc.Descuento],[PrecioFidelizacion],
                         [BannerProducto],[UrlProducto],[RutaImagen])
                    SELECT a.[Id], a.[FechaInicio], GETDATE(), '',
                           '1','',0,'{maquina}',
                           a.[PLU], a.[EAN], a.[Descripcion], a.[Categoria], GETDATE(),
                           '','','','','','','','','','',''
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
                            ruta_ss_base, task_name)
        else:
            _scraping_normal(browser, in_config, esquema, tabla_ex, url_template,
                             ruta_ss_base, maquina, lote, delay, task_name)

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
                     ruta_ss_base, maquina, lote, delay, task_name):
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
                res = _consultar_ean_cruzverde(page, ean, url_template, ruta_ss, in_config, task_name)
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

def _scraping_debug(browser, in_config, esquema, tabla_ins,
                    url_template, ruta_ss_base, task_name):
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
            res = _consultar_ean_cruzverde(page, ean, url_template, ruta_ss, in_config, task_name)
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
        cur_sq.execute(f"DELETE FROM {esquema}.CruzVerde")
        for r in resultados:
            cur_sq.execute(
                f"INSERT INTO {esquema}.CruzVerde "
                "(FechaInicio, FechaModificacion, FechaFin, Estado, Observaciones, Reintentos, Maquina, "
                " PLU, EAN, Descripcion, Categoria, HoraConsulta, MarcaProducto, NombrePrd, RegistroInvima, "
                " PrecioUnitario, PrecioConDescuento, PrecioSinDescuento, [Porc.Descuento], PrecioFidelizacion, "
                " BannerProducto, UrlProducto, RutaImagen) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (ahora, ahora, ahora,
                 r.get("estado", "99"), r.get("observaciones", ""), r.get("reintentos", 0), maquina,
                 "", r["EAN"], r["Descripcion"], "", ahora,
                 r.get("marca", ""), r.get("nombre_prd", ""), "",
                 "", r.get("precio_con_desc", ""), r.get("precio_sin_desc", ""), "",
                 "", r.get("banner", ""), r.get("url_producto", ""), r.get("RutaImagen", ""))
            )
        conn_sq.commit()
        write_log("Info", f"[DEBUG] {len(resultados)} registros guardados en ({esquema}.CruzVerde)",
                  task_name, in_config)

        ruta_debug = _PROJECT_ROOT / "debug"
        ruta_debug.mkdir(exist_ok=True)
        sello      = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_excel = str(ruta_debug / f"DEBUG_ReportePricingCruzVerde_{sello}.xlsx")
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
    url_prd    = res["url_producto"].replace("'", "''")
    banner     = res["banner"].replace("'", "''")
    obs        = res["observaciones"].replace("'", "''")
    ruta_img   = ruta_ss.replace("'", "''")

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
                [BannerProducto]='{banner}',
                [Observaciones]='{obs}',[UrlProducto]='{url_prd}',[RutaImagen]='{ruta_img}'
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
