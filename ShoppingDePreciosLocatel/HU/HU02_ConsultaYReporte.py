"""
================================================================================
HU02 - Consulta y Reporte
Nombre de la iniciativa: Shopping de Precios Locatel
Autor: KPMG Advisory, Tax & Legal SAS
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
  5. Bucle de scraping: procesa lotes de LoteLocatel productos con Selenium
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
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyodbc

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
from Funciones.utils import write_log, conectar_bd, csv_a_excel, enviar_correo


# ============================================================
# Tiempos de espera (segundos) — equivalentes a cEsperaXs AA
# ============================================================
ESPERA_3S = 3
ESPERA_5S = 5
ESPERA_7S = 7


# ============================================================
# Helpers de Selenium
# ============================================================

def _crear_driver(headless: bool = True) -> webdriver.Chrome:
    """Crea y retorna un ChromeDriver configurado."""
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


def _js(driver, script: str, default=""):
    """Ejecuta JavaScript y retorna el resultado; retorna `default` si falla."""
    try:
        result = driver.execute_script(f"return {script}")
        return result if result is not None else default
    except Exception:
        return default


def _extraer_precio_entero(texto: str) -> str:
    """
    Limpia un texto de precio: elimina espacios, puntos de miles y signo $.
    Ejemplo: "$ 15.000" -> "15000"
    """
    if not texto:
        return ""
    return re.sub(r"[^\d]", "", texto)


def _tomar_screenshot(driver, ruta: str) -> None:
    """Toma screenshot de la pagina actual y lo guarda en ruta."""
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        driver.save_screenshot(ruta)
    except Exception:
        pass


# ============================================================
# Logica de scraping por EAN en Locatel
# ============================================================

def _consultar_ean(driver, ean: str, palabra_clave: str, url_template: str,
                   ruta_screenshot: str, in_config: dict, task_name: str) -> dict:
    """
    Abre la URL de busqueda de Locatel para el EAN dado y extrae:
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
        driver.get(url_consulta)
        time.sleep(ESPERA_5S)

        # ── Obtener URL del primer producto en resultados ──────────────
        url_producto = ""
        for _ in range(5):
            url_producto = _js(
                driver,
                "document.getElementsByClassName("
                "'vtex-product-summary-2-x-clearLink "
                "vtex-product-summary-2-x-clearLink--itemList "
                "h-100 flex flex-column').item(0)?.href"
            )
            if url_producto:
                break
            time.sleep(ESPERA_3S)

        resultado["url_producto"] = url_producto or url_consulta

        # ── Titulo del producto ────────────────────────────────────────
        titulo = _js(
            driver,
            "document.querySelector('.vtex-product-summary-2-x-productBrand')"
            "?.innerText.trim()"
        )
        resultado["titulo"] = titulo or ""

        # Si no hay titulo en la pagina de resultados no hay producto
        if not titulo:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — No existe el producto en la farmacia",
                task_name, in_config
            )
            _tomar_screenshot(driver, ruta_screenshot)
            resultado["estado"]       = "99"
            resultado["observaciones"] = "No existe el producto en la farmacia"
            return resultado

        # ── Precios ───────────────────────────────────────────────────
        # Precio de venta (con descuento): primer elemento de precio visible
        precio_con_desc_raw = _js(
            driver,
            "[...document.querySelectorAll("
            "'.vtex-product-price-1-x-sellingPriceValue "
            ".vtex-product-summary-2-x-currencyInteger')]"
            ".map(el => el.innerText.trim())[0] || "
            "document.querySelector("
            "'.vtex-product-summary-2-x-currencyInteger')?.innerText.trim()"
        )
        # Precio original (sin descuento): precio de lista si existe
        precio_sin_desc_raw = _js(
            driver,
            "document.querySelector("
            "'.vtex-product-price-1-x-listPriceValue "
            ".vtex-product-summary-2-x-currencyInteger')?.innerText.trim()"
        )

        precio_con_desc = _extraer_precio_entero(str(precio_con_desc_raw or ""))
        precio_sin_desc = _extraer_precio_entero(str(precio_sin_desc_raw or ""))

        # Si no hay precio de lista, el precio de venta ES el precio sin descuento
        if not precio_sin_desc:
            precio_sin_desc = precio_con_desc
            precio_con_desc = "0"  # sin descuento

        resultado["precio_con_desc"] = precio_con_desc
        resultado["precio_sin_desc"] = precio_sin_desc

        # ── Disponibilidad / Stock ─────────────────────────────────────
        disponibilidad_raw = _js(
            driver,
            "(document.querySelector("
            "'.locatelcolombia-delivery-modal-0-x-buttonNoPdp')"
            "?.innerText.trim()) || 'Texto no encontrado'"
        )
        resultado["disponibilidad"] = str(disponibilidad_raw or "Texto no encontrado")

        # ── Marca ─────────────────────────────────────────────────────
        marca_raw = _js(
            driver,
            "document.querySelector("
            "'.vtex-product-summary-2-x-brandName')?.innerText.trim()"
        )
        resultado["marca"] = str(marca_raw or "")

        # ── Determinar estado del registro ────────────────────────────
        # Validar que el titulo corresponda a la palabra clave del EAN
        titulo_upper = titulo.upper()
        kw_upper     = (palabra_clave or "").upper().strip()

        if kw_upper and kw_upper not in titulo_upper:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — Sin coincidencia: titulo='{titulo}', "
                f"palabra_clave='{palabra_clave}'",
                task_name, in_config
            )
            _tomar_screenshot(driver, ruta_screenshot)
            resultado["estado"]       = "3"
            resultado["observaciones"] = (
                "No existe coincidencia entre la información encontrada "
                "y el producto consultado"
            )
            return resultado

        # Verificar stock
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
            resultado["estado"]       = "2"
            resultado["observaciones"] = "Sin stock"
        else:
            write_log(
                "Info",
                f"HU02: EAN ({ean}) — Producto encontrado: '{titulo}' "
                f"precio={precio_sin_desc}",
                task_name, in_config
            )
            resultado["estado"]       = "2"
            resultado["observaciones"] = ""

        _tomar_screenshot(driver, ruta_screenshot)

    except Exception as e:
        write_log(
            "Warning",
            f"HU02: Error consultando EAN ({ean}): {e}",
            task_name, in_config
        )
        resultado["estado"]       = "99"
        resultado["observaciones"] = f"Error: {e}"

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

    driver = None

    try:
        esquema      = in_config.get("Scheme", "[ShoppingDePrecios]")
        tabla_loc    = in_config.get("TablaLocatel",     "Locatel")
        tabla_ins    = in_config.get("TablaTicketInsumo","TicketInsumo")
        url_template = in_config.get("UrlLocatel", "")
        lote         = int(in_config.get("LoteLocatel", "10"))
        reintentos_r = in_config.get("ReintentosReprocesamiento", "3")
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

        # ----------------------------------------------------------------
        # PASO 2: Insertar en TablaLocatel los IDs nuevos de TicketInsumo
        # ----------------------------------------------------------------
        cursor.execute(f"""
            SELECT a.Id
            FROM {esquema}.{tabla_ins} a
            LEFT JOIN {esquema}.{tabla_loc} b ON a.Id = b.Id
            WHERE b.Id IS NULL AND a.Estado='1'
        """)
        hay_nuevos = cursor.fetchone() is not None

        if hay_nuevos:
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
                    a.[Id], a.[FechaInicio], GETDATE(), '',
                    '1', '', '0', '{maquina}', a.[FechaInicio],
                    a.[PLU], a.[EAN], a.[Descripcion], a.[Categoria],
                    '','','','','','','','','','','',''
                FROM {esquema}.{tabla_ins} a
                LEFT JOIN {esquema}.{tabla_loc} b ON a.Id = b.Id
                WHERE b.Id IS NULL AND a.Estado='1'
            """)
            write_log(
                "Info",
                f"HU02: Existen nuevos registros para cargar a la tabla ({tabla_loc})",
                task_name, in_config
            )

        # ----------------------------------------------------------------
        # PASO 3: Verificar si hay registros pendientes
        # ----------------------------------------------------------------
        cursor.execute(f"""
            SELECT TOP(1) * FROM {esquema}.{tabla_loc}
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
        # PASO 4: Estructura de carpetas de screenshots
        # ----------------------------------------------------------------
        now = datetime.now()
        ruta_screenshots = os.path.join(
            in_config.get("RutaScreenshots", ""),
            in_config.get("CarpetaLocatel", ""),
            str(now.year),
            f"{now.month:02d}",
            f"{now.day:02d}"
        )
        os.makedirs(ruta_screenshots, exist_ok=True)

        # ----------------------------------------------------------------
        # PASO 5: Bucle principal de scraping
        # ----------------------------------------------------------------
        headless = str(in_config.get("HeadlessChrome", "true")).lower() == "true"
        write_log("Info", "HU02: Inicia consulta de productos por EAN", task_name, in_config)

        hay_mas = True
        while hay_mas:
            write_log(
                "Info",
                f"HU02: Se consultara un lote de ({lote}) productos.",
                task_name, in_config
            )

            conn   = conectar_bd(in_config)
            cursor = conn.cursor()

            # Extraer lote de registros pendientes
            cursor.execute(f"""
                SELECT TOP({lote}) [Id], [EAN],
                    LEFT(
                        LTRIM(SUBSTRING(Descripcion,
                            PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100)),
                        CHARINDEX(' ',
                            LTRIM(SUBSTRING(Descripcion,
                                PATINDEX('%[a-zA-Z][a-zA-Z][a-zA-Z]%', Descripcion), 100))
                            + ' ') - 1
                    ),
                    [PLU]
                FROM {esquema}.{tabla_loc}
                WHERE Estado='1'
            """)
            registros = cursor.fetchall()
            conn.close()

            if not registros:
                hay_mas = False
                break

            driver = _crear_driver(headless=headless)

            for row in registros:
                id_ticket    = str(row[0])
                ean          = str(row[1])
                palabra_clave = str(row[2] or "")
                # plu        = str(row[3])  # disponible si se necesita

                # Actualizar FechaModificacion al iniciar consulta
                conn   = conectar_bd(in_config)
                cursor = conn.cursor()
                cursor.execute(f"""
                    UPDATE {esquema}.{tabla_loc}
                    SET FechaModificacion=GETDATE()
                    WHERE Id='{id_ticket}'
                """)
                conn.commit()
                conn.close()

                # Ruta del screenshot
                ruta_ss = os.path.join(
                    ruta_screenshots,
                    f"{ean}_{id_ticket}.jpg"
                )

                write_log(
                    "Info",
                    f"HU02: Se consultara el EAN ({ean}) en la ruta "
                    f"({url_template.replace('REEMPLAZAR', ean)})",
                    task_name, in_config
                )

                # Consultar EAN en Locatel
                resultado = _consultar_ean(
                    driver=driver,
                    ean=ean,
                    palabra_clave=palabra_clave,
                    url_template=url_template,
                    ruta_screenshot=ruta_ss,
                    in_config=in_config,
                    task_name=task_name
                )

                # Actualizar BD con resultado
                conn   = conectar_bd(in_config)
                cursor = conn.cursor()
                estado       = resultado["estado"]
                observaciones = resultado["observaciones"].replace("'", "''")
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

            # Cerrar navegador al terminar el lote
            try:
                driver.quit()
            except Exception:
                pass
            driver = None

            # Verificar si quedan registros en Estado=1
            conn   = conectar_bd(in_config)
            cursor = conn.cursor()
            cursor.execute(f"""
                SELECT TOP(1) * FROM {esquema}.{tabla_loc} WHERE Estado='1'
            """)
            hay_mas = cursor.fetchone() is not None
            conn.close()

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
        if driver:
            try:
                driver.quit()
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
            REPLACE(CONCAT(
                REPLACE(CAST(FechaInicio AS DATE),'-','_'),
                REPLACE(REPLACE(SUBSTRING(CAST(FechaInicio AS varchar),12,6),' ','_'),':','_')
            ),'__','_0')
        FROM {esquema}.{tabla_loc}
        WHERE Estado='2' OR Estado='99'
    """)
    fechas = cursor.fetchall()
    conn.close()

    for fecha_row in fechas:
        fecha_inicio = str(fecha_row[0])
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

    # Restaurar estados previos de reporte si los hubiera
    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='2' "
        f"WHERE [Estado]='100' AND FechaInicio='{fecha_inicio}'"
    )
    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='99' "
        f"WHERE [Estado]='199' AND FechaInicio='{fecha_inicio}'"
    )

    # Estadisticas
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
    total         = stats[0] if stats else 0
    extraidos     = stats[1] if stats else 0
    estado2_count = stats[2] if stats else 0
    sin_stock     = stats[3] if stats else 0
    estado99_count= stats[4] if stats else 0

    write_log(
        "Info",
        f"HU02: Reporte FechaInicio={fecha_inicio} — "
        f"Total={total}, Extraidos={extraidos}, Estado2={estado2_count}, "
        f"SinStock={sin_stock}, Estado99={estado99_count}",
        task_name, in_config
    )

    # Marcar como reportados
    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='100' "
        f"WHERE [Estado]='2' AND FechaInicio='{fecha_inicio}'"
    )
    cursor.execute(
        f"UPDATE {esquema}.{tabla_loc} SET [Estado]='199' "
        f"WHERE [Estado]='99' AND FechaInicio='{fecha_inicio}'"
    )

    # Ajustar PrecioConDescuento = 0 cuando es igual al PrecioSinDescuento
    cursor.execute(f"""
        UPDATE {esquema}.{tabla_loc}
        SET [PrecioConDescuento]='0'
        WHERE [Estado]='2'
          AND FechaInicio='{fecha_inicio}'
          AND TRY_CAST(PrecioSinDescuento AS INT) = TRY_CAST(PrecioConDescuento AS INT)
    """)

    # Calcular porcentaje de descuento
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

    conn.commit()

    # Consultar datos para el reporte
    cursor.execute(f"""
        SELECT
            [FechaInicio]       AS fechainsumo,
            [PLU]               AS plu,
            [Descripcion]       AS descripcion,
            [FechaModificacion] AS horaconsulta,
            [EAN]               AS ean,
            [Estado]            AS estado,
            [MarcaProducto]     AS marcaproducto,
            [NombrePrd]         AS nombreprd,
            [RegistroInvima]    AS registroinvima,
            [PrecioUnitario]    AS preciounitario,
            [PrecioConDescuento]AS preciocondescuento,
            [PrecioSinDescuento]AS preciosindescuento,
            [Porc.Descuento]    AS [porc.descuento],
            [PrecioFidelizacion]AS preciofidelizacion,
            [BannerProducto]    AS bannerproducto,
            [UrlProducto]       AS urlproducto,
            [RutaImagen]        AS rutaimagen
        FROM {esquema}.{tabla_loc}
        WHERE FechaInicio='{fecha_inicio}'
    """)
    cols   = [col[0] for col in cursor.description]
    filas  = cursor.fetchall()
    conn.close()

    if not filas:
        return

    # Construir DataFrame y exportar a Excel
    df = pd.DataFrame(filas, columns=cols)

    ruta_reporte = in_config.get("RutaReporte", "")
    os.makedirs(ruta_reporte, exist_ok=True)
    nombre_resultado = in_config.get("NombreResultado", "ReportePricingLOCATEL_")
    nombre_hoja      = in_config.get("NombreHojaResultado", "ReportePricingLOCATEL")
    nombre_excel     = f"{nombre_resultado}{fecha_sello}.xlsx"
    ruta_excel       = os.path.join(ruta_reporte, nombre_excel)

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)

    write_log(
        "Info",
        f"HU02: Reporte generado en ({ruta_excel})",
        task_name, in_config
    )

    # Enviar correo con reporte adjunto (codEmail=100)
    from_address = in_config.get("_correo", {}).get("usuario", "")
    reemplazo = {"$NombrePagina$": in_config.get("DrogueriaLocatel", "Locatel")}
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
        write_log("Info", f"HU02: No fue posible enviar el correo de notificacion: {err}", task_name, in_config)


if __name__ == "__main__":
    from Funciones.utils import obtener_config
    config = obtener_config()
    exc = hu02_consulta_y_reporte(config)
    if exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    print("HU02 completada exitosamente.")
