"""
================================================================================
HU01 - Validacion y Carga de Insumo
Nombre de la iniciativa: Shopping de Precios Locatel
Autor: KPMG Advisory, Tax & Legal SAS
Descripcion: Valida el archivo Excel de insumo y lo carga a la Base de Datos.
             Equivale al bot HU01_ValidacionYCargaInsumo de Automation Anywhere.
Ultima modificacion: 28/02/2025
Propiedad de Colsubsidio
================================================================================

Flujo:
  1. Verifica que el archivo insumo exista en RutaInsumos + ArchivoInsumo.
  2. Abre el Excel y valida que las columnas sean PLU / EAN / DESCRIPCION / PROVEEDOR.
  3. Convierte el Excel a CSV (separador ; / codificacion ANSI) en CarpetaTemp.
  4. Conecta a SQL Server, crea tabla temporal, hace BULK INSERT (o fallback fila-a-fila)
     y limpia EAN invalidos; luego inserta en TablaTicketInsumo.
  5. Copia el archivo a CarpetaProcesados con sello de fecha y elimina el original.
"""

import os
import sys
import time
import socket
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
from Funciones.utils import write_log, conectar_bd, conectar_bd_debug, excel_a_csv, enviar_correo


def hu01_validacion_y_carga_insumo(in_config: dict) -> str:
    """
    Ejecuta la validacion y carga del archivo insumo a la BD.

    Parametros:
        in_config: Diccionario de configuracion (ioConfig cargado en HU00).

    Retorna:
        '' si exitoso, mensaje de error si fallo.
    """
    out_system_exception = ""
    task_name = "HU01_ValidacionYCargaInsumo"
    write_log("Info", "Inicia HU01", task_name, in_config)

    try:
        # ----------------------------------------------------------------
        # PASO 1: Parametrizacion inicial
        # ----------------------------------------------------------------
        ruta_insumo = os.path.join(
            in_config.get("RutaInsumos", ""),
            in_config.get("ArchivoInsumo", "")
        )
        # En debug: usar Insumo/ local del proyecto (evita rutas de red)
        if in_config.get("_debug"):
            _nombre_ins = in_config.get("ArchivoInsumo", "InsumoPricing.xlsx")
            ruta_insumo = str(Path(__file__).resolve().parent.parent.parent / "Insumo" / _nombre_ins)

        # ----------------------------------------------------------------
        # PASO 2: Verificar existencia del archivo
        # ----------------------------------------------------------------
        if not os.path.isfile(ruta_insumo):
            write_log("Info", f"HU01: NO existe el archivo de la ruta ({ruta_insumo})", task_name, in_config)
            write_log("Info", "Finaliza HU01", task_name, in_config)
            return f"Archivo de insumo no encontrado: {ruta_insumo}"

        write_log("Info", f"HU01: Existe el archivo de la ruta ({ruta_insumo})", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 3: Validar encabezados del Excel
        # ----------------------------------------------------------------
        sheet_insumo = in_config.get("SheetTicketInsumo", "TicketInsumo")

        df_header = pd.read_excel(
            ruta_insumo,
            sheet_name=sheet_insumo,
            nrows=0,
            dtype=str,
            header=0
        )
        columnas = [str(c).strip().upper() for c in df_header.columns]

        # Columnas esperadas (A1=PLU, B1=EAN, C1=DESCRIPCION, D1=PROVEEDOR)
        header_ok = (
            len(columnas) >= 4
            and "PLU"         in columnas[0]
            and "EAN"         in columnas[1]
            and "DESCRIPCION" in columnas[2]
            and "PROVEEDOR"   in columnas[3]
        )

        if not header_ok:
            write_log(
                "Warning",
                "HU01: El archivo no cumple con la estructura de encabezados definida",
                task_name, in_config
            )
            write_log("Info", "Finaliza HU01", task_name, in_config)
            enviar_correo(
                in_config=in_config,
                i_cod_email=2,
                i_from_address=in_config.get("_correo", {}).get("usuario", ""),
                i_html_format=False
            )
            return "El archivo de insumo no cumple con la estructura de encabezados"

        write_log("Info", "HU01: El archivo cumple con la estructura", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 4: Convertir Excel a CSV temporal
        # ----------------------------------------------------------------
        ruta_red    = in_config.get("RutaRed", "")
        carpeta_tmp = in_config.get("CarpetaTemp", "")
        ruta_csv    = os.path.join(ruta_red, carpeta_tmp, "Insumo.csv")
        # En debug: CSV temporal en carpeta local debug/temp/
        if in_config.get("_debug"):
            _dir_tmp = Path(__file__).resolve().parent.parent.parent / "debug" / "temp"
            _dir_tmp.mkdir(parents=True, exist_ok=True)
            ruta_csv = str(_dir_tmp / "Insumo.csv")

        if os.path.isfile(ruta_csv):
            os.remove(ruta_csv)

        err = excel_a_csv(
            in_path_excel=ruta_insumo,
            in_path_out_csv=ruta_csv,
            in_name_sheet=sheet_insumo,
            in_last_column="F",
            in_decimal_separator=",",
            in_thousands_separator=".",
            in_config=in_config
        )
        if err:
            raise Exception(f"Error convirtiendo Excel a CSV: {err}")

        # ----------------------------------------------------------------
        # PASO 5: Cargar CSV a Base de Datos
        # ----------------------------------------------------------------
        maquina      = socket.gethostname()
        esquema      = in_config.get("Scheme", "[ShoppingDePrecios]")
        tabla_insumo = in_config.get("TablaTicketInsumo", "TicketInsumo")

        if in_config.get("_debug"):
            # DEBUG: insertar en BD Dev (SQL Server) en lugar de pruebas.db
            df_insumo = pd.read_csv(ruta_csv, sep=";", dtype=str, encoding="cp1252", errors="replace")
            df_insumo = df_insumo.fillna("")
            # Filtrar EAN invalidos
            mask_ean = df_insumo.iloc[:, 1].str.match(r"^\d+$", na=False)
            df_insumo = df_insumo[mask_ean]

            conn_sq = conectar_bd_debug(in_config)
            cur_sq  = conn_sq.cursor()
            # Limpiar tabla para que cada ejecucion debug empiece limpia
            cur_sq.execute(f"DELETE FROM {esquema}.{tabla_insumo}")
            ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for _, row in df_insumo.iterrows():
                vals = [str(v) for v in row.iloc[:5]]
                while len(vals) < 5:
                    vals.append("")
                cur_sq.execute(
                    f"INSERT INTO {esquema}.{tabla_insumo} "
                    "(FechaInicio, FechaModificacion, Estado, Observaciones, Maquina, "
                    " PLU, EAN, Descripcion, Proveedor, Categoria) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (ahora, ahora, 1, "", maquina,
                     vals[0], vals[1], vals[2], vals[3], vals[4])
                )
            conn_sq.commit()
            conn_sq.close()
            write_log("Info",
                      f"HU01: [DEBUG] {len(df_insumo)} registros cargados en BD Dev ({tabla_insumo})",
                      task_name, in_config)
        else:
            # PRODUCCION: SQL Server
            conn   = conectar_bd(in_config)
            cursor = conn.cursor()

            cursor.execute("""
                IF OBJECT_ID('tempdb.dbo.#TicketInsumo', 'U') IS NOT NULL
                    DROP TABLE #TicketInsumo;
                CREATE TABLE #TicketInsumo(
                    [PLU]         [varchar](100),
                    [EAN]         [varchar](100),
                    [Descripcion] [varchar](200),
                    [Proveedor]   [varchar](100),
                    [Categoria]   [varchar](100)
                )
            """)

            bulk_ok = False
            try:
                cursor.execute(f"""
                    BULK INSERT #TicketInsumo
                    FROM '{ruta_csv}'
                    WITH (
                        FORMAT       = 'CSV',
                        FIRSTROW     = 2,
                        FIELDTERMINATOR = ';',
                        ROWTERMINATOR   = '\\n',
                        CODEPAGE        = 'ACP'
                    )
                """)
                bulk_ok = True
                write_log("Info", "HU01: BULK INSERT ejecutado", task_name, in_config)
            except Exception as bulk_err:
                write_log(
                    "Info",
                    f"HU01: BULK INSERT no disponible ({bulk_err}), cargando fila a fila",
                    task_name, in_config
                )

            if not bulk_ok:
                df_insumo = pd.read_csv(ruta_csv, sep=";", dtype=str, encoding="cp1252", errors="replace")
                df_insumo = df_insumo.fillna("")
                for _, row in df_insumo.iterrows():
                    vals = [str(v) for v in row.iloc[:5]]
                    while len(vals) < 5:
                        vals.append("")
                    cursor.execute("INSERT INTO #TicketInsumo VALUES (?,?,?,?,?)", vals)

            cursor.execute("""
                DELETE FROM #TicketInsumo
                WHERE EAN IS NULL
                   OR EAN LIKE ''
                   OR EAN LIKE '%[^0-9]%'
            """)

            cursor.execute(f"""
                INSERT INTO {esquema}.{tabla_insumo}
                    ([FechaInicio],[FechaModificacion],[Estado],[Observaciones],[Maquina],
                     [PLU],[EAN],[Descripcion],[Proveedor],[Categoria])
                SELECT
                    GETDATE(), GETDATE(), '1', '', '{maquina}',
                    PLU, EAN, Descripcion, Proveedor, Categoria
                FROM #TicketInsumo
            """)

            conn.commit()
            conn.close()
            write_log("Info", "HU01: Insumo cargado exitosamente a BD", task_name, in_config)

        # ----------------------------------------------------------------
        # PASO 6: Mover archivo a carpeta Procesados y eliminar original
        # ----------------------------------------------------------------
        if in_config.get("_debug"):
            # En debug no se mueve el archivo (permite re-ejecutar sin regenerar insumo)
            write_log("Info",
                      "HU01: [DEBUG] Archivo insumo no movido (modo debug — reutilizable)",
                      task_name, in_config)
        else:
            ruta_insumos    = in_config.get("RutaInsumos", "")
            carpeta_proc    = in_config.get("CarpetaProcesados", "Procesados\\")
            ruta_procesados = os.path.join(ruta_insumos, carpeta_proc)
            os.makedirs(ruta_procesados, exist_ok=True)

            now = datetime.now()
            nombre_destino = (
                f"InsumoPricing_{now.year}_{now.month:02d}_{now.day:02d}"
                f"_{now.hour:02d}_{now.minute:02d}_{now.second:02d}.xlsx"
            )
            ruta_destino = os.path.join(ruta_procesados, nombre_destino)
            shutil.copy2(ruta_insumo, ruta_destino)

            time.sleep(1)
            os.remove(ruta_insumo)

        out_system_exception = ""
        write_log("Info", "Finaliza HU01", task_name, in_config)

    except Exception as e:
        out_system_exception = str(e)
        write_log("Error", f"HU01: {e}", task_name, in_config)
        write_log("Info", "Finaliza HU01", task_name, in_config)

    return out_system_exception


if __name__ == "__main__":
    from Funciones.utils import obtener_config
    config = obtener_config(Path(__file__).resolve().parent.parent / "Config" / "config.json")
    exc = hu01_validacion_y_carga_insumo(config)
    if exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    print("HU01 completada exitosamente.")
