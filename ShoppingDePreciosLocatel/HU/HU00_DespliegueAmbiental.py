"""
================================================================================
HU00 - Despliegue Ambiental
Nombre de la iniciativa: Shopping de Precios Locatel
Autor: Paula Sierra — Net Applications
Descripcion: Conecta a BD, carga todos los parametros nombrados desde la tabla
             [ShoppingDePrecios].[Parametros], deriva rutas de trabajo y realiza
             limpieza de Logs, Screenshots, Reportes, Insumos y BD.
Ultima modificacion: 22/06/2026
Propiedad de Colsubsidio
================================================================================
"""

import os
import sys
import shutil
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
from Funciones.utils import write_log, obtener_config, conectar_bd


def hu00_despliegue_ambiental() -> tuple:
    """
    Retorna (out_config, out_system_exception).
    out_config contiene todos los parametros necesarios para HU01 y HU02.
    """
    out_config = {}
    out_system_exception = ""
    task_name = "HU00_DespliegueAmbiental"

    try:
        # ----------------------------------------------------------------
        # PASO 1: Credenciales desde .env + esquema
        # ----------------------------------------------------------------
        cfg_base = obtener_config()
        out_config.update(cfg_base)           # _db, _correo, Scheme
        esquema       = out_config.get("Scheme", "[ShoppingDePrecios]")
        tabla_params  = "[Parametros]"

        # ----------------------------------------------------------------
        # PASO 2: Leer parametros NOMBRADOS desde la BD
        #         Solo filas donde Nombre parece un identificador valido
        # ----------------------------------------------------------------
        write_log("Info", "HU00: Comienza la HU00", task_name, out_config)

        conn   = conectar_bd(out_config)
        cursor = conn.cursor()
        cursor.execute(f"""
            SELECT Nombre, Valor
            FROM {esquema}.{tabla_params}
            WHERE Nombre IS NOT NULL
              AND LEN(LTRIM(RTRIM(Nombre))) BETWEEN 3 AND 60
              AND ISNUMERIC(Nombre) = 0
              AND Nombre NOT LIKE 'http%'
              AND Nombre NOT LIKE '[%'
              AND Nombre NOT LIKE '\\%'
              AND Nombre NOT LIKE '%/%'
              AND Nombre NOT LIKE '% %'
        """)
        for row in cursor.fetchall():
            if row[0] and row[1] is not None:
                out_config[row[0]] = row[1]
        conn.close()

        write_log("Info", "HU00: Se cargaron los parametros desde BD", task_name, out_config)

        # ----------------------------------------------------------------
        # PASO 3: Defaults de seguridad — solo aplican si BD no tiene el parametro
        # ----------------------------------------------------------------
        out_config["TablaParametros"] = tabla_params  # constante de arranque

        out_config.setdefault("NombreIniciativaLocatel", "Shopping de precios Locatel")
        out_config.setdefault("DrogueriaLocatel",        "LOCATEL")
        out_config.setdefault("TablaLocatel",            "[Locatel]")
        out_config.setdefault("LoteLocatel",             "50")
        out_config.setdefault("ArchivoInsumo",           "InsumoPricing.xlsx")
        out_config.setdefault("ArchivoEnvioCorreos",     "EnvioCorreos.xlsx")
        out_config.setdefault("SheetTicketInsumo",       "TicketInsumo")
        out_config.setdefault("SheetEnvioCorreos",       "EnvioCorreos")
        out_config.setdefault("TablaTicketInsumo",       "[TicketInsumo]")
        out_config.setdefault("TablaEnvioCorreos",       "[EnvioCorreos]")
        out_config.setdefault("CodigoRobot",             "SPIDER")
        out_config.setdefault("ActivarLog",              "True")
        out_config.setdefault("ReintentosHu",              "3")
        out_config.setdefault("ReintentosReprocesamiento", "0")
        out_config.setdefault("DiasCaidaDb",               "365")
        out_config.setdefault("MeseslimpiezaLog",           "12")
        out_config.setdefault("MesesLimpiezaScreenshots",  "12")
        out_config.setdefault("MesesLimpiezaReportes",     "12")
        out_config.setdefault("MesesLimpiezaInsumos",      "12")
        out_config.setdefault("NombreResultado",           "ReportePricing")
        out_config.setdefault("NombreHojaResultado",       "ReportePricingLOCATEL")

        # ----------------------------------------------------------------
        # PASO 4: Validacion de carpetas (crear si no existen)
        # ----------------------------------------------------------------
        carpetas = [
            out_config.get("RutaInsumos", ""),
            out_config.get("PathLog", ""),
            out_config.get("RutaTemp", ""),
            out_config.get("RutaReporte", ""),
            out_config.get("RutaScreenshots", ""),
        ]
        for carpeta in carpetas:
            if carpeta and not os.path.isdir(carpeta):
                os.makedirs(carpeta, exist_ok=True)

        write_log("Info", "HU00: Se realizo validacion de carpetas", task_name, out_config)

        # ----------------------------------------------------------------
        # PASO 5: Limpieza de Logs
        # ----------------------------------------------------------------
        meses_log = int(out_config.get("MeseslimpiezaLog", "12"))
        fecha_corte_log = datetime.now() - relativedelta(months=meses_log)
        path_log = out_config["PathLog"]

        if os.path.isdir(path_log):
            for archivo in Path(path_log).glob("*.txt"):
                try:
                    if datetime.fromtimestamp(archivo.stat().st_mtime) < fecha_corte_log:
                        archivo.unlink()
                except Exception:
                    pass

        write_log("Info", f"HU00: Limpieza de Logs — {meses_log} meses", task_name, out_config)

        # ----------------------------------------------------------------
        # PASO 6: Limpieza de Screenshots
        # ----------------------------------------------------------------
        meses_ss = int(out_config.get("MesesLimpiezaScreenshots", "3"))
        fecha_corte_ss = datetime.now() - relativedelta(months=meses_ss)

        if os.path.isdir(out_config["RutaScreenshots"]):
            for carpeta_anio in Path(out_config["RutaScreenshots"]).iterdir():
                if carpeta_anio.is_dir():
                    for carpeta_mes in carpeta_anio.iterdir():
                        if carpeta_mes.is_dir():
                            try:
                                if datetime.fromtimestamp(carpeta_mes.stat().st_mtime) < fecha_corte_ss:
                                    shutil.rmtree(carpeta_mes, ignore_errors=True)
                            except Exception:
                                pass

        write_log("Info", f"HU00: Limpieza de Screenshots — {meses_ss} meses", task_name, out_config)

        # ----------------------------------------------------------------
        # PASO 7: Limpieza de Reportes
        # ----------------------------------------------------------------
        meses_rep = int(out_config.get("MesesLimpiezaReportes", "6"))
        fecha_corte_rep = datetime.now() - relativedelta(months=meses_rep)

        if os.path.isdir(out_config["RutaReporte"]):
            for d in Path(out_config["RutaReporte"]).iterdir():
                if d.is_dir():
                    try:
                        if datetime.fromtimestamp(d.stat().st_mtime) < fecha_corte_rep:
                            shutil.rmtree(d, ignore_errors=True)
                    except Exception:
                        pass

        write_log("Info", f"HU00: Limpieza de Reportes — {meses_rep} meses", task_name, out_config)

        # ----------------------------------------------------------------
        # PASO 8: Limpieza de Insumos procesados
        # ----------------------------------------------------------------
        meses_ins = int(out_config.get("MesesLimpiezaInsumos", "6"))
        fecha_corte_ins = datetime.now() - relativedelta(months=meses_ins)
        ruta_proc = out_config.get("RutaProcesados", "")

        if os.path.isdir(ruta_proc):
            for archivo in Path(ruta_proc).iterdir():
                if archivo.is_file():
                    try:
                        if datetime.fromtimestamp(archivo.stat().st_mtime) < fecha_corte_ins:
                            archivo.unlink()
                    except Exception:
                        pass

        write_log("Info", f"HU00: Limpieza de Insumos procesados — {meses_ins} meses", task_name, out_config)

        # ----------------------------------------------------------------
        # PASO 9: Limpieza de BD
        # ----------------------------------------------------------------
        dias_caida = out_config.get("DiasCaidaDb", "365")
        conn   = conectar_bd(out_config)
        cursor = conn.cursor()
        cursor.execute(
            f"DELETE FROM {esquema}.[TicketInsumo] "
            f"WHERE FechaInicio < DATEADD(DAY, -{dias_caida}, GETDATE())"
        )
        cursor.execute(
            f"DELETE FROM {esquema}.{out_config.get('TablaLocatel', '[Locatel]')} "
            f"WHERE FechaInicio < DATEADD(DAY, -{dias_caida}, GETDATE())"
        )
        cursor.execute(
            f"UPDATE {esquema}.{tabla_params} "
            f"SET Valor = CAST(GETDATE() AS DATE) WHERE Nombre = 'LimpiezaDB'"
        )
        conn.commit()
        conn.close()

        write_log("Info", f"HU00: Limpieza de BD — {dias_caida} dias", task_name, out_config)
        write_log("Info", "HU00: Finaliza la HU00", task_name, out_config)

    except Exception as e:
        out_system_exception = str(e)
        write_log("Error", f"HU00: {e}", task_name, out_config)
        write_log("Info", "HU00: Finaliza la HU00", task_name, out_config)

    return out_config, out_system_exception


if __name__ == "__main__":
    config, excepcion = hu00_despliegue_ambiental()
    if excepcion:
        print(f"ERROR: {excepcion}")
        sys.exit(1)
    print("HU00 completada exitosamente.")
