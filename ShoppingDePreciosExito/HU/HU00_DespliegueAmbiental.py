"""
================================================================================
HU00 - Despliegue Ambiental
Nombre de la iniciativa: Shopping de Precios Exito
Autor: KPMG Advisory, Tax & Legal SAS
Descripcion: Conecta a BD, carga todos los parametros nombrados desde la tabla
             [ShoppingDePrecios].[Parametros] y realiza limpieza de
             Logs, Screenshots, Reportes, Insumos y BD.
             Las rutas de trabajo vienen directamente de la tabla Parametros
             (RutaRed, RutaInsumos, PathLog, RutaScreenshots, RutaReporte, etc.)
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

# ── Constantes propias de Exito ────────────────────────────────────────────
_KEY_URL           = "UrlExito"
_NOMBRE_INICIATIVA = "Shopping de precios Exito"
_DROGUERIA         = "EXITO"


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
        # PASO 1: Credenciales desde Azure Key Vault + esquema base
        # ----------------------------------------------------------------
        cfg_base = obtener_config()
        out_config.update(cfg_base)           # _db, _correo, Scheme
        esquema      = out_config.get("Scheme", "[ShoppingDePrecios]")
        tabla_params = "[Parametros]"

        # ----------------------------------------------------------------
        # PASO 2: Leer TODOS los parametros nombrados desde la BD
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
        # PASO 3: Ajustes puntuales — rutas y URL de busqueda
        #         Las rutas principales (RutaRed, RutaInsumos, PathLog,
        #         RutaScreenshots, RutaReporte, RutaTemp) ya estan en
        #         out_config cargadas desde Parametros.
        # ----------------------------------------------------------------

        # URL de busqueda Exito: {base}/s?q=<EAN>&sort=score_desc&page=0
        url_base = out_config.get(_KEY_URL, "https://www.exito.com/").rstrip("/")
        out_config["UrlExito"] = f"{url_base}/s?q=REEMPLAZAR&sort=score_desc&page=0"

        # Garantizar defaults para valores que pueden no estar en la tabla
        out_config.setdefault("NombreIniciativaExito", _NOMBRE_INICIATIVA)
        out_config.setdefault("DrogueriaExito",        _DROGUERIA)
        out_config.setdefault("NombreHojaResultado",   "ReportePricingExito")
        out_config.setdefault("ReintentosHu",              "3")
        out_config.setdefault("ReintentosReprocesamiento", "0")
        out_config.setdefault("DiasCaidaDb",               "365")
        out_config.setdefault("MeseslimpiezaLog",           "12")
        out_config.setdefault("MesesLimpiezaScreenshots",  "12")
        out_config.setdefault("MesesLimpiezaReportes",     "12")
        out_config.setdefault("MesesLimpiezaInsumos",      "12")
        out_config.setdefault("NombreResultado",           "ReportePricing")
        out_config.setdefault("CantExito",                 "100")
        out_config.setdefault("SegExito",                  "10")

        # ----------------------------------------------------------------
        # PASO 4: Validacion de carpetas (crear si no existen)
        # ----------------------------------------------------------------
        carpetas = [
            out_config.get("RutaInsumos",    ""),
            out_config.get("PathLog",        ""),
            out_config.get("RutaTemp",       ""),
            out_config.get("RutaScreenshots",""),
            out_config.get("RutaReporte",    ""),
        ]
        for carpeta in carpetas:
            if carpeta and not os.path.isdir(carpeta):
                try:
                    os.makedirs(carpeta, exist_ok=True)
                except Exception:
                    pass

        write_log("Info", "HU00: Se realizo validacion de carpetas", task_name, out_config)

        # ----------------------------------------------------------------
        # PASO 5: Limpieza de Logs
        # ----------------------------------------------------------------
        meses_log       = int(out_config.get("MeseslimpiezaLog", "12"))
        fecha_corte_log = datetime.now() - relativedelta(months=meses_log)
        path_log        = out_config.get("PathLog", "")

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
        meses_ss       = int(out_config.get("MesesLimpiezaScreenshots", "12"))
        fecha_corte_ss = datetime.now() - relativedelta(months=meses_ss)
        ruta_ss        = out_config.get("RutaScreenshots", "")

        if os.path.isdir(ruta_ss):
            for carpeta_anio in Path(ruta_ss).iterdir():
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
        meses_rep       = int(out_config.get("MesesLimpiezaReportes", "12"))
        fecha_corte_rep = datetime.now() - relativedelta(months=meses_rep)
        ruta_rep        = out_config.get("RutaReporte", "")

        if os.path.isdir(ruta_rep):
            for d in Path(ruta_rep).iterdir():
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
        meses_ins       = int(out_config.get("MesesLimpiezaInsumos", "12"))
        fecha_corte_ins = datetime.now() - relativedelta(months=meses_ins)
        ruta_ins        = out_config.get("RutaInsumos", "")
        carpeta_proc    = out_config.get("CarpetaProcesados", "Procesados\\")
        ruta_proc       = os.path.join(ruta_ins, carpeta_proc)

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
        #         Se ejecuta solo una vez por dia (verifica LimpiezaDB)
        # ----------------------------------------------------------------
        fecha_hoy    = datetime.now().strftime("%Y-%m-%d")
        ultima_limpi = str(out_config.get("LimpiezaDB", "")).strip()[:10]

        if ultima_limpi != fecha_hoy:
            dias_caida = out_config.get("DiasCaidaDb", "365")
            tabla_ex   = out_config.get("TablaExito", "[Exito]")

            conn   = conectar_bd(out_config)
            cursor = conn.cursor()
            cursor.execute(
                f"DELETE FROM {esquema}.[TicketInsumo] "
                f"WHERE FechaInicio < DATEADD(DAY, -{dias_caida}, GETDATE())"
            )
            cursor.execute(
                f"DELETE FROM {esquema}.{tabla_ex} "
                f"WHERE FechaInicio < DATEADD(DAY, -{dias_caida}, GETDATE())"
            )
            cursor.execute(
                f"UPDATE {esquema}.{tabla_params} "
                f"SET Valor = CAST(GETDATE() AS DATE) WHERE Nombre = 'LimpiezaDB'"
            )
            conn.commit()
            conn.close()
            write_log("Info", f"HU00: Limpieza de BD — {dias_caida} dias", task_name, out_config)
        else:
            write_log("Info", "HU00: Limpieza de BD ya realizada hoy, se omite", task_name, out_config)

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
    # Mostrar rutas cargadas desde BD para verificacion
    for k in ("RutaRed", "RutaInsumos", "PathLog", "RutaScreenshots", "RutaReporte", "UrlExito"):
        print(f"  {k}: {config.get(k, '(no encontrado)')}")
