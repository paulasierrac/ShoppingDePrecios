"""
================================================================================
HU00 - Despliegue Ambiental
Nombre de la iniciativa: Shopping de Precios Cafam
Autor: Paula Sierra — Net Applications
Descripcion: Conecta a BD, carga todos los parametros nombrados desde la tabla
             [ShoppingDePrecios].[Parametros] y realiza limpieza de
             Logs, Screenshots, Reportes, Insumos y BD.
             Las rutas vienen directamente de la tabla Parametros.
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

_KEY_URL           = "UrlCafam"
_NOMBRE_INICIATIVA = "Shopping de precios Cafam"
_DROGUERIA         = "CAFAM"


def hu00_despliegue_ambiental() -> tuple:
    out_config = {}
    out_system_exception = ""
    task_name = "HU00_DespliegueAmbiental"

    try:
        cfg_base = obtener_config()
        out_config.update(cfg_base)
        out_config["_debug"] = os.environ.get("RPA_DEBUG", "").lower() == "true"
        esquema      = out_config["Scheme"]
        tabla_params = "[Parametros]"

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

        # URL de busqueda Cafam: UrlCafam ya termina en q=, solo agregar EAN
        url_base = out_config[_KEY_URL]
        out_config["UrlCafam"] = url_base + "REEMPLAZAR"

        out_config["NombreIniciativaCafam"] = _NOMBRE_INICIATIVA
        out_config["DrogueriaCafam"]        = _DROGUERIA
        out_config["NombreResultado"]       = "ReportePricingCafam_"
        out_config["NombreHojaResultado"]   = "ReportePricingCafam"
        out_config.setdefault("HeadlessChrome", "true")

        carpetas = [
            out_config.get("RutaInsumos")     or "",
            out_config.get("PathLog")         or "",
            out_config.get("RutaTemp")        or "",
            out_config.get("RutaScreenshots") or "",
            out_config.get("RutaReporte")     or "",
        ]
        for carpeta in carpetas:
            if carpeta and not os.path.isdir(carpeta):
                try:
                    os.makedirs(carpeta, exist_ok=True)
                except Exception:
                    pass

        write_log("Info", "HU00: Se realizo validacion de carpetas", task_name, out_config)

        meses_log       = int(out_config["MeseslimpiezaLog"])
        fecha_corte_log = datetime.now() - relativedelta(months=meses_log)
        path_log        = out_config.get("PathLog") or ""
        if os.path.isdir(path_log):
            for archivo in Path(path_log).glob("*.txt"):
                try:
                    if datetime.fromtimestamp(archivo.stat().st_mtime) < fecha_corte_log:
                        archivo.unlink()
                except Exception:
                    pass
        write_log("Info", f"HU00: Limpieza de Logs — {meses_log} meses", task_name, out_config)

        meses_ss       = int(out_config["MesesLimpiezaScreenshots"])
        fecha_corte_ss = datetime.now() - relativedelta(months=meses_ss)
        ruta_ss        = out_config.get("RutaScreenshots") or ""
        if os.path.isdir(ruta_ss):
            for ca in Path(ruta_ss).iterdir():
                if ca.is_dir():
                    for cm in ca.iterdir():
                        if cm.is_dir():
                            try:
                                if datetime.fromtimestamp(cm.stat().st_mtime) < fecha_corte_ss:
                                    shutil.rmtree(cm, ignore_errors=True)
                            except Exception:
                                pass
        write_log("Info", f"HU00: Limpieza de Screenshots — {meses_ss} meses", task_name, out_config)

        meses_rep       = int(out_config["MesesLimpiezaReportes"])
        fecha_corte_rep = datetime.now() - relativedelta(months=meses_rep)
        ruta_rep        = out_config.get("RutaReporte") or ""
        if os.path.isdir(ruta_rep):
            for d in Path(ruta_rep).iterdir():
                if d.is_dir():
                    try:
                        if datetime.fromtimestamp(d.stat().st_mtime) < fecha_corte_rep:
                            shutil.rmtree(d, ignore_errors=True)
                    except Exception:
                        pass
        write_log("Info", f"HU00: Limpieza de Reportes — {meses_rep} meses", task_name, out_config)

        meses_ins       = int(out_config["MesesLimpiezaInsumos"])
        fecha_corte_ins = datetime.now() - relativedelta(months=meses_ins)
        ruta_proc       = out_config.get("RutaProcesados") or ""
        if os.path.isdir(ruta_proc):
            for archivo in Path(ruta_proc).iterdir():
                if archivo.is_file():
                    try:
                        if datetime.fromtimestamp(archivo.stat().st_mtime) < fecha_corte_ins:
                            archivo.unlink()
                    except Exception:
                        pass
        write_log("Info", f"HU00: Limpieza de Insumos procesados — {meses_ins} meses", task_name, out_config)

        fecha_hoy    = datetime.now().strftime("%Y-%m-%d")
        ultima_limpi = str(out_config.get("LimpiezaDB") or "").strip()[:10]
        if ultima_limpi != fecha_hoy:
            dias_caida = out_config["DiasCaidaDb"]
            tabla_ex   = out_config["TablaCafam"]
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
    for k in ("RutaInsumos", "PathLog", "RutaScreenshots", "RutaReporte", "UrlCafam"):
        print(f"  {k}: {config.get(k, '(no encontrado)')}")
