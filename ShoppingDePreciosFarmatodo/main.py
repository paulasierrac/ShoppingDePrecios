"""
================================================================================
Main - Shopping de Precios Farmatodo
Nombre de la iniciativa: Shopping de Precios Farmatodo
Autor: KPMG Advisory, Tax & Legal SAS
Descripcion: Orquesta secuencialmente las tareas del proceso de shopping
             de precios para Farmatodo.
             Equivale al bot Main_ShoppingDePreciosFarmatodo de Automation Anywhere.
Ultima modificacion: 22/06/2026
Propiedad de Colsubsidio
================================================================================

Modo debug: set RPA_DEBUG=true  (sin escrituras en BD ni correos, Chrome visible)
"""

import os
import sys
from pathlib import Path

_DEBUG = os.environ.get("RPA_DEBUG", "").lower() in ("1", "true", "si", "yes")

_PHARMACY_ROOT = Path(__file__).resolve().parent
_PROJECT_ROOT  = _PHARMACY_ROOT.parent
sys.path.insert(0, str(_PROJECT_ROOT))
sys.path.insert(0, str(_PHARMACY_ROOT))

from Funciones.utils import write_log, cargar_tabla_envio_correos, enviar_correo
from HU.HU00_DespliegueAmbiental    import hu00_despliegue_ambiental
from HU.HU01_ValidacionYCargaInsumo import hu01_validacion_y_carga_insumo
from HU.HU02_ConsultaYReporte       import hu02_consulta_y_reporte

TASK_NAME = "Main_ShoppingDePreciosFarmatodo"


def main() -> None:
    io_config          = {}
    p_system_exception = ""
    p_detener_bot      = False
    p_aux_string       = ""

    for _ in range(3):
        io_config, p_system_exception = hu00_despliegue_ambiental()
        if not p_system_exception:
            break

    if p_system_exception:
        write_log("Error",
                  f"Se presento el error ({p_system_exception}) en el despliegue ambiental",
                  TASK_NAME, io_config)
        sys.exit(1)

    io_config["_debug"] = _DEBUG
    if _DEBUG:
        io_config["HeadlessChrome"] = "false"
        print("[DEBUG] Modo desarrollo activo — sin escrituras en BD ni envio de correos")

    write_log("Info", "-" * 100, TASK_NAME, io_config)
    write_log("Info",
              f"Inicia ejecucion del BOT ({io_config.get('NombreIniciativaFarmatodo', '')})",
              TASK_NAME, io_config)

    p_system_exception = cargar_tabla_envio_correos(io_config)
    if p_system_exception:
        write_log("Error",
                  f"Se presento el error ({p_system_exception}) en la funcion CargarTablaEnvioCorreos",
                  TASK_NAME, io_config)
        p_system_exception = ""

    try:
        p_aux_dic    = {"$NombrePagina$": io_config.get("DrogueriaFarmatodo", "Farmatodo")}
        from_address = io_config.get("_correo", {}).get("usuario", "")

        if not _DEBUG:
            err = enviar_correo(in_config=io_config, i_cod_email=0,
                                i_from_address=from_address,
                                i_replace_in_message=p_aux_dic,
                                i_replace_in_subject=p_aux_dic,
                                i_html_format=False)
            if err:
                write_log("Info", "No fue posible enviar el correo de notificacion", TASK_NAME, io_config)
        else:
            write_log("Info", "[DEBUG] Correo de inicio omitido", TASK_NAME, io_config)

        reintentos_hu = int(io_config.get("ReintentosHu", "3"))

        while not p_detener_bot:
            p_system_exception = ""
            for _ in range(reintentos_hu):
                p_system_exception = hu01_validacion_y_carga_insumo(io_config)
                if not p_system_exception:
                    break
            if p_system_exception:
                write_log("Info",
                          f"Existe error en HU01, se finaliza proceso: {p_system_exception}",
                          TASK_NAME, io_config)
                raise RuntimeError(p_system_exception)

            p_system_exception = ""
            for _ in range(reintentos_hu):
                p_system_exception = hu02_consulta_y_reporte(io_config)
                if not p_system_exception:
                    break
            if p_system_exception:
                p_aux_string = p_system_exception
                write_log("Info",
                          f"Existe error en HU02, se finaliza proceso: {p_system_exception}",
                          TASK_NAME, io_config)
                raise RuntimeError(p_system_exception)
            else:
                p_detener_bot = True

        if not _DEBUG:
            err = enviar_correo(in_config=io_config, i_cod_email=100,
                                i_from_address=from_address,
                                i_replace_in_message=p_aux_dic,
                                i_replace_in_subject=p_aux_dic,
                                i_html_format=False)
            if err:
                write_log("Info", "No fue posible enviar el correo de notificacion", TASK_NAME, io_config)
        else:
            write_log("Info", "[DEBUG] Correo de finalizacion omitido", TASK_NAME, io_config)

        write_log("Info",
                  f"Finaliza ejecucion del BOT ({io_config.get('NombreIniciativaFarmatodo', '')})",
                  TASK_NAME, io_config)
        write_log("Info", "-" * 100, TASK_NAME, io_config)

    except Exception as e:
        p_system_exception = str(e)
        write_log("Error", p_system_exception, TASK_NAME, io_config)

        p_aux_dic_error = {
            "$NombrePagina$": io_config.get("DrogueriaFarmatodo", "Farmatodo"),
            "$Proceso$":      io_config.get("NombreIniciativaFarmatodo", ""),
            "$Error$":        p_aux_string or p_system_exception,
        }
        if not _DEBUG:
            err = enviar_correo(in_config=io_config, i_cod_email=99,
                                i_from_address=from_address,
                                i_replace_in_message=p_aux_dic_error,
                                i_replace_in_subject=p_aux_dic_error,
                                i_html_format=False)
            if err:
                write_log("Info", "No fue posible enviar el correo de notificacion", TASK_NAME, io_config)
        else:
            write_log("Info", "[DEBUG] Correo de error omitido", TASK_NAME, io_config)

        write_log("Info",
                  f"Finaliza ejecucion del BOT ({io_config.get('NombreIniciativaFarmatodo', '')})",
                  TASK_NAME, io_config)
        write_log("Info", "-" * 100, TASK_NAME, io_config)
        sys.exit(1)


if __name__ == "__main__":
    main()
