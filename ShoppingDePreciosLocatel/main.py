"""
================================================================================
Main - Shopping de Precios Locatel
Nombre de la iniciativa: Shopping de Precios Locatel
Autor: KPMG Advisory, Tax & Legal SAS
Descripcion: Orquesta secuencialmente las tareas del proceso de shopping
             de precios para Locatel.
             Equivale al bot Main_ShoppingDePreciosLocatel de Automation Anywhere.
Ultima modificacion: 27/05/2025
Propiedad de Colsubsidio
================================================================================

Flujo:
  1. Ejecuta HU00_DespliegueAmbiental (hasta 3 intentos) para obtener el
     diccionario de configuracion desde la BD.
  2. Carga la tabla de envio de correos (CargarTablaEnvioCorreos).
  3. Envia correo de inicio del proceso (codEmail=0).
  4. Bucle principal:
       a. Ejecuta HU01 (hasta ReintentosHu veces); si persiste el error, aborta.
       b. Ejecuta HU02 (hasta ReintentosHu veces); si tiene exito, termina el bucle.
          Si persiste el error, aborta.
  5. En caso de exito: envia correo de finalizacion (codEmail=100).
  6. En caso de error: envia correo de error (codEmail=99).
"""

import sys
from pathlib import Path

_PHARMACY_ROOT = Path(__file__).resolve().parent
_PROJECT_ROOT  = _PHARMACY_ROOT.parent
sys.path.insert(0, str(_PROJECT_ROOT))
sys.path.insert(0, str(_PHARMACY_ROOT))

from Funciones.utils import write_log, cargar_tabla_envio_correos, enviar_correo
from HU.HU00_DespliegueAmbiental    import hu00_despliegue_ambiental
from HU.HU01_ValidacionYCargaInsumo import hu01_validacion_y_carga_insumo
from HU.HU02_ConsultaYReporte       import hu02_consulta_y_reporte


TASK_NAME = "Main_ShoppingDePreciosLocatel"


def main() -> None:
    io_config          = {}
    p_system_exception = ""
    p_detener_bot      = False
    p_aux_string       = ""

    # ----------------------------------------------------------------
    # PASO 1: Despliegue ambiental (hasta 3 reintentos)
    # ----------------------------------------------------------------
    for _ in range(3):
        io_config, p_system_exception = hu00_despliegue_ambiental()
        if not p_system_exception:
            break

    if p_system_exception:
        write_log(
            "Error",
            f"Se presento el error ({p_system_exception}) en el despliegue ambiental",
            TASK_NAME, io_config
        )
        sys.exit(1)

    write_log("Info", "-" * 100, TASK_NAME, io_config)
    write_log(
        "Info",
        f"Inicia ejecucion del BOT ({io_config.get('NombreIniciativaLocatel', '')})",
        TASK_NAME, io_config
    )

    # ----------------------------------------------------------------
    # PASO 2: Cargar tabla de envio de correos
    # ----------------------------------------------------------------
    p_system_exception = cargar_tabla_envio_correos(io_config)
    if p_system_exception:
        write_log(
            "Error",
            f"Se presento el error ({p_system_exception}) en la funcion CargarTablaEnvioCorreos",
            TASK_NAME, io_config
        )
        p_system_exception = ""  # no critico — se continua

    try:
        # ----------------------------------------------------------------
        # PASO 3: Correo de inicio (codEmail=0)
        # ----------------------------------------------------------------
        p_aux_dic = {"$NombrePagina$": io_config.get("DrogueriaLocatel", "Locatel")}
        from_address = io_config.get("_correo", {}).get("usuario", "")

        err = enviar_correo(
            in_config=io_config,
            i_cod_email=0,
            i_from_address=from_address,
            i_replace_in_message=p_aux_dic,
            i_replace_in_subject=p_aux_dic,
            i_html_format=False
        )
        if err:
            write_log("Info", "No fue posible enviar el correo de notificacion", TASK_NAME, io_config)

        # ----------------------------------------------------------------
        # PASO 4: Bucle principal (HU01 + HU02) hasta exito o error definitivo
        # ----------------------------------------------------------------
        reintentos_hu = int(io_config.get("ReintentosHu", "3"))

        while not p_detener_bot:

            # ── HU01: Validacion y Carga de Insumo ───────────────────
            p_system_exception = ""
            for _ in range(reintentos_hu):
                p_system_exception = hu01_validacion_y_carga_insumo(io_config)
                if not p_system_exception:
                    break

            if p_system_exception:
                write_log(
                    "Info",
                    f"Existe error en la ejecucion de la HU01, se finaliza proceso: "
                    f"{p_system_exception}",
                    TASK_NAME, io_config
                )
                raise RuntimeError(p_system_exception)

            # ── HU02: Consulta y Reporte ─────────────────────────────
            p_system_exception = ""
            for _ in range(reintentos_hu):
                p_system_exception = hu02_consulta_y_reporte(io_config)
                if not p_system_exception:
                    break

            if p_system_exception:
                p_aux_string = p_system_exception
                write_log(
                    "Info",
                    f"Existe error en la ejecucion de la HU02, se finaliza proceso: "
                    f"{p_system_exception}",
                    TASK_NAME, io_config
                )
                raise RuntimeError(p_system_exception)
            else:
                p_detener_bot = True

        # ----------------------------------------------------------------
        # PASO 5: Correo de finalizacion exitosa (codEmail=100)
        # ----------------------------------------------------------------
        err = enviar_correo(
            in_config=io_config,
            i_cod_email=100,
            i_from_address=from_address,
            i_replace_in_message=p_aux_dic,
            i_replace_in_subject=p_aux_dic,
            i_html_format=False
        )
        if err:
            write_log("Info", "No fue posible enviar el correo de notificacion", TASK_NAME, io_config)

        write_log(
            "Info",
            f"Finaliza ejecucion del BOT ({io_config.get('NombreIniciativaLocatel', '')})",
            TASK_NAME, io_config
        )
        write_log("Info", "-" * 100, TASK_NAME, io_config)

    except Exception as e:
        p_system_exception = str(e)
        write_log("Error", p_system_exception, TASK_NAME, io_config)

        p_aux_dic_error = {
            "$NombrePagina$": io_config.get("DrogueriaLocatel", "Locatel"),
            "$Proceso$":      io_config.get("NombreIniciativaLocatel", ""),
            "$Error$":        p_aux_string or p_system_exception,
        }
        err = enviar_correo(
            in_config=io_config,
            i_cod_email=99,
            i_from_address=from_address,
            i_replace_in_message=p_aux_dic_error,
            i_replace_in_subject=p_aux_dic_error,
            i_html_format=False
        )
        if err:
            write_log("Info", "No fue posible enviar el correo de notificacion", TASK_NAME, io_config)

        write_log(
            "Info",
            f"Finaliza ejecucion del BOT ({io_config.get('NombreIniciativaLocatel', '')})",
            TASK_NAME, io_config
        )
        write_log("Info", "-" * 100, TASK_NAME, io_config)

        sys.exit(1)


if __name__ == "__main__":
    main()
