"""
================================================================================
Main Orquestador - Shopping de Precios (todas las farmacias)
Autor: Paula Sierra (NetApplications)
Descripcion: Ejecuta el main.py de cada farmacia de forma consecutiva.
             Un fallo en una farmacia se registra pero NO detiene las demas.
Propiedad de Colsubsidio
================================================================================
"""

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

_ROOT = Path(__file__).resolve().parent

# Orden de ejecucion de las farmacias
FARMACIAS = [
    "ShoppingDePreciosLocatel",
    "ShoppingDePreciosCafam",
    "ShoppingDePreciosComfandi",
    "ShoppingDePreciosCruzVerde",
    "ShoppingDePreciosExito",
    "ShoppingDePreciosFarmatodo",
    "ShoppingDePreciosLaRebaja",
    "ShoppingDePreciosMedipiel",
    "ShoppingDePreciosOlimpica",
    "ShoppingDePreciosOrtopedicos",
    "ShoppingDePreciosPasteur",
    "ShoppingDePreciosAlemana",
]


def ejecutar_farmacia(nombre: str) -> tuple[bool, float]:
    """
    Lanza ShoppingDePreciosXxx/main.py como subproceso.
    Retorna (exitoso, segundos_transcurridos).
    """
    script = _ROOT / nombre / "main.py"
    if not script.exists():
        print(f"  [SKIP] No existe {script.relative_to(_ROOT)}")
        return False, 0.0

    inicio = time.time()
    resultado = subprocess.run(
        [sys.executable, str(script)],
        cwd=str(_ROOT),
    )
    duracion = time.time() - inicio
    return resultado.returncode == 0, duracion


def main() -> None:
    inicio_total = datetime.now()
    print("=" * 70)
    print(f"  Shopping de Precios — inicio: {inicio_total:%d/%m/%Y %H:%M:%S}")
    print("=" * 70)

    resumen = []

    for farmacia in FARMACIAS:
        print(f"\n>>> {farmacia}")
        print("-" * 70)
        exitoso, seg = ejecutar_farmacia(farmacia)
        estado = "OK" if exitoso else "ERROR"
        resumen.append((farmacia, estado, seg))
        print(f"    {estado}  ({seg:.1f}s)")

    # ── Resumen final ──────────────────────────────────────────────────────
    fin_total = datetime.now()
    total_seg = (fin_total - inicio_total).total_seconds()

    print("\n" + "=" * 70)
    print(f"  Resumen — fin: {fin_total:%d/%m/%Y %H:%M:%S}  (total: {total_seg:.0f}s)")
    print("=" * 70)
    print(f"  {'Farmacia':<40} {'Estado':<8} {'Tiempo':>8}")
    print(f"  {'-'*40} {'-'*7} {'-'*8}")
    for farmacia, estado, seg in resumen:
        print(f"  {farmacia:<40} {estado:<8} {seg:>7.1f}s")

    errores = [f for f, e, _ in resumen if e == "ERROR"]
    if errores:
        print(f"\n  Farmacias con error: {', '.join(errores)}")
        sys.exit(1)
    else:
        print("\n  Todas las farmacias completadas exitosamente.")


if __name__ == "__main__":
    main()
