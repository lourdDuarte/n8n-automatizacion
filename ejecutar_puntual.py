"""Ejecuta una sola fuente y conserva la salida compatible con n8n."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


RAIZ = Path(__file__).resolve().parent

RUTAS = {
    "dnrpa": "ejecutar_general.py",
    "ieric": "ejecutar_general.py",
    "ipc": "ejecutar_general.py",
    "iaraf": "ejecutar_general.py",
    "ipi": "ejecutar_general.py",
    "canasta": "ejecutar_general.py",
    "sipa": "ejecutar_general.py",
    "emae": "ejecutar_complejos.py",
    "salarios": "ejecutar_complejos.py",
    "icc": "ejecutar_complejos.py",
    "supermercados": "ejecutar_complejos.py",
    "comercializacion": "ejecutar_productivos.py",
}


def main() -> int:
    parser = argparse.ArgumentParser(description="Ejecuta un indicador puntual")
    parser.add_argument("--fuente", required=True)
    parser.add_argument("--listar", action="store_true")
    args = parser.parse_args()

    if args.listar:
        print("\n".join(RUTAS))
        return 0

    fuente = args.fuente.strip().casefold()
    ejecutor = RUTAS.get(fuente)
    if ejecutor is None:
        print("ERROR=Fuente desconocida: " + fuente)
        print("DISPONIBLES=" + ",".join(RUTAS))
        return 2

    entorno = os.environ.copy()
    entorno["PYTHONIOENCODING"] = "utf-8"
    entorno["PYTHONUTF8"] = "1"
    proceso = subprocess.run(
        [sys.executable, "-u", str(RAIZ / ejecutor), "--fuentes", fuente, "--estricto"],
        cwd=RAIZ,
        env=entorno,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proceso.stdout:
        print(proceso.stdout, end="" if proceso.stdout.endswith("\n") else "\n")
    if proceso.stderr:
        print(proceso.stderr, file=sys.stderr, end="" if proceso.stderr.endswith("\n") else "\n")
    return proceso.returncode


if __name__ == "__main__":
    raise SystemExit(main())
