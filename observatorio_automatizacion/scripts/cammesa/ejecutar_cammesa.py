"""Ejecuta descarga, extracción y actualización de CAMMESA para n8n."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import subprocess
import sys


RAIZ = Path(__file__).resolve().parent.parent
LOGS = RAIZ / "logs" / "cammesa"
ETAPAS = (
    ("DESCARGA", "descargar_cammesa.py"),
    ("EXTRACCION", "extraer_cammesa.py"),
    ("ACTUALIZACION_EXCEL", "actualizar_base_cammesa.py"),
)


def configurar_salida() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def main() -> int:
    configurar_salida()
    inicio = datetime.now()
    LOGS.mkdir(parents=True, exist_ok=True)
    log = LOGS / f"cammesa_{inicio:%Y%m%d_%H%M%S}.log"
    registros = ["# AUTOMATIZACIÓN — CAMMESA DEMANDA FORMOSA\n"]
    print("=" * 72)
    print("AUTOMATIZACIÓN COMPLETA — CAMMESA")
    print("=" * 72)
    for nombre, archivo in ETAPAS:
        encabezado = f"\n[{nombre}] {archivo}\n" + "-" * 72
        print(encabezado)
        proceso = subprocess.run(
            [sys.executable, "-u", str(Path(__file__).parent / archivo)],
            cwd=RAIZ,
            text=True,
            capture_output=True,
            encoding="utf-8",
            errors="replace",
        )
        print(proceso.stdout, end="")
        if proceso.stderr:
            print(proceso.stderr, end="", file=sys.stderr)
        registros.append(encabezado + "\n" + proceso.stdout + proceso.stderr)
        if proceso.returncode != 0:
            mensaje = f"RESULTADO: ERROR — Falló {nombre} (código {proceso.returncode})."
            registros.append("\n" + mensaje + "\n")
            log.write_text("".join(registros), encoding="utf-8")
            print(mensaje, file=sys.stderr)
            print(f"LOG_CAMMESA={log}", file=sys.stderr)
            return 1
    registros.append(f"\nRESULTADO: OK\nFin: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
    log.write_text("".join(registros), encoding="utf-8")
    print("RESULTADO: OK — PROCESO CAMMESA COMPLETADO")
    print(f"LOG_CAMMESA={log}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
