"""Ejecuta descarga, extracción y actualización del IPI para n8n."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import subprocess
import sys


RAIZ = Path(__file__).resolve().parent.parent
LOGS = RAIZ / "logs" / "ipi"
ETAPAS = (
    ("DESCARGA", "descargar_ipi.py"),
    ("EXTRACCION", "extraer_ipi.py"),
    ("ACTUALIZACION_EXCEL", "actualizar_base.py"),
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
    log = LOGS / f"ipi_{inicio:%Y%m%d_%H%M%S}.log"
    registros = ["# AUTOMATIZACIÓN — IPI MANUFACTURERO\n"]
    print("=" * 72)
    print("AUTOMATIZACIÓN COMPLETA — IPI MANUFACTURERO")
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
        salida = proceso.stdout
        error = proceso.stderr
        print(salida, end="")
        if error:
            print(error, end="", file=sys.stderr)
        registros.append(encabezado + "\n" + salida + error)
        if proceso.returncode != 0:
            mensaje = f"RESULTADO: ERROR — Falló {nombre} (código {proceso.returncode})."
            registros.append("\n" + mensaje + "\n")
            log.write_text("".join(registros), encoding="utf-8")
            print(mensaje, file=sys.stderr)
            print(f"LOG_IPI={log}", file=sys.stderr)
            return 1
    registros.append(f"\nRESULTADO: OK\nFin: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
    log.write_text("".join(registros), encoding="utf-8")
    print("RESULTADO: OK — PROCESO IPI COMPLETADO")
    print(f"LOG_IPI={log}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())