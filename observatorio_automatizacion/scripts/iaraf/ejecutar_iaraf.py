"""Ejecuta de punta a punta la actualización mensual de IARAF."""

from __future__ import annotations

import os
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path


RAIZ = Path(__file__).resolve().parents[2]
CARPETA_SCRIPTS = RAIZ / "scripts" / "iaraf"
CARPETA_LOGS = RAIZ / "logs" / "iaraf"
ETAPAS = (
    ("DESCARGA", "descargar_iaraf.py"),
    ("EXTRACCION", "extraer_iaraf.py"),
    ("ACTUALIZACION_EXCEL", "actualizar_base_iaraf.py"),
)


class ErrorEtapa(RuntimeError):
    """Conserva stdout/stderr de una etapa que terminó con error."""

    def __init__(self, mensaje: str, registro: str) -> None:
        super().__init__(mensaje)
        self.registro = registro


def configurar_salida() -> None:
    if os.name == "nt":
        os.system("")
    for flujo in (sys.stdout, sys.stderr):
        if hasattr(flujo, "reconfigure"):
            flujo.reconfigure(encoding="utf-8", errors="replace")


def ejecutar_etapa(nombre: str, archivo: str, argumentos: list[str] | None = None) -> str:
    ruta = CARPETA_SCRIPTS / archivo
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el script requerido: {ruta}")

    print(f"\n[{nombre}] {archivo}")
    print("-" * 72)

    comando = [sys.executable, str(ruta)]
    if argumentos:
        comando.extend(argumentos)

    resultado = subprocess.run(
        comando,
        cwd=RAIZ,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    salida = resultado.stdout or ""
    error = resultado.stderr or ""
    if salida:
        print(salida.rstrip())
    if error:
        print(error.rstrip(), file=sys.stderr)

    registro = (
        f"\n[{nombre}] {archivo}\n"
        f"CODIGO_SALIDA={resultado.returncode}\n"
        f"{salida}"
        f"{error}"
    )

    if resultado.returncode != 0:
        raise ErrorEtapa(
            f"Falló la etapa {nombre} (código {resultado.returncode}).",
            registro,
        )

    return registro

def main() -> None:
    configurar_salida()
    inicio = datetime.now()
    marca = inicio.strftime("%Y%m%d_%H%M%S")
    CARPETA_LOGS.mkdir(parents=True, exist_ok=True)
    ruta_log = CARPETA_LOGS / f"iaraf_{marca}.log"

    print("=" * 72)
    print("AUTOMATIZACIÓN COMPLETA — IARAF")
    print("=" * 72)
    print(f"Inicio: {inicio:%Y-%m-%d %H:%M:%S}")

    registros = []
    pdf_iaraf = None

    try:
        for nombre, archivo in ETAPAS:
            try:
                argumentos = None

                if nombre == "EXTRACCION" and pdf_iaraf:
                    argumentos = ["--pdf", pdf_iaraf]

                registro = ejecutar_etapa(nombre, archivo, argumentos)
                registros.append(registro)

                if nombre == "DESCARGA":
                    coincidencia = re.search(r"PDF_IARAF=(.+)", registro)
                    if coincidencia:
                        pdf_iaraf = coincidencia.group(1).strip()

            except ErrorEtapa as error:
                # Si la web de IARAF no deja consultar (p. ej. verificación anti-robots), se sigue
                # con el último PDF guardado en fuentes/iaraf (descargado antes o puesto a mano).
                if nombre != "DESCARGA" or not list((RAIZ / "fuentes" / "iaraf").glob("*.pdf")):
                    raise
                registros.append(error.registro)
                print("AVISO: no se pudo consultar la página de IARAF; se usa el último PDF guardado.")
                print(
                    "NOTA_TELEGRAM=La página de IARAF tiene una verificación anti-robots "
                    "y no se pudo consultar; se usó el último informe guardado. "
                    "Si salió uno nuevo, descargalo a mano y guardalo en la carpeta fuentes\\iaraf."
                )

    except Exception as error:
        if isinstance(error, ErrorEtapa):
            registros.append(error.registro)
        registros.append(f"\nRESULTADO: ERROR\n{error}\n")
        ruta_log.write_text("".join(registros), encoding="utf-8")
        print("=" * 72, file=sys.stderr)
        print(f"RESULTADO: ERROR — {error}", file=sys.stderr)
        print(f"LOG_IARAF={ruta_log}", file=sys.stderr)
        print("=" * 72, file=sys.stderr)
        raise SystemExit(1) from error

    fin = datetime.now()
    registros.append(
        f"\nRESULTADO: OK\nFin: {fin:%Y-%m-%d %H:%M:%S}\n"
    )
    ruta_log.write_text("".join(registros), encoding="utf-8")

    reportes = sorted(
        (RAIZ / "reportes" / "iaraf").glob("iaraf_*.json"),
        key=lambda ruta: ruta.stat().st_mtime,
        reverse=True,
    )

    if not reportes:
        raise SystemExit("La ejecución IARAF terminó sin generar su reporte JSON")

    resultado_fuente = json.loads(reportes[0].read_text(encoding="utf-8"))

    cantidades = [
        int(valor)
        for registro in registros
        for valor in re.findall(r"Celdas actualizadas:\s*(\d+)", registro)
    ]

    resultado_fuente["fuente"] = "IARAF"
    resultado_fuente["actualizacion"] = {
        "celdas_actualizadas": cantidades[-1] if cantidades else 0,
        "cambios": [],
        "detalle_completo": [],
    }

    print("\n" + "=" * 72)
    print("RESULTADO: OK — PROCESO IARAF COMPLETADO")
    print(f"Duración: {(fin - inicio).total_seconds():.1f} segundos")
    print(f"LOG_IARAF={ruta_log}")
    print(
        "RESULTADO_IARAF="
        + json.dumps(resultado_fuente, ensure_ascii=False, separators=(",", ":"))
    )
    print("=" * 72)
if __name__ == "__main__":
    main()
