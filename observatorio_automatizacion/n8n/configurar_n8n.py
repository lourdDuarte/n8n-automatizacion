"""Genera los workflows de n8n listos para importar en el servidor.

Toma las plantillas exportadas (carpeta ``plantillas``) y reemplaza las rutas
de la PC original por las del servidor. En Linux además traduce los comandos
de PowerShell a comandos de shell.

Uso (desde esta carpeta):

    python3 configurar_n8n.py --so linux \
        --raiz /opt/observatorio/automatizacion_indicadores \
        --python /opt/observatorio/venv/bin/python \
        --n8n-files /opt/observatorio/n8n-files

    python configurar_n8n.py --so windows \
        --raiz "D:\\observatorio\\automatizacion_indicadores" \
        --python "C:\\Python314\\python.exe" \
        --n8n-files "C:\\Users\\servidor\\.n8n-files"

Los archivos generados quedan en ``importar_<so>/``.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

AQUI = Path(__file__).resolve().parent
PLANTILLAS = AQUI / "plantillas"

# Rutas de la PC donde se desarrolló la automatización.
RAIZ_ORIGINAL = r"C:\Users\daeco\Desktop\automatizacion_indicadores"
N8N_FILES_ORIGINAL = r"C:\Users\daeco\.n8n-files"
PYTHON_ORIGINAL = r"C:\Users\daeco\AppData\Local\Python\pythoncore-3.14-64\python.exe"

# Para cada nodo "Execute Command": script de Python y expresión de argumentos.
COMANDOS = {
    "Ejecutar indicadores generales": (
        "ejecutar_general.py",
        "resumen_general.json",
        "--fuentes '{{ $json.fuente === 'general' || $json.fuente === 'generales' ? 'todas' : ($json.fuente || 'todas') }}'"
        " --estricto {{ $json.fuente ? '--forzar-google' : '' }}",
    ),
    "Ejecutar indicadores complejos": (
        "ejecutar_complejos.py",
        "resumen_complejos.json",
        "--fuentes '{{ $json.fuente === 'complejo' || $json.fuente === 'complejos' ? 'todas' : ($json.fuente || 'todas') }}'"
        " --estricto",
    ),
    "Ejecutar indicadores productivos": (
        "ejecutar_productivos.py",
        "resumen_productivos.json",
        "--fuentes '{{ $json.fuente === 'productivo' || $json.fuente === 'productivos' ? 'todas' : ($json.fuente || 'todas') }}'"
        " --estricto",
    ),
}


def comando_linux(raiz: str, python: str, n8n_files: str, script: str, resumen: str, argumentos: str) -> str:
    salida = f"{raiz}/salidas_n8n/{resumen}"
    copia = f"{n8n_files}/{resumen}"
    return (
        f"=cd '{raiz}' && rm -f '{salida}' '{copia}'; "
        f"'{python}' -u '{script}' {argumentos}; "
        f"if [ -f '{salida}' ]; then mkdir -p '{n8n_files}' && cp -f '{salida}' '{copia}'; exit 0; fi; exit 1"
    )


def comando_windows(raiz: str, python: str, n8n_files: str, script: str, resumen: str, argumentos: str) -> str:
    salida = f"{raiz}\\salidas_n8n\\{resumen}"
    copia = f"{n8n_files}\\{resumen}"
    return (
        "=powershell.exe -NoProfile -ExecutionPolicy Bypass -Command \"$ErrorActionPreference='Stop'; "
        f"Set-Location '{raiz}'; Remove-Item -LiteralPath '{salida}','{copia}' -ErrorAction SilentlyContinue; "
        f"& '{python}' -u '.\\{script}' {argumentos}; "
        f"if (Test-Path -LiteralPath '{salida}') {{ Copy-Item -LiteralPath '{salida}' -Destination '{copia}' -Force; exit 0 }}; exit 1\""
    )


def configurar(so: str, raiz: str, python: str, n8n_files: str) -> Path:
    raiz = raiz.rstrip("/\\")
    n8n_files = n8n_files.rstrip("/\\")
    sep = "/" if so == "linux" else "\\"
    destino = AQUI / f"importar_{so}"
    destino.mkdir(exist_ok=True)

    for plantilla in sorted(PLANTILLAS.glob("*.json")):
        flujo = json.loads(plantilla.read_text(encoding="utf-8"))
        for nodo in flujo["nodes"]:
            p = nodo.get("parameters", {})
            nombre = nodo["name"]

            if nombre in COMANDOS:
                script, resumen, argumentos = COMANDOS[nombre]
                armar = comando_linux if so == "linux" else comando_windows
                p["command"] = armar(raiz, python, n8n_files, script, resumen, argumentos)

            elif nodo["type"].endswith("readWriteFile") and "fileSelector" in p:
                archivo = p["fileSelector"].replace(N8N_FILES_ORIGINAL + "\\", "")
                p["fileSelector"] = f"{n8n_files}{sep}{archivo}"

            elif nombre == "Seleccionar diagnóstico":
                vieja = "const base = 'C:\\\\Users\\\\daeco\\\\Desktop\\\\automatizacion_indicadores\\\\salidas_n8n\\\\';"
                if vieja not in p.get("jsCode", ""):
                    raise SystemExit("No se encontró la ruta base en 'Seleccionar diagnóstico'.")
                ruta = f"{raiz}{sep}salidas_n8n{sep}"
                p["jsCode"] = p["jsCode"].replace(vieja, "const base = " + json.dumps(ruta) + ";")

            elif nombre == "Leer diagnóstico del indicador" and so == "linux":
                p["command"] = (
                    "={{ $json.ruta_resumen ? 'cat \"' + $json.ruta_resumen + '\" 2>/dev/null || echo \"{}\"' : 'echo \"{}\"' }}"
                )

        texto = json.dumps(flujo, ensure_ascii=False, indent=2)
        for viejo in (RAIZ_ORIGINAL, N8N_FILES_ORIGINAL, PYTHON_ORIGINAL):
            if viejo.replace("\\", "\\\\") in texto:
                raise SystemExit(f"Quedó una ruta de la PC original en {plantilla.name}: {viejo}")
        (destino / plantilla.name).write_text(texto, encoding="utf-8")
        print(f"OK  {destino.name}/{plantilla.name}")
    return destino


def main() -> int:
    ap = argparse.ArgumentParser(description="Genera los workflows de n8n con las rutas del servidor")
    ap.add_argument("--so", choices=("linux", "windows"), required=True)
    ap.add_argument("--raiz", required=True, help="Carpeta donde se copió automatizacion_indicadores")
    ap.add_argument("--python", required=True, help="Ejecutable de Python (idealmente el del entorno virtual)")
    ap.add_argument("--n8n-files", required=True, help="Carpeta que n8n tiene permitida para leer archivos")
    a = ap.parse_args()
    configurar(a.so, a.raiz, a.python, a.n8n_files)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
