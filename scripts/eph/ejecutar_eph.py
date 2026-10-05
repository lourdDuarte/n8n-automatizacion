from datetime import datetime
from pathlib import Path
import subprocess
import sys


ARCHIVO_ACTUAL = Path(__file__).resolve()
BASE_DIR = ARCHIVO_ACTUAL.parents[2]

CARPETA_SCRIPTS = BASE_DIR / "scripts" / "eph"
CARPETA_LOGS = BASE_DIR / "logs" / "eph"

SCRIPT_DESCARGA = CARPETA_SCRIPTS / "descargar_eph.py"
SCRIPT_EXTRACCION = CARPETA_SCRIPTS / "extraer_eph.py"

CARPETA_LOGS.mkdir(parents=True, exist_ok=True)


def ejecutar_script(nombre_etapa: str, ruta_script: Path) -> dict:
    print("=" * 72)
    print(f"ETAPA: {nombre_etapa}")
    print("=" * 72)

    if not ruta_script.exists():
        return {
            "etapa": nombre_etapa,
            "codigo": 1,
            "estado": "ERROR",
            "salida": f"No existe el archivo: {ruta_script}",
        }

    resultado = subprocess.run(
        [sys.executable, "-u", str(ruta_script)],
        cwd=str(BASE_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    salida = resultado.stdout or ""
    error = resultado.stderr or ""

    if salida:
        print(salida, end="" if salida.endswith("\n") else "\n")

    if error:
        print("SALIDA DE ERROR:")
        print(error, end="" if error.endswith("\n") else "\n")

    texto_completo = f"{salida}\n{error}"

    if resultado.returncode != 0 or "ESTADO=ERROR" in texto_completo:
        estado = "ERROR"
    elif "ESTADO=PENDIENTE" in texto_completo:
        estado = "PENDIENTE"
    elif "ESTADO=OK" in texto_completo:
        estado = "OK"
    else:
        estado = "DESCONOCIDO"

    return {
        "etapa": nombre_etapa,
        "codigo": resultado.returncode,
        "estado": estado,
        "salida": texto_completo.strip(),
    }


def guardar_log(resultados: list[dict], estado_final: str) -> Path:
    momento = datetime.now()
    nombre = momento.strftime("eph_%Y%m%d_%H%M%S.log")
    ruta_log = CARPETA_LOGS / nombre

    lineas = [
        "=" * 72,
        "AUTOMATIZACIÓN EPH - GRAN FORMOSA",
        "=" * 72,
        f"Fecha y hora: {momento.strftime('%Y-%m-%d %H:%M:%S')}",
        f"Python: {sys.executable}",
        f"Proyecto: {BASE_DIR}",
        f"Estado final: {estado_final}",
        "",
    ]

    for resultado in resultados:
        lineas.extend(
            [
                "-" * 72,
                f"ETAPA: {resultado['etapa']}",
                f"CÓDIGO: {resultado['codigo']}",
                f"ESTADO: {resultado['estado']}",
                "-" * 72,
                resultado["salida"],
                "",
            ]
        )

    ruta_log.write_text(
        "\n".join(lineas),
        encoding="utf-8",
    )

    return ruta_log


def main() -> int:
    print("=" * 72)
    print("PROCESO GENERAL — EPH GRAN FORMOSA")
    print("=" * 72)
    print(f"Inicio: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Python utilizado: {sys.executable}")
    print(f"Proyecto: {BASE_DIR}")

    resultados = []

    descarga = ejecutar_script(
        "DESCARGA",
        SCRIPT_DESCARGA,
    )
    resultados.append(descarga)

    if descarga["estado"] == "ERROR":
        estado_final = "ERROR"
        ruta_log = guardar_log(resultados, estado_final)

        print(f"LOG_EPH={ruta_log}")
        print("ESTADO_FINAL=ERROR")
        return 1

    if descarga["estado"] == "PENDIENTE":
        estado_final = "PENDIENTE"
        ruta_log = guardar_log(resultados, estado_final)

        print("-" * 72)
        print("La publicación todavía no está disponible.")
        print("No se ejecutarán las etapas posteriores.")
        print(f"LOG_EPH={ruta_log}")
        print("ESTADO_FINAL=PENDIENTE")
        print("=" * 72)
        return 0

    extraccion = ejecutar_script(
        "EXTRACCIÓN E INVENTARIO",
        SCRIPT_EXTRACCION,
    )
    resultados.append(extraccion)

    if extraccion["estado"] == "ERROR":
        estado_final = "ERROR"
        ruta_log = guardar_log(resultados, estado_final)

        print(f"LOG_EPH={ruta_log}")
        print("ESTADO_FINAL=ERROR")
        return 1

    if extraccion["estado"] == "PENDIENTE":
        estado_final = "PENDIENTE"
        ruta_log = guardar_log(resultados, estado_final)

        print(f"LOG_EPH={ruta_log}")
        print("ESTADO_FINAL=PENDIENTE")
        print("=" * 72)
        return 0

    estado_final = "INVENTARIO_LISTO"
    ruta_log = guardar_log(resultados, estado_final)

    print("-" * 72)
    print("El ZIP fue descargado y sus archivos fueron identificados.")
    print("Falta validar columnas y completar el cálculo de las tasas.")
    print("El Excel no fue modificado.")
    print(f"LOG_EPH={ruta_log}")
    print("ESTADO_FINAL=INVENTARIO_LISTO")
    print("=" * 72)

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nEjecución cancelada por el usuario.")
        print("ESTADO_FINAL=CANCELADO")
        sys.exit(130)
    except Exception as error:
        print("=" * 72)
        print(f"ERROR NO CONTROLADO: {error}")
        print("ESTADO_FINAL=ERROR")
        print("=" * 72)
        sys.exit(1)