from pathlib import Path
import json
import shutil
import sys
import zipfile


# ---------------------------------------------------------------------
# RUTAS DEL PROYECTO
# ---------------------------------------------------------------------

ARCHIVO_ACTUAL = Path(__file__).resolve()
BASE_DIR = ARCHIVO_ACTUAL.parents[2]

CARPETA_FUENTES = BASE_DIR / "fuentes" / "eph"
CARPETA_EXTRAIDOS = CARPETA_FUENTES / "extraidos"
CARPETA_REPORTES = BASE_DIR / "reportes" / "eph"

ARCHIVO_INVENTARIO = CARPETA_REPORTES / "inventario_eph_2026_T1.json"

CARPETA_FUENTES.mkdir(parents=True, exist_ok=True)
CARPETA_EXTRAIDOS.mkdir(parents=True, exist_ok=True)
CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------
# FUNCIONES GENERALES
# ---------------------------------------------------------------------

def imprimir_encabezado():
    print("=" * 72)
    print("EXTRACCIÓN E INVENTARIO — EPH INDEC")
    print("=" * 72)
    print("Período esperado: primer trimestre de 2026")
    print(f"Carpeta fuente: {CARPETA_FUENTES}")
    print(f"Carpeta de extracción: {CARPETA_EXTRAIDOS}")
    print("-" * 72)


def es_zip_valido(ruta: Path) -> bool:
    if not ruta.exists() or ruta.stat().st_size == 0:
        return False

    try:
        with zipfile.ZipFile(ruta, "r") as archivo_zip:
            return archivo_zip.testzip() is None
    except zipfile.BadZipFile:
        return False


def localizar_zip() -> Path | None:
    candidatos = []

    for archivo in CARPETA_FUENTES.glob("*.zip"):
        if es_zip_valido(archivo):
            candidatos.append(archivo)

    if not candidatos:
        return None

    candidatos.sort(
        key=lambda archivo: archivo.stat().st_mtime,
        reverse=True,
    )

    return candidatos[0]


def ruta_segura(destino_base: Path, nombre_interno: str) -> Path:
    """
    Evita que un archivo ZIP pueda escribir fuera de la carpeta autorizada.
    """
    destino_base_resuelto = destino_base.resolve()
    destino = (destino_base / nombre_interno).resolve()

    if (
        destino != destino_base_resuelto
        and destino_base_resuelto not in destino.parents
    ):
        raise ValueError(
            f"Ruta insegura detectada dentro del ZIP: {nombre_interno}"
        )

    return destino


def extraer_zip(archivo_zip: Path) -> list[Path]:
    archivos_extraidos = []

    with zipfile.ZipFile(archivo_zip, "r") as comprimido:
        miembros = [
            miembro
            for miembro in comprimido.infolist()
            if not miembro.is_dir()
        ]

        if not miembros:
            raise ValueError("El ZIP no contiene archivos.")

        for miembro in miembros:
            destino = ruta_segura(
                CARPETA_EXTRAIDOS,
                miembro.filename,
            )

            destino.parent.mkdir(parents=True, exist_ok=True)

            with comprimido.open(miembro, "r") as origen:
                with destino.open("wb") as salida:
                    shutil.copyfileobj(origen, salida)

            archivos_extraidos.append(destino)

    return archivos_extraidos


# ---------------------------------------------------------------------
# IDENTIFICACIÓN DE ARCHIVOS
# ---------------------------------------------------------------------

def normalizar_nombre(ruta: Path) -> str:
    return ruta.name.lower().replace("-", "_").replace(" ", "_")


def es_archivo_tabular(ruta: Path) -> bool:
    return ruta.suffix.lower() in {
        ".xls",
        ".xlsx",
        ".csv",
        ".txt",
    }


def clasificar_archivo(ruta: Path) -> str:
    nombre = normalizar_nombre(ruta)

    referencias_individual = [
        "individual",
        "usu_individual",
        "usu_ind",
        "personas",
        "persona",
    ]

    referencias_hogar = [
        "hogar",
        "usu_hogar",
        "usu_hog",
        "hogares",
    ]

    referencias_diseno = [
        "diseno",
        "diseño",
        "registro",
        "diccionario",
        "documentacion",
        "documentación",
        "manual",
    ]

    if any(referencia in nombre for referencia in referencias_individual):
        return "base_individual"

    if any(referencia in nombre for referencia in referencias_hogar):
        return "base_hogar"

    if any(referencia in nombre for referencia in referencias_diseno):
        return "documentacion"

    if es_archivo_tabular(ruta):
        return "archivo_tabular_no_clasificado"

    return "otro"


def construir_inventario(
    archivo_zip: Path,
    archivos: list[Path],
) -> dict:
    inventario_archivos = []

    for ruta in sorted(
        archivos,
        key=lambda elemento: str(elemento).lower(),
    ):
        inventario_archivos.append(
            {
                "nombre": ruta.name,
                "ruta": str(ruta),
                "extension": ruta.suffix.lower(),
                "tamano_bytes": ruta.stat().st_size,
                "clasificacion": clasificar_archivo(ruta),
            }
        )

    bases_individuales = [
        elemento
        for elemento in inventario_archivos
        if elemento["clasificacion"] == "base_individual"
    ]

    bases_hogares = [
        elemento
        for elemento in inventario_archivos
        if elemento["clasificacion"] == "base_hogar"
    ]

    return {
        "periodo_esperado": {
            "anio": 2026,
            "trimestre": 1,
        },
        "fuente": "INDEC - Encuesta Permanente de Hogares",
        "archivo_zip": str(archivo_zip),
        "cantidad_archivos": len(inventario_archivos),
        "base_individual_detectada": (
            bases_individuales[0]["ruta"]
            if len(bases_individuales) == 1
            else None
        ),
        "base_hogar_detectada": (
            bases_hogares[0]["ruta"]
            if len(bases_hogares) == 1
            else None
        ),
        "cantidad_bases_individuales_detectadas": len(
            bases_individuales
        ),
        "cantidad_bases_hogares_detectadas": len(
            bases_hogares
        ),
        "archivos": inventario_archivos,
    }


def guardar_inventario(inventario: dict):
    with ARCHIVO_INVENTARIO.open(
        "w",
        encoding="utf-8",
    ) as archivo:
        json.dump(
            inventario,
            archivo,
            ensure_ascii=False,
            indent=2,
        )


def mostrar_resultado(inventario: dict):
    print(f"Archivos encontrados: {inventario['cantidad_archivos']}")
    print("-" * 72)

    for elemento in inventario["archivos"]:
        print(
            f"{elemento['clasificacion']}: "
            f"{elemento['nombre']} "
            f"({elemento['tamano_bytes']} bytes)"
        )

    print("-" * 72)

    individual = inventario["base_individual_detectada"]
    hogar = inventario["base_hogar_detectada"]

    if individual:
        print(f"BASE_INDIVIDUAL={individual}")
    else:
        print("BASE_INDIVIDUAL=NO_IDENTIFICADA_AUTOMATICAMENTE")

    if hogar:
        print(f"BASE_HOGAR={hogar}")
    else:
        print("BASE_HOGAR=NO_IDENTIFICADA_AUTOMATICAMENTE")

    print(f"INVENTARIO={ARCHIVO_INVENTARIO}")


# ---------------------------------------------------------------------
# EJECUCIÓN
# ---------------------------------------------------------------------

def main() -> int:
    imprimir_encabezado()

    archivo_zip = localizar_zip()

    if archivo_zip is None:
        print(
            "No se encontró todavía un ZIP válido de la EPH "
            "en la carpeta fuente."
        )
        print("No se extrajo ni modificó ningún archivo.")
        print("ESTADO=PENDIENTE")
        print("=" * 72)
        return 0

    print(f"ZIP localizado: {archivo_zip.name}")
    print(f"Tamaño: {archivo_zip.stat().st_size} bytes")
    print("-" * 72)

    archivos_extraidos = extraer_zip(archivo_zip)

    inventario = construir_inventario(
        archivo_zip,
        archivos_extraidos,
    )

    guardar_inventario(inventario)
    mostrar_resultado(inventario)

    print("-" * 72)
    print("Extracción e inventario completados.")
    print("ESTADO=OK")
    print("=" * 72)

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nEjecución cancelada por el usuario.")
        sys.exit(130)
    except Exception as error:
        print("=" * 72)
        print(f"ERROR NO CONTROLADO: {error}")
        print("ESTADO=ERROR")
        print("=" * 72)
        sys.exit(1)