from pathlib import Path
from urllib.parse import urljoin
import re
import sys
import zipfile

import requests
from bs4 import BeautifulSoup


# ---------------------------------------------------------------------
# RUTAS DEL PROYECTO
# ---------------------------------------------------------------------

ARCHIVO_ACTUAL = Path(__file__).resolve()
BASE_DIR = ARCHIVO_ACTUAL.parents[2]

CARPETA_FUENTES = BASE_DIR / "fuentes" / "eph"
CARPETA_EXTRAIDOS = CARPETA_FUENTES / "extraidos"

CARPETA_FUENTES.mkdir(parents=True, exist_ok=True)
CARPETA_EXTRAIDOS.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------
# CONFIGURACIÓN DE LA PUBLICACIÓN
# ---------------------------------------------------------------------

ANIO = 2026
TRIMESTRE = 1

PAGINA_BASES = (
    "https://www.indec.gob.ar/indec/web/"
    "Institucional-Indec-BasesDeDatos"
)

CARPETA_FTP_EPH = (
    "https://www.indec.gob.ar/ftp/cuadros/menusuperior/eph/"
)

# El primer nombre es el formato que INDEC suele utilizar.
# Los demás quedan como alternativas por si modifica mayúsculas
# o la forma de escribir el trimestre.
NOMBRES_POSIBLES = [
    "EPH_usu_1_Trim_2026_xls.zip",
    "EPH_usu_1_Trim_2026_XLS.zip",
    "EPH_usu_1er_Trim_2026_xls.zip",
    "EPH_usu_1T2026_xls.zip",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/124 Safari/537.36"
    )
}

TIMEOUT = 45


# ---------------------------------------------------------------------
# FUNCIONES
# ---------------------------------------------------------------------

def imprimir_encabezado():
    print("=" * 72)
    print("DESCARGA AUTOMÁTICA — EPH INDEC")
    print("=" * 72)
    print(f"Período buscado: {ANIO} - trimestre {TRIMESTRE}")
    print(f"Carpeta de destino: {CARPETA_FUENTES}")
    print("-" * 72)


def es_zip_valido(ruta: Path) -> bool:
    if not ruta.exists() or ruta.stat().st_size == 0:
        return False

    try:
        with zipfile.ZipFile(ruta, "r") as archivo_zip:
            return archivo_zip.testzip() is None
    except zipfile.BadZipFile:
        return False


def buscar_en_pagina() -> list[str]:
    urls_encontradas = []

    try:
        respuesta = requests.get(
            PAGINA_BASES,
            headers=HEADERS,
            timeout=TIMEOUT,
        )
        respuesta.raise_for_status()

        soup = BeautifulSoup(respuesta.text, "html.parser")

        for enlace in soup.find_all("a", href=True):
            href = enlace.get("href", "").strip()
            texto = enlace.get_text(" ", strip=True)
            contenido = f"{href} {texto}".lower()

            es_zip = ".zip" in contenido
            es_eph = "eph" in contenido or "usu" in contenido
            es_anio = str(ANIO) in contenido

            referencias_trimestre = [
                "1_trim",
                "1 trim",
                "1er",
                "1t2026",
                "1t 2026",
                "primer trimestre",
            ]

            es_trimestre = any(
                referencia in contenido
                for referencia in referencias_trimestre
            )

            es_excel = (
                "xls" in contenido
                or "excel" in contenido
            )

            if es_zip and es_eph and es_anio and es_trimestre and es_excel:
                url_completa = urljoin(PAGINA_BASES, href)
                urls_encontradas.append(url_completa)

    except requests.RequestException as error:
        print(f"Aviso: no se pudo leer la página general: {error}")

    return urls_encontradas


def armar_urls_posibles() -> list[str]:
    urls = buscar_en_pagina()

    for nombre in NOMBRES_POSIBLES:
        urls.append(urljoin(CARPETA_FTP_EPH, nombre))

    # Elimina direcciones repetidas conservando el orden.
    return list(dict.fromkeys(urls))


def nombre_desde_url(url: str) -> str:
    nombre = url.split("?")[0].rstrip("/").split("/")[-1]

    if not nombre.lower().endswith(".zip"):
        nombre = f"EPH_usu_1_Trim_{ANIO}_xls.zip"

    nombre = re.sub(r'[<>:"/\\|?*]', "_", nombre)
    return nombre


def descargar_zip(url: str) -> Path | None:
    nombre_archivo = nombre_desde_url(url)
    destino = CARPETA_FUENTES / nombre_archivo
    temporal = destino.with_suffix(".zip.part")

    if es_zip_valido(destino):
        print(f"Sin descarga: el ZIP ya existe y es válido: {destino}")
        return destino

    try:
        print(f"Probando: {url}")

        respuesta = requests.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
            stream=True,
            allow_redirects=True,
        )

        if respuesta.status_code == 404:
            return None

        respuesta.raise_for_status()

        tipo_contenido = respuesta.headers.get(
            "Content-Type", ""
        ).lower()

        if "text/html" in tipo_contenido:
            return None

        with temporal.open("wb") as archivo:
            for bloque in respuesta.iter_content(chunk_size=1024 * 1024):
                if bloque:
                    archivo.write(bloque)

        if not es_zip_valido(temporal):
            temporal.unlink(missing_ok=True)
            return None

        temporal.replace(destino)
        return destino

    except requests.RequestException:
        temporal.unlink(missing_ok=True)
        return None

    except OSError as error:
        temporal.unlink(missing_ok=True)
        print(f"Error al guardar el archivo: {error}")
        return None


def buscar_zip_ya_descargado() -> Path | None:
    patrones = [
        f"*1*Trim*{ANIO}*xls*.zip",
        f"*1*T*{ANIO}*xls*.zip",
        f"*EPH*{ANIO}*.zip",
    ]

    for patron in patrones:
        for archivo in CARPETA_FUENTES.glob(patron):
            if es_zip_valido(archivo):
                return archivo

    return None


def main() -> int:
    imprimir_encabezado()

    archivo_existente = buscar_zip_ya_descargado()

    if archivo_existente:
        print(f"ZIP válido ya disponible: {archivo_existente}")
        print(f"EPH_ZIP={archivo_existente}")
        print("ESTADO=OK")
        print("=" * 72)
        return 0

    urls_posibles = armar_urls_posibles()

    for url in urls_posibles:
        archivo_descargado = descargar_zip(url)

        if archivo_descargado:
            print("-" * 72)
            print("Descarga completada y ZIP validado.")
            print(f"EPH_ZIP={archivo_descargado}")
            print("ESTADO=OK")
            print("=" * 72)
            return 0

    print("-" * 72)
    print(
        "La base EPH del primer trimestre de 2026 "
        "todavía no está disponible."
    )
    print("No se modificó ningún archivo del proyecto.")
    print("ESTADO=PENDIENTE")
    print("=" * 72)

    # No se considera error porque la publicación puede estar pendiente.
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