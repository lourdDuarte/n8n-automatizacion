"""Descubre y descarga publicaciones oficiales de actividad pecuaria."""

from __future__ import annotations

import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path

import fitz


PAGINAS = {
    "bovinos_estadistica": "https://www.magyp.gob.ar/sitio/areas/bovinos/informacion_sectorial/",
    "bovinos_informes": "https://www.magyp.gob.ar/sitio/areas/bovinos/informacion_interes/informes/index.php",
    "porcinos": "https://www.magyp.gob.ar/sitio/areas/porcinos/estadistica/informes.php?accion=imp",
    "aves_carne": "https://www.magyp.gob.ar/sitio/areas/aves/estadistica/carne/index.php",
    "aves_huevos": "https://www.magyp.gob.ar/sitio/areas/aves/estadistica/huevos/index.php",
}


@dataclass(frozen=True)
class Publicacion:
    clave: str
    url: str
    contenido: bytes
    nombre: str


def descargar(url: str, timeout: int = 90, intentos: int = 3) -> bytes:
    ultimo = None
    for intento in range(1, intentos + 1):
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
            "Accept-Language": "es-AR,es;q=0.9",
        })
        try:
            with urllib.request.urlopen(req, timeout=timeout) as respuesta:
                return respuesta.read()
        except (urllib.error.URLError, TimeoutError, OSError) as error:
            ultimo = error
            if intento < intentos:
                time.sleep(2 * intento)
    raise ConnectionError(f"No se pudo descargar {url}: {ultimo}")


def enlaces(html: str, pagina: str) -> list[tuple[str, str]]:
    resultado = []
    patron = re.compile(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', re.I | re.S)
    for href, cuerpo in patron.findall(html):
        texto = re.sub(r"<[^>]+>", " ", cuerpo)
        texto = " ".join(texto.replace("&nbsp;", " ").split())
        url_bruta = urllib.parse.urljoin(pagina, href)
        partes = urllib.parse.urlsplit(url_bruta)
        url = urllib.parse.urlunsplit((partes.scheme, partes.netloc, urllib.parse.quote(urllib.parse.unquote(partes.path), safe="/%=()"), partes.query, partes.fragment))
        if url.lower().split("?", 1)[0].endswith((".pdf", ".xls", ".xlsx")):
            resultado.append((texto, url))
    return resultado


def elegir(items: list[tuple[str, str]], *terminos: str, extension: str | None = None) -> str:
    candidatos = []
    for texto, url in items:
        combinado = urllib.parse.unquote(texto + " " + url).casefold()
        if all(t.casefold() in combinado for t in terminos):
            if extension and not url.lower().split("?", 1)[0].endswith(extension):
                continue
            candidatos.append(url)
    if not candidatos:
        raise ValueError(f"No se encontró publicación con términos {terminos}")
    def orden(url: str):
        anios = [int(x) for x in re.findall(r"20\d{2}", urllib.parse.unquote(url))]
        return (max(anios, default=0), url.lower().endswith(".xlsx"))
    return max(candidatos, key=orden)


def descubrir() -> dict[str, str]:
    paginas = {}
    todos = {}
    for clave, url in PAGINAS.items():
        html = descargar(url).decode("utf-8", errors="replace")
        paginas[clave] = enlaces(html, url)

    todos["stock_bovino"] = elegir(paginas["bovinos_informes"], "serie", "stock", "bovino", "provincia", extension=".xlsx")
    todos["faena_bovina"] = elegir(paginas["bovinos_informes"], "informe", "mensual", "faena", "bovina", extension=".pdf")
    todos["faena_bovina_excel"] = "https://www.magyp.gob.ar/sitio/areas/ss_ganaderia/archivos/Planilla_Informe_Faena_Bovina_.xls"
    tableros_bovinos = [
        (texto, url) for texto, url in paginas["bovinos_estadistica"]
        if "tablero de indicadores sectoriales bovino" in texto.casefold()
        and url.lower().split("?", 1)[0].endswith(".pdf")
    ]
    if not tableros_bovinos:
        raise ValueError("No se encontró el Tablero de indicadores sectoriales Bovino")
    todos["indicadores_bovinos_pdf"] = tableros_bovinos[0][1]
    todos["aves_oferta_demanda"] = elegir(paginas["aves_carne"], "indicadores", "oferta", "demanda", extension=".xlsx")
    huevos_validos = [(t, u) for t, u in paginas["aves_huevos"]
                      if "import" not in (t + " " + urllib.parse.unquote(u)).casefold()
                      and "export" not in (t + " " + urllib.parse.unquote(u)).casefold()]
    todos["aves_huevos"] = elegir(huevos_validos, "huevo", extension=".pdf")

    # El portal porcino cambia con más frecuencia. Se eligen los rótulos y no una URL anual fija.
    todos["stock_porcino"] = elegir(paginas["porcinos"], "existencias", "porcinas", "31-03", extension=".xlsx")
    todos["indicadores_porcinos"] = elegir(paginas["porcinos"], "evolucion", "mensual", "indicadores", extension=".xlsx")
    return todos


def enlace_excel_embebido(pdf: bytes, contiene: str) -> str:
    """Obtiene el XLS/XLSX oficial enlazado desde un tablero PDF."""
    with fitz.open(stream=pdf, filetype="pdf") as documento:
        urls = [
            enlace.get("uri", "")
            for pagina in documento
            for enlace in pagina.get_links()
            if enlace.get("uri")
        ]
    candidatas = [
        url for url in urls
        if contiene.casefold() in urllib.parse.unquote(url).casefold()
        and url.lower().split("?", 1)[0].endswith((".xls", ".xlsx"))
    ]
    if not candidatas:
        raise ValueError(f"El tablero no contiene el enlace Excel esperado: {contiene}")
    url = candidatas[0].replace("http://", "https://", 1)
    partes = urllib.parse.urlsplit(url)
    return urllib.parse.urlunsplit((
        partes.scheme,
        partes.netloc,
        urllib.parse.quote(urllib.parse.unquote(partes.path), safe="/%=()"),
        partes.query,
        partes.fragment,
    ))


def obtener(carpeta: Path) -> dict[str, Publicacion]:
    carpeta.mkdir(parents=True, exist_ok=True)
    salida = {}
    for clave, url in descubrir().items():
        contenido = descargar(url)
        nombre = urllib.parse.unquote(Path(urllib.parse.urlparse(url).path).name)
        ruta = carpeta / f"{clave}_{nombre}"
        ruta.write_bytes(contenido)
        salida[clave] = Publicacion(clave, url, contenido, ruta.name)
    url = enlace_excel_embebido(
        salida["indicadores_bovinos_pdf"].contenido,
        "Planilla_indicadores_bovinos",
    )
    contenido = descargar(url)
    nombre = urllib.parse.unquote(Path(urllib.parse.urlparse(url).path).name)
    ruta = carpeta / f"indicadores_bovinos_{nombre}"
    ruta.write_bytes(contenido)
    salida["indicadores_bovinos"] = Publicacion(
        "indicadores_bovinos", url, contenido, ruta.name
    )
    return salida
