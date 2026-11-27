"""Detecta y descarga automáticamente el informe mensual más reciente de IARAF."""

from collections import deque
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urljoin, urlparse
import hashlib
import re
import sys

import requests
from bs4 import BeautifulSoup
import fitz


CARPETA_SCRIPT = Path(__file__).resolve().parent
CARPETA_PROYECTO = CARPETA_SCRIPT.parent.parent
CARPETA_DESTINO = CARPETA_PROYECTO / "fuentes" / "iaraf"

DOMINIO_IARAF = "www.iaraf.org"
PAGINAS_INICIALES = [
    "https://www.iaraf.org/index.php",
    "https://www.iaraf.org/index.php/informes-economicos/area-fiscal",
]
PALABRAS_CLAVE = (
    "coparticipación",
    "coparticipacion",
    "transferencias automáticas",
    "transferencias automaticas",
)
MAXIMO_PAGINAS = 80
TIEMPO_ESPERA = 30
ENCABEZADOS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0 Safari/537.36"
    ),
    "Accept-Language": "es-AR,es;q=0.9",
}


def configurar_salida():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def normalizar(texto):
    tabla = str.maketrans("áéíóúüñ", "aeiouun")
    return " ".join((texto or "").lower().translate(tabla).split())


def contiene_palabra_clave(texto):
    limpio = normalizar(texto)
    return any(normalizar(palabra) in limpio for palabra in PALABRAS_CLAVE)


def es_url_iaraf(url):
    return urlparse(url).netloc.lower() in {DOMINIO_IARAF, "iaraf.org"}


def fecha_publicacion(url, texto):
    # Las notas de IARAF suelen terminar en DD-MM-AA.
    coincidencias = re.findall(r"(?<!\d)(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})(?!\d)", f"{url} {texto}")
    fechas = []
    for dia, mes, anio in coincidencias:
        anio = int(anio)
        if anio < 100:
            anio += 2000
        try:
            fechas.append(datetime(anio, int(mes), int(dia)))
        except ValueError:
            continue
    return max(fechas) if fechas else datetime.min


def obtener_html(sesion, url):
    respuesta = sesion.get(url, timeout=TIEMPO_ESPERA)
    respuesta.raise_for_status()
    return BeautifulSoup(respuesta.text, "html.parser")


def descubrir_notas(sesion):
    pendientes = deque(PAGINAS_INICIALES)
    visitadas = set()
    candidatas = {}

    while pendientes and len(visitadas) < MAXIMO_PAGINAS:
        url = pendientes.popleft()
        if url in visitadas:
            continue
        visitadas.add(url)
        try:
            sopa = obtener_html(sesion, url)
        except requests.RequestException as error:
            print(f"AVISO: no se pudo revisar {url}: {error}")
            continue

        texto_pagina = sopa.get_text(" ", strip=True)
        # Si ya estamos dentro de una nota, el título puede no estar en el
        # enlace que llevó hasta ella. La propia página y su botón de descarga
        # son evidencia suficiente para considerarla candidata.
        tiene_descarga = any(
            "drive.google.com" in urljoin(url, a.get("href", ""))
            or urljoin(url, a.get("href", "")).lower().split("?", 1)[0].endswith(".pdf")
            or "descargar" in normalizar(a.get_text(" ", strip=True))
            or "informe completo" in normalizar(a.get_text(" ", strip=True))
            for a in sopa.select("a[href]")
        )
        if contiene_palabra_clave(texto_pagina) and tiene_descarga:
            titulo = sopa.title.get_text(" ", strip=True) if sopa.title else texto_pagina[:300]
            candidatas[url] = titulo

        for enlace in sopa.select("a[href]"):
            destino = urljoin(url, enlace.get("href", "")).split("#", 1)[0]
            texto = enlace.get_text(" ", strip=True)
            contexto = enlace.parent.get_text(" ", strip=True) if enlace.parent else texto
            if es_url_iaraf(destino) and contiene_palabra_clave(f"{texto} {contexto}"):
                candidatas[destino] = contexto

            # Recorre páginas y notas del área fiscal aunque el botón diga
            # solamente "Leer más". La validación por palabras clave se hace
            # al abrir cada página, no por una clase CSS específica.
            ruta_destino = normalizar(urlparse(destino).path)
            consulta = urlparse(destino).query.casefold()
            es_area_fiscal = (
                es_url_iaraf(destino)
                and destino != url
                and (
                    "informes economicos area fiscal" in ruta_destino
                    or "area fiscal" in ruta_destino
                    or "start=" in consulta
                    or "limitstart=" in consulta
                )
            )
            if es_area_fiscal and destino not in visitadas and destino not in pendientes:
                pendientes.append(destino)

        # En portada, el enlace puede decir solo "Leer más". Se evalúa el bloque
        # contenedor completo para asociarlo con el título de transferencias.
        for bloque in sopa.select("article, .item, .blog-item, .item-content, .leading-0"):
            contexto = bloque.get_text(" ", strip=True)
            if not contiene_palabra_clave(contexto):
                continue
            for enlace in bloque.select("a[href]"):
                destino = urljoin(url, enlace.get("href", "")).split("#", 1)[0]
                if es_url_iaraf(destino):
                    candidatas[destino] = contexto

        # Respaldo específico para la portada Joomla: relaciona cada "Leer más"
        # con el encabezado inmediatamente anterior.
        for encabezado in sopa.select("h1, h2, h3, h4"):
            titulo = encabezado.get_text(" ", strip=True)
            if not contiene_palabra_clave(titulo):
                continue
            contenedor = encabezado.find_parent(["article", "div"])
            if contenedor:
                for enlace in contenedor.select("a[href]"):
                    destino = urljoin(url, enlace.get("href", "")).split("#", 1)[0]
                    if es_url_iaraf(destino):
                        candidatas[destino] = titulo

            # La portada nueva separa el encabezado y el botón "Leer más"
            # en contenedores hermanos. Buscar los enlaces inmediatamente
            # posteriores permite asociarlos sin depender de una clase CSS.
            for enlace in encabezado.find_all_next("a", href=True, limit=6):
                destino = urljoin(url, enlace.get("href", "")).split("#", 1)[0]
                texto_enlace = normalizar(enlace.get_text(" ", strip=True))
                ruta = urlparse(destino).path.lower()
                es_nota = (
                    destino != url
                    and es_url_iaraf(destino)
                    and (
                        "leer mas" in texto_enlace
                        or "informe-economico" in ruta
                        or "area-fiscal" in ruta
                    )
                )
                if es_nota:
                    candidatas[destino] = titulo
                    break

    if not candidatas:
        raise RuntimeError("No se encontraron publicaciones de transferencias automáticas en IARAF.")
    return sorted(candidatas.items(), key=lambda item: fecha_publicacion(item[0], item[1]), reverse=True)


def id_google_drive(url):
    parsed = urlparse(url)
    coincidencia = re.search(r"/file/d/([^/]+)", parsed.path)
    if coincidencia:
        return coincidencia.group(1)
    return parse_qs(parsed.query).get("id", [None])[0]


def localizar_descarga(sesion, url_nota):
    sopa = obtener_html(sesion, url_nota)
    enlaces = []
    for enlace in sopa.select("a[href]"):
        destino = urljoin(url_nota, enlace.get("href", ""))
        texto = normalizar(enlace.get_text(" ", strip=True))
        if "drive.google.com" in destino:
            archivo_id = id_google_drive(destino)
            if archivo_id:
                return f"https://drive.google.com/uc?export=download&id={archivo_id}"
        if destino.lower().split("?", 1)[0].endswith(".pdf"):
            enlaces.append(destino)
        elif "descargar" in texto or "informe completo" in texto:
            enlaces.append(destino)
    if enlaces:
        return enlaces[0]
    raise RuntimeError(f"La publicación no contiene un enlace descargable: {url_nota}")


def descargar_pdf(sesion, url):
    respuesta = sesion.get(url, timeout=TIEMPO_ESPERA, allow_redirects=True)
    respuesta.raise_for_status()
    contenido = respuesta.content

    # Google Drive puede pedir una confirmación para archivos grandes.
    if not contenido.startswith(b"%PDF") and "drive.google.com" in respuesta.url:
        sopa = BeautifulSoup(respuesta.text, "html.parser")
        formulario = sopa.select_one("form#download-form")
        if formulario:
            parametros = {
                campo.get("name"): campo.get("value", "")
                for campo in formulario.select("input[name]")
            }
            url_confirmacion = urljoin(respuesta.url, formulario.get("action", ""))
            respuesta = sesion.get(url_confirmacion, params=parametros, timeout=TIEMPO_ESPERA)
            respuesta.raise_for_status()
            contenido = respuesta.content

    if not contenido.startswith(b"%PDF"):
        raise RuntimeError("El enlace encontrado no devolvió un PDF válido.")
    return contenido

def validar_informe_mensual(contenido, fecha):
    """Verifica que el PDF sea el informe mensual de transferencias automáticas."""
    try:
        documento = fitz.open(stream=contenido, filetype="pdf")
        texto = " ".join(pagina.get_text("text") for pagina in documento)
        documento.close()
    except Exception as error:
        raise RuntimeError(f"No se pudo leer el PDF para validarlo: {error}")

    limpio = normalizar(texto)

    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
    ]
    mes = meses[fecha.month - 1]
    anio = fecha.year

    patron_periodo = rf"transferencias automaticas de {mes} de {anio}"

    if not re.search(patron_periodo, limpio):
        raise RuntimeError(
            f"El PDF no corresponde al informe mensual de transferencias "
            f"automáticas de {mes} de {anio}."
        )

    if "variacion nominal" not in limpio:
        raise RuntimeError(
            "El PDF no contiene la variación nominal esperada del informe mensual."
        )

    if "acumulado" not in limpio:
        raise RuntimeError(
            "El PDF no contiene información acumulada esperada del informe mensual."
        )

    return True


def guardar_pdf(contenido, fecha):
    CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
    nombre = f"iaraf_transferencias_{fecha:%Y_%m}.pdf"
    destino = CARPETA_DESTINO / nombre
    huella_nueva = hashlib.sha256(contenido).hexdigest()
    if destino.exists() and hashlib.sha256(destino.read_bytes()).hexdigest() == huella_nueva:
        print(f"Sin descarga: el informe ya existe y coincide: {destino}")
    else:
        destino.write_bytes(contenido)
        print(f"Informe descargado: {destino}")
    return destino


def main():
    configurar_salida()
    print("=" * 72)
    print("DESCARGA AUTOMÁTICA — IARAF")
    print("=" * 72)
    sesion = requests.Session()
    sesion.headers.update(ENCABEZADOS)

    publicaciones = descubrir_notas(sesion)
    ultimo_error = None
    for url_nota, titulo in publicaciones:
        try:
            fecha = fecha_publicacion(url_nota, titulo)
            if fecha == datetime.min:
                continue
            print(f"Publicación seleccionada: {titulo}")
            print(f"URL: {url_nota}")
            url_descarga = localizar_descarga(sesion, url_nota)
            contenido = descargar_pdf(sesion, url_descarga)
            validar_informe_mensual(contenido, fecha)
            destino = guardar_pdf(contenido, fecha)
            print(f"PDF_IARAF={destino}")
            print("RESULTADO: OK")
            print("=" * 72)
            return 0
        except (requests.RequestException, RuntimeError) as error:
            ultimo_error = error
            print(f"AVISO: publicación descartada: {error}")

    raise RuntimeError(f"No se pudo descargar ningún informe candidato. Último error: {ultimo_error}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
