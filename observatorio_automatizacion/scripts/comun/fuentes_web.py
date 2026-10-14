"""Búsqueda de archivos en las páginas oficiales, aunque el organismo cambie el nombre.

Cada programa prueba primero los nombres de archivo que conoce y, además, todos los
archivos que publica la página oficial del tema. Cada candidato se abre con el mismo
lector del indicador (que valida la estructura); se usa el que trae el período más
reciente. Un archivo que no corresponde falla la validación y simplemente se descarta.
"""
from __future__ import annotations

import re
import unicodedata
import urllib.parse
import urllib.request

INDEC = 'https://www.indec.gob.ar'
# Páginas de cada tema en el sitio del INDEC (listan los archivos vigentes).
PAGINAS_INDEC = {
    'ipc': INDEC + '/Nivel4/Tema/3/5/31',
    'emae': INDEC + '/Nivel4/Tema/3/9/48',
    'supermercados': INDEC + '/Nivel4/Tema/3/1/34',
    'salarios': INDEC + '/Nivel4/Tema/4/31/61',
    'pobreza': INDEC + '/Nivel4/Tema/4/46/152',
    'canasta': INDEC + '/Nivel4/Tema/4/43/149',
    'ipi': INDEC + '/Nivel4/Tema/3/6/14',
    'eph': INDEC + '/Nivel4/Tema/4/31/58',
    'pbi': INDEC + '/Nivel4/Tema/3/9/47',
}
CABECERAS = {'User-Agent': 'Mozilla/5.0 (Observatorio Formosa)'}


def pedir(url: str, timeout: int = 60) -> bytes:
    with urllib.request.urlopen(urllib.request.Request(url, headers=CABECERAS), timeout=timeout) as r:
        return r.read()


def _norm(texto: str) -> str:
    texto = unicodedata.normalize('NFKD', urllib.parse.unquote(texto))
    return ''.join(c for c in texto if not unicodedata.combining(c)).casefold()


def archivos_en_pagina(pagina: str, extensiones=('.xls', '.xlsx'), timeout: int = 30) -> list[str]:
    """Links (absolutos) a archivos con esas extensiones que muestra la página."""
    try:
        html = pedir(pagina, timeout).decode('utf-8', errors='replace')
    except Exception:  # noqa: BLE001 - sin página, se sigue con los nombres conocidos
        return []
    salida = []
    for href in re.findall(r'''href\s*=\s*["']([^"']+)["']''', html, re.I):
        url = urllib.parse.urljoin(pagina, href.replace('&amp;', '&'))
        ruta = urllib.parse.urlparse(url).path.lower()
        if ruta.endswith(tuple(extensiones)) and url not in salida:
            salida.append(url)
    return salida


def candidatos(preferidos=(), pagina: str | None = None, incluir=(), excluir=(),
               extensiones=('.xls', '.xlsx')) -> list[str]:
    """Primero los nombres conocidos; después los archivos de la página cuyo nombre
    contiene todas las palabras de `incluir` y ninguna de `excluir`."""
    lista = [u for u in preferidos if u]
    if pagina:
        for url in archivos_en_pagina(pagina, extensiones):
            nombre = _norm(url.rsplit('/', 1)[-1])
            if all(_norm(p) in nombre for p in incluir) and not any(_norm(p) in nombre for p in excluir):
                lista.append(url)
    vistos, unicos = set(), []
    for u in lista:
        if u not in vistos:
            vistos.add(u)
            unicos.append(u)
    return unicos


def elegir_mas_reciente(urls, procesar, timeout: int = 90):
    """Descarga cada candidato y lo procesa con procesar(contenido, url) -> (periodo, resultado).

    Devuelve (resultado, url, probados) del que trae el período más reciente; ante
    empate gana el primero de la lista (el nombre conocido).
    """
    mejor = None
    probados = []
    for url in urls:
        try:
            contenido = pedir(url, timeout)
            periodo, resultado = procesar(contenido, url)
        except Exception as error:  # noqa: BLE001 - un candidato que no sirve se descarta
            probados.append((url, f'descartado: {str(error)[:120]}'))
            continue
        probados.append((url, f'período {periodo}'))
        if mejor is None or periodo > mejor[0]:
            mejor = (periodo, resultado, url)
    if mejor is None:
        detalle = '; '.join(f'{u.rsplit("/", 1)[-1]} ({m})' for u, m in probados[:6])
        raise FileNotFoundError('Ningún archivo publicado sirvió: ' + (detalle or 'no hay candidatos'))
    return mejor[1], mejor[2], probados
