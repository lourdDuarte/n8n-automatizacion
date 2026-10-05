"""Detecta y descarga automáticamente la última base mensual de CAMMESA."""

from __future__ import annotations

from html import unescape
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.parse import urljoin
from urllib.request import HTTPCookieProcessor, Request, build_opener
import io
import re
import sys
import time
import zipfile


RAIZ = Path(__file__).resolve().parent.parent
DESTINO = RAIZ / "fuentes" / "cammesa"
LISTADO = "https://cammesaweb.cammesa.com/download-category/sintesis-mensual/"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/150 Safari/537.36"
ABRIDOR = build_opener(HTTPCookieProcessor(CookieJar()))


def configurar_salida() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def obtener(url: str, intentos: int = 3, timeout: int = 60) -> bytes:
    ultimo = None
    for intento in range(1, intentos + 1):
        try:
            pedido = Request(url, headers={"User-Agent": UA, "Accept-Language": "es-AR,es;q=0.9"})
            with ABRIDOR.open(pedido, timeout=timeout) as respuesta:
                return respuesta.read()
        except Exception as error:
            ultimo = error
            if intento < intentos:
                time.sleep(2 * intento)
    raise RuntimeError(f"No se pudo descargar {url}: reintentos agotados: {ultimo}")


def zip_valido(contenido: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(contenido)) as archivo:
            nombres = [n.replace("\\", "/") for n in archivo.namelist()]
            return any(n.endswith("Bases_Demanda_INFORME_MENSUAL/Demanda Mensual.xlsx") for n in nombres)
    except zipfile.BadZipFile:
        return False


def periodos_publicados() -> list[str]:
    encontrados = set()
    # Las publicaciones más nuevas están en las primeras páginas.
    for pagina in range(1, 3):
        url = LISTADO if pagina == 1 else urljoin(LISTADO, f"page/{pagina}/")
        html = obtener(url).decode("utf-8", errors="replace")
        encontrados.update(re.findall(r"BASE_INFORME_MENSUAL_(20\d{2}-(?:0[1-9]|1[0-2]))", html, re.I))
    if not encontrados:
        raise RuntimeError("CAMMESA no mostró publicaciones BASE_INFORME_MENSUAL en el listado oficial.")
    return sorted(encontrados, reverse=True)


def urls_descarga(html: str, pagina: str) -> list[str]:
    texto = unescape(html).replace("\\/", "/")
    candidatos = []
    patrones = (
        r'''data-downloadurl=["']([^"']+)["']''',
        r'''href=["']([^"']*(?:wpdmdl=\d+|\.zip(?:\?[^"']*)?))['"]''',
        r'''(?:location\.href|window\.open)\s*\(?\s*["']([^"']*(?:wpdmdl=\d+|\.zip(?:\?[^"']*)?))['"]''',
        r'''(https?://[^"'<>\s]+(?:wpdmdl=\d+|\.zip(?:\?[^"'<>\s]*)?))''',
    )
    for patron in patrones:
        candidatos.extend(re.findall(patron, texto, re.I))
    vistos = set()
    resultado = []
    for candidato in candidatos:
        url = urljoin(pagina, candidato.strip())
        if url not in vistos:
            vistos.add(url)
            resultado.append(url)
    return resultado


def descargar_periodo(periodo: str) -> tuple[Path, str]:
    DESTINO.mkdir(parents=True, exist_ok=True)
    local = DESTINO / ("BASE_INFORME_MENSUAL_" + periodo + ".zip")
    if local.is_file() and zip_valido(local.read_bytes()):
        return local, "archivo local válido"

    pagina = f"https://cammesaweb.cammesa.com/download/base_informe_mensual_{periodo}/"
    html = obtener(pagina).decode("utf-8", errors="replace")
    candidatos = urls_descarga(html, pagina)
    if not candidatos:
        raise RuntimeError(f"CAMMESA publicó {periodo}, pero la página no expuso el enlace de descarga.")
    errores = []
    for url in candidatos:
        try:
            contenido = obtener(url)
            if not zip_valido(contenido):
                errores.append(f"contenido no válido: {url}")
                continue
            local.write_bytes(contenido)
            return local, url
        except Exception as error:
            errores.append(str(error))
    raise RuntimeError("No se obtuvo un ZIP válido de CAMMESA. " + " | ".join(errores[-3:]))


def main() -> int:
    configurar_salida()
    errores = []
    for periodo in periodos_publicados():
        try:
            ruta, origen = descargar_periodo(periodo)
            print("=" * 72)
            print("DESCARGA — CAMMESA BASE INFORME MENSUAL")
            print("=" * 72)
            print(f"Período oficial más reciente: {periodo}")
            print(f"Origen: {origen}")
            print(f"ZIP_CAMMESA={ruta.resolve()}")
            print("RESULTADO: OK")
            return 0
        except Exception as error:
            errores.append(f"{periodo}: {error}")
    raise RuntimeError("No se pudo descargar ninguna base mensual válida. " + " | ".join(errores[:3]))


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        configurar_salida()
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
