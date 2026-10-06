import argparse
import hashlib
import html
import json
import re
import sys
import unicodedata
import urllib.request
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urljoin, urlparse

import fitz

RAIZ = Path(__file__).resolve().parents[2]
PAGINA = "https://www.iaraf.org/index.php"

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11,
    "diciembre": 12,
}

TITULOS = [
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ F",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Nominal F",
    "Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Var% Anual Real F",
    "Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Acumulado anual. En millones de $ F",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ N",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Nominal N",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Real N",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Acumulado anual. En millones de $ N",
]

def normalizar(texto):
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(
        c for c in texto if unicodedata.category(c) != "Mn"
    )
    return " ".join(texto.split()).lower()

def limpiar_html(texto):
    return html.unescape(re.sub(r"<[^>]+>", " ", texto))

def descargar(url):
    pedido = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(pedido, timeout=45) as respuesta:
        return respuesta.read()

def fecha_publicacion(url, titulo):
    m = re.search(
        r"(\d{2})-(\d{2})-(\d{2}|\d{4})(?:$|[/?])",
        url,
    )
    if m:
        anio = int(m[3])
        if anio < 100:
            anio += 2000
        return datetime(anio, int(m[2]), int(m[1]))

    patron = "|".join(MESES)
    m = re.search(
        rf"\b({patron})\s+(?:de\s+)?(20\d{{2}})\b",
        normalizar(titulo),
    )
    if m:
        return datetime(int(m[2]), MESES[m[1]], 1)

    raise ValueError(
        f"No se identifica la fecha del informe: {titulo}"
    )

def descubrir():
    # Busca publicaciones recientes, incluyendo paginacion.
    candidatos = {}
    for inicio in range(0, 50, 5):
        url = PAGINA if inicio == 0 else f"{PAGINA}?start={inicio}"
        pagina = descargar(url).decode("utf-8", errors="replace")

        bloques = re.findall(
            r"<h2\b[^>]*>(.*?)</h2>(.*?)(?=<h2\b|$)",
            pagina,
            flags=re.I | re.S,
        )

        for titulo_html, cuerpo in bloques:
            titulo = limpiar_html(titulo_html)
            texto = normalizar(titulo)

            if not (
                "transferencias automaticas" in texto
                or "coparticipacion nacional" in texto
            ):
                continue

            for etiqueta, interior in re.findall(
                r"<a\b([^>]*)>(.*?)</a>",
                cuerpo,
                flags=re.I | re.S,
            ):
                if "leer mas" not in normalizar(limpiar_html(interior)):
                    continue
                enlace = re.search(
                    r'href=["\']([^"\']+)["\']',
                    etiqueta,
                    flags=re.I,
                )
                if enlace:
                    destino = urljoin(
                        PAGINA,
                        html.unescape(enlace[1]),
                    )
                    candidatos[destino] = (
                        fecha_publicacion(destino, titulo),
                        titulo,
                    )
                    break

        if candidatos:
            # El inicio presenta las publicaciones mas recientes.
            break

    if not candidatos:
        raise RuntimeError(
            "IARAF no mostro un informe mensual de "
            "transferencias en las paginas revisadas."
        )

    url = max(candidatos, key=lambda x: candidatos[x][0])
    return url, candidatos[url][1]

def enlace_pdf(url_nota):
    pagina = descargar(url_nota).decode("utf-8", errors="replace")
    enlaces = [
        html.unescape(x)
        for x in re.findall(
            r'href=["\']([^"\']+)["\']',
            pagina,
            flags=re.I,
        )
    ]

    for enlace in enlaces:
        absoluto = urljoin(url_nota, enlace)
        if "drive.google.com" in absoluto:
            partes = urlparse(absoluto)
            identificador = parse_qs(partes.query).get("id", [None])[0]
            if not identificador:
                m = re.search(r"/file/d/([^/]+)", partes.path)
                identificador = m[1] if m else None
            if identificador:
                return (
                    "https://drive.google.com/uc?export=download&id="
                    + identificador
                )

        if urlparse(absoluto).path.lower().endswith(".pdf"):
            return absoluto

    raise RuntimeError(
        "El informe mensual existe, pero no se encontro "
        "su enlace de descarga."
    )

def numero(texto):
    texto = texto.strip().replace("−", "-").replace("–", "-")
    if not re.fullmatch(r"-?\d[\d.]*(?:,\d+)?%?", texto):
        raise ValueError(f"Valor no numerico en el cuadro: {texto}")
    return float(
        texto.replace(".", "").replace(",", ".").replace("%", "")
    )

def fila(texto, jurisdiccion):
    lineas = [x.strip() for x in texto.splitlines() if x.strip()]
    posiciones = [
        i for i, linea in enumerate(lineas)
        if normalizar(linea) == normalizar(jurisdiccion)
    ]
    if len(posiciones) != 1:
        raise ValueError(
            f"Fila inexistente o duplicada: {jurisdiccion}"
        )
    inicio = posiciones[0] + 1
    datos = lineas[inicio:inicio + 3]
    if len(datos) != 3:
        raise ValueError(f"Fila incompleta: {jurisdiccion}")
    return [numero(x) for x in datos]

def extraer_pdf(contenido):
    documento = fitz.open(stream=contenido, filetype="pdf")
    paginas = [pagina.get_text() for pagina in documento]
    documento.close()

    def cuadro(numero_cuadro):
        coincidencias = [
            texto for texto in paginas
            if re.search(
                rf"\bCuadro\s+{numero_cuadro}\b",
                texto,
                flags=re.I,
            )
            and "Jurisdicción" in texto
        ]
        if len(coincidencias) != 1:
            raise ValueError(
                f"No se identifica un unico Cuadro {numero_cuadro}. "
                "Posible cambio de estructura del informe."
            )
        return coincidencias[0]

    mensual = cuadro(1)
    anual = cuadro(5)

    patron = "|".join(MESES)
    fechas_mensuales = set(re.findall(
        rf"\b({patron})\s+(?:de\s+)?(20\d{{2}})\b",
        normalizar(mensual),
    ))
    if len(fechas_mensuales) != 1:
        raise ValueError(
            "No se identifica un periodo mensual unico."
        )

    nombre_mes, anio = next(iter(fechas_mensuales))
    periodo = f"{int(anio):04d}-{MESES[nombre_mes]:02d}-01"

    if not re.search(
        rf"\b{nombre_mes}\s+(?:de\s+)?{anio}\b",
        normalizar(anual),
    ):
        raise ValueError(
            "El cuadro mensual y el acumulado no tienen "
            "el mismo periodo."
        )

    formosa = fila(mensual, "Formosa")
    nacion = fila(mensual, "Total")
    acumulado_formosa = fila(anual, "Formosa")[0]
    acumulado_nacion = fila(anual, "Total")[0]

    if not (
        0 < formosa[0] <= nacion[0]
        and acumulado_formosa >= formosa[0]
        and acumulado_nacion >= nacion[0]
    ):
        raise ValueError(
            "Los montos mensuales y acumulados no son coherentes."
        )

    valores = dict(zip(TITULOS, [
        *formosa, acumulado_formosa,
        *nacion, acumulado_nacion,
    ]))

    return periodo, valores

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--solo-extraer", action="store_true")
    parser.add_argument("--pdf", type=Path)
    args = parser.parse_args()

    if args.pdf:
        contenido = args.pdf.read_bytes()
        url_nota = None
        url_pdf = None
        titulo = args.pdf.name
    else:
        url_nota, titulo = descubrir()
        url_pdf = enlace_pdf(url_nota)
        contenido = descargar(url_pdf)

    if not contenido.startswith(b"%PDF"):
        raise RuntimeError(
            "La descarga no devolvio un PDF. "
            "Puede haber una pantalla de acceso o un cambio de enlace."
        )

    periodo, valores = extraer_pdf(contenido)
    carpeta = RAIZ / "reportes" / "transferencias_fiscales"
    carpeta.mkdir(parents=True, exist_ok=True)

    pdf = carpeta / f"transferencias_{periodo[:7]}.pdf"
    pdf.write_bytes(contenido)

    resultado = {
        "fuente": "transferencias_fiscales",
        "nombre": "Transferencias automaticas - IARAF",
        "estado": "OK",
        "periodo": periodo,
        "consultado_en": datetime.now().astimezone().isoformat(),
        "url_informe": url_nota,
        "url_pdf": url_pdf,
        "titulo_informe": titulo,
        "sha256_pdf": hashlib.sha256(contenido).hexdigest(),
        "valores": valores,
        "columnas_preparadas": 8,
        "nota": (
            "Valores publicados por IARAF. Las variaciones reales "
            "pueden utilizar su estimacion de inflacion del mes. "
            "El acumulado corresponde a enero hasta el mes informado."
        ),
    }

    reporte = carpeta / f"transferencias_{periodo[:7]}.json"
    resultado["reporte"] = str(reporte)
    reporte.write_text(
        json.dumps(resultado, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(json.dumps(resultado, ensure_ascii=False, indent=2))
    print("ESTADO=OK")

if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"ERROR_TRANSFERENCIAS={error}")
        print("ESTADO=ERROR")
        raise SystemExit(1)