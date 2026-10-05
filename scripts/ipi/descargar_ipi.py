"""Descarga desde INDEC el IPI esperado para la fecha de ejecución."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from urllib.parse import urljoin
from urllib.request import Request, urlopen
import re
import sys
import time

import fitz


RAIZ = Path(__file__).resolve().parent.parent
DESTINO = RAIZ / "fuentes" / "ipi"
PAGINA = "https://www.indec.gob.ar/Nivel4/Tema/3/6/14"
MESES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126 Safari/537.36"


def configurar_salida() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def obtener(url: str, intentos: int = 3, timeout: int = 75) -> bytes:
    ultimo_error: Exception | None = None
    for intento in range(1, intentos + 1):
        pedido = Request(
            url,
            headers={
                "User-Agent": UA,
                "Accept-Language": "es-AR,es;q=0.9",
            },
        )
        try:
            with urlopen(pedido, timeout=timeout) as respuesta:
                return respuesta.read()
        except Exception as error:
            ultimo_error = error
            if intento < intentos:
                time.sleep(2 * intento)

    assert ultimo_error is not None
    raise ultimo_error


def periodo_esperado(hoy: date | None = None) -> str:
    """El IPI difundido en un mes corresponde a dos meses antes."""
    hoy = hoy or date.today()
    indice = hoy.year * 12 + hoy.month - 1 - 2
    return f"{indice // 12:04d}-{indice % 12 + 1:02d}"


def detectar_periodo_pdf(contenido: bytes) -> str:
    with fitz.open(stream=contenido, filetype="pdf") as documento:
        texto = "\n".join(p.get_text("text") for p in list(documento)[:5])

    patron = re.compile(
        rf"\b({'|'.join(MESES)})\s+de\s+(20\d{{2}})\b",
        re.IGNORECASE,
    )
    halladas = patron.findall(texto)
    if not halladas:
        raise ValueError("el PDF no informa un período reconocible")

    mes, anio = halladas[-1]
    return f"{int(anio):04d}-{MESES[mes.lower()]:02d}"


def enlaces_ipi(html: str) -> list[str]:
    rutas = re.findall(
        r'''(?:href|src)=["']([^"']*ipi_manufacturero_[^"']+?\.pdf(?:\?[^"']*)?)["']''',
        html,
        re.IGNORECASE,
    )
    vistos: set[str] = set()
    enlaces: list[str] = []

    for ruta in rutas:
        url = urljoin(PAGINA, ruta)
        if url not in vistos:
            vistos.add(url)
            enlaces.append(url)

    return enlaces


def enlace_planilla(contenido_pdf: bytes) -> str:
    """Obtiene del informe el XLS con los cuadros web detallados."""
    patron = re.compile(
        r"https?://[^\s<>]+/sh_ipi_manufacturero_\d{4}\.xls",
        re.IGNORECASE,
    )

    with fitz.open(stream=contenido_pdf, filetype="pdf") as documento:
        candidatos: list[str] = []
        for pagina in documento:
            candidatos.extend(
                enlace.get("uri", "")
                for enlace in pagina.get_links()
                if enlace.get("uri")
            )
            candidatos.extend(patron.findall(pagina.get_text("text")))

    for candidato in candidatos:
        coincidencia = patron.search(candidato)
        if coincidencia:
            return coincidencia.group(0).rstrip(".,;)")

    raise RuntimeError(
        "El informe IPI no contiene el enlace a la planilla XLS de cuadros web."
    )


def asegurar_planilla(contenido_pdf: bytes, periodo: str) -> Path:
    url = enlace_planilla(contenido_pdf)
    contenido = obtener(url)

    # Los archivos XLS clásicos comienzan con esta firma.
    if not contenido.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        raise RuntimeError(
            "La planilla detallada descargada no es un XLS válido."
        )

    destino = (
        DESTINO
        / f"sh_ipi_manufacturero_{periodo.replace('-', '_')}.xls"
    )
    DESTINO.mkdir(parents=True, exist_ok=True)

    if destino.exists() and destino.read_bytes() == contenido:
        print(
            f"Sin descarga: la planilla detallada ya coincide: {destino}"
        )
    else:
        destino.write_bytes(contenido)
        print(f"Planilla detallada descargada: {destino}")

    print(f"XLS_IPI={destino.resolve()}")
    return destino


def main() -> int:
    configurar_salida()
    esperado = periodo_esperado()

    print("=" * 70)
    print("DESCARGA AUTOMÁTICA — IPI MANUFACTURERO")
    print("=" * 70)
    print(f"Período esperado: {esperado}")

    # Se consulta INDEC en cada ejecución, aunque exista un PDF local.
    html = obtener(PAGINA).decode("utf-8", errors="replace")
    enlaces = enlaces_ipi(html)

    if not enlaces:
        raise RuntimeError(
            "INDEC no publicó enlaces PDF del IPI en la página oficial."
        )

    encontrados: list[str] = []

    for url in enlaces:
        try:
            contenido = obtener(url)
            if not contenido.startswith(b"%PDF"):
                continue

            periodo = detectar_periodo_pdf(contenido)
            encontrados.append(periodo)

            if periodo != esperado:
                continue

            DESTINO.mkdir(parents=True, exist_ok=True)
            destino = (
                DESTINO
                / f"ipi_manufacturero_{periodo.replace('-', '_')}.pdf"
            )

            if destino.exists() and destino.read_bytes() == contenido:
                print(
                    f"Sin descarga: el informe oficial coincide: {destino}"
                )
            else:
                destino.write_bytes(contenido)
                print(f"Informe oficial descargado: {destino}")

            asegurar_planilla(contenido, periodo)

            print(f"URL: {url}")
            print(f"PERIODO_IPI={periodo}")
            print(f"PDF_IPI={destino.resolve()}")
            print("RESULTADO: OK")
            return 0

        except Exception as error:
            print(f"AVISO: enlace descartado: {error}")

    ultimo = max(encontrados) if encontrados else "ninguno"
    raise RuntimeError(
        f"El IPI {esperado} todavía no está disponible. "
        f"Último período oficial detectado: {ultimo}."
    )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        configurar_salida()
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)