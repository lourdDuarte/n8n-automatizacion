"""Consulta ATP Formosa y actualiza la recaudación tributaria mensual.

El destino se resuelve por el encabezado PERIODO y por el título del
indicador. No depende de letras de columnas ni de números de filas.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import urllib.request
from copy import copy
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook


RAIZ = Path(__file__).resolve().parents[2]
SCRIPTS = RAIZ / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from comun.columnas_por_titulo import resolver_columnas


URL_FUENTE = "https://www.formosa.gob.ar/atp/recaudaciones"
BASE_PREDETERMINADA = RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA_PREDETERMINADA = "BASE DE DATOS"
TITULO_RECAUDACION = "Recaudación tributaria F"

MESES = {
    "ene": 1, "enero": 1, "feb": 2, "febrero": 2,
    "mar": 3, "marzo": 3, "abr": 4, "abril": 4,
    "may": 5, "mayo": 5, "jun": 6, "junio": 6,
    "jul": 7, "julio": 7, "ago": 8, "agosto": 8,
    "sep": 9, "sept": 9, "septiembre": 9, "setiembre": 9,
    "oct": 10, "octubre": 10, "nov": 11, "noviembre": 11,
    "dic": 12, "diciembre": 12,
}


def configurar_salida() -> None:
    for flujo in (sys.stdout, sys.stderr):
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def normalizar(valor: object) -> str:
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", texto.casefold()).strip()


def numero_argentino(texto: str) -> float:
    limpio = re.sub(r"[^0-9,.-]", "", texto).replace(".", "").replace(",", ".")
    if not limpio or limpio in {"-", ".", "-."}:
        raise ValueError(f"Importe ATP inválido: {texto!r}")
    return float(limpio)


def descargar_html(timeout: int = 60) -> str:
    solicitud = urllib.request.Request(
        URL_FUENTE,
        headers={"User-Agent": "Mozilla/5.0 (automatizacion-indicadores/1.0)"},
    )
    with urllib.request.urlopen(solicitud, timeout=timeout) as respuesta:
        contenido = respuesta.read()
        codificacion = respuesta.headers.get_content_charset() or "utf-8"
    return contenido.decode(codificacion, errors="replace")


class LectorTablas(HTMLParser):
    """Convierte las tablas HTML en matrices de texto sin dependencias externas."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tablas: list[list[list[str]]] = []
        self._tabla: list[list[str]] | None = None
        self._fila: list[str] | None = None
        self._celda: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.casefold()
        if tag == "table":
            self._tabla = []
        elif tag == "tr" and self._tabla is not None:
            self._fila = []
        elif tag in {"td", "th"} and self._fila is not None:
            self._celda = []

    def handle_data(self, data: str) -> None:
        if self._celda is not None:
            self._celda.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.casefold()
        if tag in {"td", "th"} and self._celda is not None and self._fila is not None:
            self._fila.append(" ".join("".join(self._celda).split()))
            self._celda = None
        elif tag == "tr" and self._fila is not None and self._tabla is not None:
            if self._fila:
                self._tabla.append(self._fila)
            self._fila = None
        elif tag == "table" and self._tabla is not None:
            if self._tabla:
                self.tablas.append(self._tabla)
            self._tabla = None


def extraer_ultimo_periodo(html: str, anio: int) -> dict:
    lector = LectorTablas()
    lector.feed(html)
    candidatos: list[dict] = []

    for filas in lector.tablas:
        if not filas:
            continue
        encabezados = [normalizar(celda) for celda in filas[0]]
        if not encabezados or "totales" not in encabezados:
            continue
        columna_total = encabezados.index("totales")

        for fila in filas[1:]:
            celdas = fila
            if len(celdas) <= columna_total:
                continue
            mes_texto = normalizar(celdas[0])
            mes = MESES.get(mes_texto)
            if mes is None:
                continue
            candidatos.append({
                "periodo": f"{anio:04d}-{mes:02d}",
                "anio": anio,
                "mes": mes,
                "mes_fuente": celdas[0],
                "recaudacion": numero_argentino(celdas[columna_total]),
            })

    if not candidatos:
        raise ValueError("No se encontró la tabla mensual de recaudación de ATP.")
    return max(candidatos, key=lambda item: (item["anio"], item["mes"]))


def periodo_celda(valor: object) -> tuple[int, int] | None:
    if isinstance(valor, (date, datetime)):
        return valor.year, valor.month
    texto = normalizar(valor)
    if not texto:
        return None
    coincidencia = re.fullmatch(r"(20\d{2}) (0?[1-9]|1[0-2])", texto)
    if coincidencia:
        return int(coincidencia.group(1)), int(coincidencia.group(2))
    coincidencia = re.fullmatch(r"([a-z]+) (\d{2}|20\d{2})", texto)
    if coincidencia and coincidencia.group(1) in MESES:
        anio = int(coincidencia.group(2))
        return (anio + 2000 if anio < 100 else anio), MESES[coincidencia.group(1)]
    return None


def buscar_fila(hoja, columna_periodo: int, periodo: str) -> int:
    objetivo = tuple(map(int, periodo.split("-")))
    coincidencias = [
        fila
        for fila in range(2, hoja.max_row + 1)
        if periodo_celda(hoja.cell(fila, columna_periodo).value) == objetivo
    ]
    if not coincidencias:
        raise ValueError(f"No se encontró el período {periodo} en PERIODO.")
    if len(coincidencias) != 1:
        raise ValueError(f"El período {periodo} está repetido en las filas {coincidencias}.")
    return coincidencias[0]


def iguales(actual: object, nuevo: float, tolerancia: float = 0.005) -> bool:
    return (
        isinstance(actual, (int, float))
        and not isinstance(actual, bool)
        and abs(float(actual) - nuevo) <= tolerancia
    )


def guardar_atomico(libro, ruta: Path) -> None:
    descriptor, nombre = tempfile.mkstemp(prefix=f"{ruta.stem}_", suffix=ruta.suffix, dir=ruta.parent)
    os.close(descriptor)
    temporal = Path(nombre)
    try:
        libro.save(temporal)
        os.replace(temporal, ruta)
    finally:
        if temporal.exists():
            temporal.unlink()


def actualizar_base(ruta: Path, hoja_nombre: str, dato: dict, sobrescribir: bool) -> dict:
    libro = load_workbook(ruta, data_only=False, keep_links=True)
    if hoja_nombre not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {hoja_nombre!r}.")
    hoja = libro[hoja_nombre]
    columnas = resolver_columnas(
        hoja,
        {"periodo": "PERIODO", "recaudacion": TITULO_RECAUDACION},
    )
    fila = buscar_fila(hoja, columnas["periodo"], dato["periodo"])
    celda = hoja.cell(fila, columnas["recaudacion"])
    anterior = celda.value
    nuevo = float(dato["recaudacion"])

    if isinstance(anterior, str) and anterior.startswith("="):
        raise ValueError(f"Seguridad: {celda.coordinate} contiene una fórmula.")
    if iguales(anterior, nuevo):
        return {
            "archivo_excel": str(ruta.resolve()), "hoja": hoja_nombre,
            "periodo": dato["periodo"], "fila_localizada": fila,
            "celdas_actualizadas": 0, "hubo_cambios": False,
            "cambios": [], "respaldo": None,
        }
    vacio_o_cero = anterior in (None, "") or (
        isinstance(anterior, (int, float)) and not isinstance(anterior, bool) and float(anterior) == 0.0
    )
    if not vacio_o_cero and not sobrescribir:
        raise ValueError(
            f"{celda.coordinate} ya contiene {anterior!r}, distinto de {nuevo}. "
            "Revisá el dato o usá --sobrescribir."
        )

    carpeta = ruta.parent / "respaldos"
    carpeta.mkdir(parents=True, exist_ok=True)
    respaldo = carpeta / f"antes_atp_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(ruta, respaldo)

    celda.value = nuevo
    if not celda.number_format or celda.number_format == "General":
        celda.number_format = copy(hoja.cell(fila - 1, celda.column).number_format)
    libro.calculation.fullCalcOnLoad = True
    libro.calculation.forceFullCalc = True
    libro.calculation.calcMode = "auto"
    guardar_atomico(libro, ruta)

    control = load_workbook(ruta, data_only=False, read_only=True, keep_links=True)
    try:
        guardado = control[hoja_nombre][celda.coordinate].value
    finally:
        control.close()
    if not iguales(guardado, nuevo):
        raise RuntimeError(f"Falló la verificación de {celda.coordinate}.")

    return {
        "archivo_excel": str(ruta.resolve()), "hoja": hoja_nombre,
        "periodo": dato["periodo"], "fila_localizada": fila,
        "celdas_actualizadas": 1, "hubo_cambios": True,
        "cambios": [{
            "celda": celda.coordinate, "titulo": TITULO_RECAUDACION,
            "antes": anterior, "despues": nuevo,
        }],
        "respaldo": str(respaldo.resolve()),
    }


def main() -> int:
    configurar_salida()
    parser = argparse.ArgumentParser(description="Actualiza la recaudación tributaria de ATP Formosa.")
    parser.add_argument("--excel", type=Path, default=BASE_PREDETERMINADA)
    parser.add_argument("--hoja", default=HOJA_PREDETERMINADA)
    parser.add_argument("--html", type=Path, help="HTML local, únicamente para pruebas.")
    parser.add_argument("--anio", type=int, default=date.today().year)
    parser.add_argument("--sobrescribir", action="store_true")
    args = parser.parse_args()

    ruta = args.excel.expanduser().resolve()
    if not ruta.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta}")
    html = args.html.read_text(encoding="utf-8") if args.html else descargar_html()
    dato = extraer_ultimo_periodo(html, args.anio)
    actualizacion = actualizar_base(ruta, args.hoja, dato, args.sobrescribir)

    resultado = {
        "fuente": "ATP - Recaudación Tributaria Provincial",
        "url": URL_FUENTE,
        "consultado_en": datetime.now().astimezone().isoformat(timespec="seconds"),
        "periodo": dato["periodo"],
        "recaudacion": dato["recaudacion"],
        "actualizacion": actualizacion,
    }
    carpeta_reporte = RAIZ / "scripts" / "reportes" / "atp"
    carpeta_reporte.mkdir(parents=True, exist_ok=True)
    reporte = carpeta_reporte / f"atp_{dato['anio']:04d}_{dato['mes']:02d}.json"
    reporte.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    resultado["reporte"] = str(reporte.resolve())

    print("RESULTADO_ATP=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
    print("ESTADO=OK")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        raise SystemExit(1)
