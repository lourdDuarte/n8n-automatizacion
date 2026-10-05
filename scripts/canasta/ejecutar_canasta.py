"""Descarga el ultimo informe CBA/CBT de INDEC y actualiza YQ:ZL.

Dependencias: openpyxl, pdfplumber.
"""

from __future__ import annotations

import argparse
import io
import json
import math
import re
import sys
import urllib.parse
import urllib.request
from datetime import date, datetime
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
from comun.columnas_por_titulo import resolver_columnas


# El endpoint /Nivel4/Tema/... contiene los enlaces reales. La ruta
# /indec/web/... es solo la envoltura de navegación del navegador.
PAGINA_INDEC = "https://www.indec.gob.ar/Nivel4/Tema/4/43/149"
HOJA_BASE = "BASE DE DATOS"
ARCHIVO_BASE = "BASE DE DATOS INDICADORES PROPIOS.xlsx"
MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}

DESTINOS = {
    "YQ": "CBA. Valores mensuales por adulto equivalente. Línea de indigencia. (Valor en $)",
    "YR": "VAR% mensual CBA adulto equivalente",
    "YS": "VAR% anual CBA adulto equivalente",
    "YT": "Acumulado CBA adulto equivalente año en curso",
    "YU": "CBT. Valores mensuales por adulto equivalente. Línea de pobreza (Valor en $)",
    "YV": "VAR% mensual CBT adulto equivalente",
    "YW": "VAR% anual CBT adulto equivalente",
    "YX": "Acumulado CBT adulto equivalente año en curso",
    "YY": "Cálculo de la CBA para un hogar de ejemplo. Hogar 1",
    "YZ": "Cálculo de la CBA para un hogar de ejemplo. Hogar 2",
    "ZA": "VAR% mensual CBA hogar 2",
    "ZB": "VAR% anual CBA hogar 2",
    "ZC": "Cálculo de la CBA para un hogar de ejemplo. Hogar 3",
    "ZD": "VAR% mensual CBA hogar 3",
    "ZE": "VAR% anual CBA hogar 3",
    "ZF": "Cálculo de la CBT para un hogar de ejemplo. Hogar 1",
    "ZG": "Cálculo de la CBT para un hogar de ejemplo. Hogar 2",
    "ZH": "VAR% mensual CBT hogar 2",
    "ZI": "VAR% anual CBT hogar 2",
    "ZJ": "Cálculo de la CBT para un hogar de ejemplo. Hogar 3",
    "ZK": "VAR% mensual CBT hogar 3",
    "ZL": "VAR% anual CBT hogar 3",
}

# Estas son las columnas fuente publicadas. Los acumulados y las variaciones
# de hogares son fórmulas propias del observatorio y nunca se reemplazan.
CODIGOS_ESCRITURA = (
    "YQ", "YR", "YS", "YU", "YV", "YW",
    "YY", "YZ", "ZC", "ZF", "ZG", "ZJ",
)


def raiz_proyecto() -> Path:
    return Path(__file__).resolve().parents[2]


def pedir(url: str, timeout: int = 60) -> bytes:
    solicitud = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (automatizacion-indicadores/1.0)"},
    )
    with urllib.request.urlopen(solicitud, timeout=timeout) as respuesta:
        return respuesta.read()


def descubrir_informe() -> tuple[str, bytes]:
    html = pedir(PAGINA_INDEC).decode("utf-8", errors="replace")
    enlaces = re.findall(
        r'''(?:href\s*=\s*["'])([^"']*canasta_[^"']*\.pdf[^"']*)''',
        html,
        flags=re.IGNORECASE,
    )
    if not enlaces:
        raise FileNotFoundError("No se encontro el enlace del informe CBA/CBT en INDEC")

    candidatos = []
    for enlace in enlaces:
        url = urllib.parse.urljoin(PAGINA_INDEC, enlace.replace("&amp;", "&"))
        nombre = urllib.parse.urlparse(url).path.rsplit("/", 1)[-1]
        coincidencia = re.search(r"canasta_(\d{2})_(\d{2})", nombre, re.IGNORECASE)
        orden = (0, 0)
        if coincidencia:
            orden = (2000 + int(coincidencia.group(2)), int(coincidencia.group(1)))
        candidatos.append((orden, url))

    for _, url in sorted(candidatos, reverse=True):
        contenido = pedir(url)
        if contenido.startswith(b"%PDF"):
            return url, contenido
    raise ValueError("Los enlaces encontrados no devolvieron un PDF valido")


def numero(texto: str) -> float:
    limpio = texto.strip().replace("$", "").replace(" ", "")
    limpio = limpio.replace(".", "").replace(",", ".")
    return float(limpio)


def fila_periodo(texto: str, mes: str) -> list[float]:
    patron = re.compile(
        rf"(?:20\d{{2}}\s+)?{mes}\s+"
        r"([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+"
        r"(-?\d+,\d)\s+(-?\d+,\d)\s+(-?\d+,\d)",
        flags=re.IGNORECASE,
    )
    coincidencias = patron.findall(texto)
    if not coincidencias:
        raise ValueError(f"No se encontro la fila de {mes} en el informe")
    return [numero(item) for item in coincidencias[-1]]


def extraer(contenido: bytes, url: str) -> dict:
    with pdfplumber.open(io.BytesIO(contenido)) as pdf:
        textos = [pagina.extract_text(x_tolerance=2, y_tolerance=3) or "" for pagina in pdf.pages]
    texto = "\n".join(textos)

    titulo = re.search(r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s+(?:de\s+)?(20\d{2})", texto, re.IGNORECASE)
    if not titulo:
        raise ValueError("No se pudo determinar el periodo del informe")
    mes_nombre = titulo.group(1).casefold()
    anio = int(titulo.group(2))

    # Se toma la última coincidencia de cada cuadro: es el período más reciente.
    nombres_meses = "|".join(MESES)
    patron = re.compile(
        rf"(?:20\d{{2}}\s+)?(?:{nombres_meses})\s+"
        r"([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+"
        r"(-?\d+,\d)\s+(-?\d+,\d)\s+(-?\d+,\d)",
        flags=re.IGNORECASE,
    )
    filas_cba, filas_cbt = [], []
    for pagina in textos:
        coincidencias = [[numero(item) for item in grupo] for grupo in patron.findall(pagina)]
        if "Cuadro 1." in pagina:
            filas_cba.extend(coincidencias)
        if "Cuadro 2." in pagina:
            filas_cbt.extend(coincidencias)
    if len(filas_cba) < 13 or len(filas_cbt) < 13:
        raise ValueError(f"Se esperaban las filas CBA y CBT de {mes_nombre} {anio}")
    cba, cbt = filas_cba[-1], filas_cbt[-1]
    cba_anterior, cbt_anterior = filas_cba[-2], filas_cbt[-2]
    cba_interanual, cbt_interanual = filas_cba[-13], filas_cbt[-13]

    def variacion(actual: float, anterior: float) -> float:
        if anterior == 0:
            raise ValueError("No se puede calcular una variacion contra cero")
        return round((actual / anterior - 1) * 100, 2)

    # Cada fila: adulto, hogar 1, hogar 2, hogar 3, mensual, acumulada, interanual.
    valores = {
        "YQ": cba[0],
        "YR": cba[4],
        "YS": cba[6],
        # YT y YX se recalculan al actualizar la base usando el acumulado anterior.
        "YT": cba[5],
        "YU": cbt[0],
        "YV": cbt[4],
        "YW": cbt[6],
        "YX": cbt[5],
        "YY": cba[1],
        "YZ": cba[2],
        "ZA": variacion(cba[2], cba_anterior[2]),
        "ZB": variacion(cba[2], cba_interanual[2]),
        "ZC": cba[3],
        "ZD": variacion(cba[3], cba_anterior[3]),
        "ZE": variacion(cba[3], cba_interanual[3]),
        "ZF": cbt[1],
        "ZG": cbt[2],
        "ZH": variacion(cbt[2], cbt_anterior[2]),
        "ZI": variacion(cbt[2], cbt_interanual[2]),
        "ZJ": cbt[3],
        "ZK": variacion(cbt[3], cbt_anterior[3]),
        "ZL": variacion(cbt[3], cbt_interanual[3]),
    }
    return {
        "fuente": "INDEC - CBA y CBT",
        "periodo": date(anio, MESES[mes_nombre], 1).isoformat(),
        "anio": anio,
        "mes_numero": MESES[mes_nombre],
        "url_informe": url,
        "consultado_en": datetime.now().isoformat(timespec="seconds"),
        "valores": valores,
        "validaciones": {"cantidad_valores": len(valores), "filas_cba_cbt_detectadas": True},
    }


def localizar_fila(ws, anio: int, mes: int) -> int:
    for fila in range(2, ws.max_row + 1):
        valor = ws.cell(fila, 1).value
        if isinstance(valor, (date, datetime)) and valor.year == anio and valor.month == mes:
            return fila
    raise ValueError(f"No existe el periodo {anio}-{mes:02d} en la columna A")


def actualizar_base(resultado: dict, ruta_base: Path) -> dict:
    if not ruta_base.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta_base}")
    libro = load_workbook(ruta_base, data_only=False)
    if HOJA_BASE not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {HOJA_BASE!r}")
    ws = libro[HOJA_BASE]
    columnas = resolver_columnas(ws, DESTINOS)
    fila = localizar_fila(ws, resultado["anio"], resultado["mes_numero"])
    cambios, coincidencias = [], []
    for codigo in CODIGOS_ESCRITURA:
        nuevo = resultado["valores"][codigo]
        celda = ws.cell(fila, columnas[codigo])
        if isinstance(celda.value, str) and celda.value.startswith("="):
            raise ValueError(f"No se reemplazó la fórmula existente en {celda.coordinate}")
        anterior = celda.value
        if isinstance(anterior, (int, float)) and math.isclose(float(anterior), nuevo, rel_tol=0, abs_tol=1e-9):
            coincidencias.append(celda.coordinate)
        elif anterior == nuevo:
            coincidencias.append(celda.coordinate)
        else:
            celda.value = nuevo
            cambios.append({"celda": celda.coordinate, "titulo": DESTINOS[codigo], "antes": anterior, "despues": nuevo})
    if cambios:
        libro.save(ruta_base)
    return {"archivo_excel": str(ruta_base.resolve()), "fila": fila, "celdas_actualizadas": len(cambios), "cambios": cambios, "coincidencias": coincidencias}


def guardar_reporte(resultado: dict, carpeta: Path) -> Path:
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / f"canasta_{resultado['anio']}_{resultado['mes_numero']:02d}.json"
    ruta.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    return ruta


def main() -> int:
    raiz = raiz_proyecto()
    parser = argparse.ArgumentParser(description="Actualiza CBA/CBT en YQ:ZL.")
    parser.add_argument("--base", type=Path, default=raiz / "base" / ARCHIVO_BASE)
    parser.add_argument("--solo-extraer", action="store_true")
    parser.add_argument("--pdf", type=Path, help="Usa un informe PDF local para pruebas.")
    args = parser.parse_args()
    try:
        if args.pdf:
            contenido, url = args.pdf.read_bytes(), args.pdf.resolve().as_uri()
        else:
            url, contenido = descubrir_informe()
        resultado = extraer(contenido, url)
        if not args.solo_extraer:
            resultado["actualizacion"] = actualizar_base(resultado, args.base)
        reporte = guardar_reporte(resultado, raiz / "scripts" / "reportes" / "canasta")
        resultado["reporte"] = str(reporte.resolve())
        print(json.dumps(resultado, ensure_ascii=False, indent=2))
        print("ESTADO=OK")
        return 0
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
