r"""Actualiza en Excel los datos fuente extraídos del informe IARAF.

Uso, desde la raíz de automatizacion_indicadores:
    python .\scripts\iaraf\actualizar_base_iaraf.py

Opciones:
    --json RUTA          Usa un control JSON específico.
    --excel RUTA         Usa otro libro de Excel.
    --sobrescribir       Permite reemplazar valores distintos de cero.

Escribe los seis datos fuente publicados por IARAF localizando sus encabezados:
monto mensual y variaciones nominal y real para Formosa y Nación.
Los acumulados del observatorio son fórmulas y nunca se reemplazan.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import unicodedata
from copy import copy
from datetime import date, datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit("Falta openpyxl. Instalalo con: pip install openpyxl") from exc


RAIZ = Path(__file__).resolve().parents[2]
EXCEL_PREDETERMINADO = RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
CARPETA_REPORTES = RAIZ / "reportes" / "iaraf"
HOJA_PREDETERMINADA = "BASE DE DATOS"

MESES = {
    "ene": 1, "enero": 1, "feb": 2, "febrero": 2,
    "mar": 3, "marzo": 3, "abr": 4, "abril": 4,
    "may": 5, "mayo": 5, "jun": 6, "junio": 6,
    "jul": 7, "julio": 7, "ago": 8, "agosto": 8,
    "sep": 9, "sept": 9, "septiembre": 9, "setiembre": 9,
    "oct": 10, "octubre": 10, "nov": 11, "noviembre": 11,
    "dic": 12, "diciembre": 12,
}

CAMPOS = {
    ("formosa", "monto_millones"): (
        "transferencias automaticas por coparticipacion leyes compl y compensac en millones f",
    ),
    ("formosa", "variacion_nominal_pct"): (
        "transferencias automaticas por coparticipacion leyes compl y compensac var anual nominal f",
        "transferencias automaticas por coparticipacion leyes complementarias y compensaciones var anual nominal f",
    ),
    ("formosa", "variacion_real_pct"): (
        "transferencias automaticas por coparticipacion leyes complementarias y compensaciones var anual real f",
        "transferencias automaticas por coparticipacion leyes compl y compensac var anual real f",
    ),
    ("formosa", "acumulado_anual_millones"): (
        "transferencias automaticas por coparticipacion leyes complementarias y compensaciones acumulado anual en millones de f",
        "transferencias automaticas por coparticipacion leyes compl y compensac acumulado anual en millones de f",
    ),
    ("nacion", "monto_millones"): (
        "transferencias automaticas por coparticipacion leyes compl y compensac en millones n",
    ),
    ("nacion", "variacion_nominal_pct"): (
        "transferencias automaticas por coparticipacion leyes compl y compensac var anual nominal n",
        "transferencias automaticas por coparticipacion leyes complementarias y compensaciones var anual nominal n",
    ),
    ("nacion", "variacion_real_pct"): (
        "transferencias automaticas por coparticipacion leyes complementarias y compensaciones var anual real n",
        "transferencias automaticas por coparticipacion leyes compl y compensac var anual real n",
    ),
    ("nacion", "acumulado_anual_millones"): (
        "transferencias automaticas por coparticipacion leyes complementarias y compensaciones acumulado anual en millones de n",
        "transferencias automaticas por coparticipacion leyes compl y compensac acumulado anual en millones de n",
    ),
}


def configurar_salida() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def normalizar(valor: object) -> str:
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower().replace("%", " ")
    return re.sub(r"[^a-z0-9]+", " ", texto).strip()


def seleccionar_json(ruta_indicada: str | None) -> Path:
    if ruta_indicada:
        ruta = Path(ruta_indicada).expanduser().resolve()
        if not ruta.is_file():
            raise FileNotFoundError(f"No existe el JSON indicado: {ruta}")
        return ruta
    archivos = list(CARPETA_REPORTES.glob("iaraf_????_??.json"))
    if not archivos:
        raise FileNotFoundError(f"No hay controles IARAF en: {CARPETA_REPORTES}")
    # El nombre YYYY_MM permite seleccionar por período, no por fecha de copia.
    return max(archivos, key=lambda ruta: ruta.name)


def cargar_control(ruta: Path) -> dict:
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    if datos.get("control", {}).get("estado") != "OK":
        raise ValueError("El control de extracción no tiene estado OK.")
    if not re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", str(datos.get("periodo", ""))):
        raise ValueError("El período del JSON no tiene formato YYYY-MM.")
    for jurisdiccion, campo in CAMPOS:
        valor = datos.get(jurisdiccion, {}).get(campo)
        if not isinstance(valor, (int, float)):
            raise ValueError(f"Falta un valor numérico: {jurisdiccion}.{campo}")
    return datos


def buscar_encabezados(hoja) -> tuple[int, dict[tuple[str, str], int], int]:
    mejor: tuple[int, dict[tuple[str, str], int], int] | None = None
    limite_filas = min(20, hoja.max_row)
    for fila in range(1, limite_filas + 1):
        textos = {col: normalizar(hoja.cell(fila, col).value) for col in range(1, hoja.max_column + 1)}
        columnas: dict[tuple[str, str], int] = {}
        for clave, variantes in CAMPOS.items():
            coincidencias = [
                col for col, texto in textos.items()
                if any(texto == variante or variante in texto for variante in variantes)
            ]
            if len(coincidencias) == 1:
                columnas[clave] = coincidencias[0]
            elif len(coincidencias) > 1:
                raise ValueError(f"Encabezado ambiguo para {clave}: columnas {coincidencias}")
        periodo = next((col for col, texto in textos.items() if texto in {"periodo", "periodo mes"}), 0)
        candidato = (fila, columnas, periodo)
        if mejor is None or len(columnas) > len(mejor[1]):
            mejor = candidato
    assert mejor is not None
    fila, columnas, periodo = mejor
    faltantes = [f"{j}.{c}" for (j, c) in CAMPOS if (j, c) not in columnas]
    if faltantes:
        raise ValueError("No se encontraron estos encabezados: " + ", ".join(faltantes))
    if not periodo:
        raise ValueError("No se encontró la columna PERIODO en la fila de encabezados.")
    return fila, columnas, periodo


def convertir_periodo(valor: object) -> tuple[int, int] | None:
    if isinstance(valor, (datetime, date)):
        return valor.year, valor.month
    if isinstance(valor, (int, float)) and 1900 <= valor <= 2200:
        return int(valor), 1
    texto = normalizar(valor)
    if not texto:
        return None
    m = re.fullmatch(r"(20\d{2})[- /](0?[1-9]|1[0-2])", texto)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.fullmatch(r"(0?[1-9]|1[0-2])[- /](20\d{2})", texto)
    if m:
        return int(m.group(2)), int(m.group(1))
    m = re.fullmatch(r"([a-z]+)[ -]?(\d{2}|20\d{2})", texto)
    if m and m.group(1) in MESES:
        anio = int(m.group(2))
        if anio < 100:
            anio += 2000
        return anio, MESES[m.group(1)]
    m = re.fullmatch(r"(\d{2}|20\d{2})[ -]?([a-z]+)", texto)
    if m and m.group(2) in MESES:
        anio = int(m.group(1))
        if anio < 100:
            anio += 2000
        return anio, MESES[m.group(2)]
    return None


def buscar_fila_periodo(hoja, fila_encabezados: int, columna: int, periodo: str) -> int:
    objetivo = tuple(map(int, periodo.split("-")))
    filas = [
        fila for fila in range(fila_encabezados + 1, hoja.max_row + 1)
        if convertir_periodo(hoja.cell(fila, columna).value) == objetivo
    ]
    if not filas:
        raise ValueError(f"No se encontró el período {periodo} en la columna PERIODO.")
    if len(filas) > 1:
        raise ValueError(f"El período {periodo} está repetido en las filas {filas}.")
    return filas[0]


def es_formula(valor: object) -> bool:
    return isinstance(valor, str) and valor.startswith("=")


def son_iguales(actual: object, nuevo: float, tolerancia: float = 1e-7) -> bool:
    return isinstance(actual, (int, float)) and abs(float(actual) - float(nuevo)) <= tolerancia


def preparar_cambios(hoja, fila: int, columnas: dict, datos: dict, sobrescribir: bool) -> list[dict]:
    cambios = []
    for (jurisdiccion, campo), columna in columnas.items():
        celda = hoja.cell(fila, columna)
        actual = celda.value
        nuevo = float(datos[jurisdiccion][campo])
        if es_formula(actual):
            # Los acumulados anuales del observatorio se calculan en Excel.
            # Se resuelven por título para validar la estructura, pero se conservan.
            continue
        if son_iguales(actual, nuevo):
            continue
        vacio_o_cero = actual is None or actual == "" or (
            isinstance(actual, (int, float)) and float(actual) == 0.0
        )
        # El informe del período es la fuente: si el valor guardado difiere
        # (p. ej. una extracción anterior equivocada), se corrige y se informa.
        cambios.append({
            "corrige_valor_anterior": not vacio_o_cero,
            "celda": celda,
            "jurisdiccion": jurisdiccion,
            "campo": campo,
            "anterior": actual,
            "nuevo": nuevo,
        })
    return cambios


def guardar_atomico(libro, ruta_excel: Path) -> None:
    descriptor, nombre_temporal = tempfile.mkstemp(
        prefix=f"{ruta_excel.stem}_", suffix=ruta_excel.suffix, dir=ruta_excel.parent
    )
    os.close(descriptor)
    temporal = Path(nombre_temporal)
    try:
        libro.save(temporal)
        os.replace(temporal, ruta_excel)
    finally:
        if temporal.exists():
            temporal.unlink()


def main() -> int:
    configurar_salida()
    parser = argparse.ArgumentParser(description="Actualiza datos fuente IARAF en el Excel base.")
    parser.add_argument("--json", help="Control JSON específico; por defecto usa el período más reciente.")
    parser.add_argument("--excel", help="Libro Excel; por defecto usa la base principal.")
    parser.add_argument("--hoja", default=HOJA_PREDETERMINADA, help="Hoja a actualizar.")
    parser.add_argument("--sobrescribir", action="store_true", help="Reemplaza valores no vacíos distintos.")
    args = parser.parse_args()

    ruta_json = seleccionar_json(args.json)
    ruta_excel = Path(args.excel).expanduser().resolve() if args.excel else EXCEL_PREDETERMINADO
    if not ruta_excel.is_file():
        raise FileNotFoundError(f"No existe el Excel: {ruta_excel}")
    datos = cargar_control(ruta_json)

    libro = load_workbook(ruta_excel, data_only=False, keep_links=True)
    if args.hoja not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {args.hoja!r}. Disponibles: {', '.join(libro.sheetnames)}")
    hoja = libro[args.hoja]
    fila_encabezados, columnas, columna_periodo = buscar_encabezados(hoja)
    fila = buscar_fila_periodo(hoja, fila_encabezados, columna_periodo, datos["periodo"])
    cambios = preparar_cambios(hoja, fila, columnas, datos, args.sobrescribir)

    print("=" * 72)
    print("ACTUALIZACIÓN DE BASE — IARAF")
    print("=" * 72)
    print(f"Período: {datos['periodo']} | hoja: {args.hoja} | fila: {fila}")
    print(f"Control fuente: {ruta_json}")
    if not cambios:
        print("RESULTADO: SIN CAMBIOS — los datos fuente ya coinciden y las fórmulas se conservaron.")
        print("=" * 72)
        return 0

    for cambio in cambios:
        celda = cambio["celda"]
        celda.value = cambio["nuevo"]
        # Conserva el formato existente; si estaba vacío, toma el de la fila anterior.
        if not celda.number_format or celda.number_format == "General":
            anterior = hoja.cell(celda.row - 1, celda.column)
            celda.number_format = copy(anterior.number_format)

    try:
        guardar_atomico(libro, ruta_excel)
    except PermissionError as exc:
        raise PermissionError("No se pudo guardar el Excel. Cerralo en Excel y volvé a ejecutar.") from exc

    # Verificación independiente después de guardar.
    control = load_workbook(ruta_excel, data_only=False, read_only=True, keep_links=True)
    hoja_control = control[args.hoja]
    for cambio in cambios:
        valor = hoja_control[cambio["celda"].coordinate].value
        if not son_iguales(valor, cambio["nuevo"]):
            raise RuntimeError(f"Falló la verificación de {cambio['celda'].coordinate}.")

    for cambio in cambios:
        celda = cambio["celda"]
        etiqueta = f"{cambio['jurisdiccion'].upper()} {cambio['campo']}"
        print(f"{celda.coordinate} | {etiqueta}: {cambio['anterior']!r} -> {cambio['nuevo']}")
    print(f"Celdas actualizadas: {len(cambios)}")
    print(f"EXCEL_IARAF={ruta_excel}")
    print("RESULTADO: OK")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        configurar_salida()
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
