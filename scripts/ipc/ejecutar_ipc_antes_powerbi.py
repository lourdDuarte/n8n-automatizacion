"""Detecta la ultima planilla mensual del IPC y actualiza PT:QY.

Dependencias: openpyxl, xlrd.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import load_workbook


URL_XLS = "https://www.indec.gob.ar/ftp/cuadros/economia/sh_ipc_{mes:02d}_{anio:02d}.xls"
HOJA_BASE = "BASE DE DATOS"
ARCHIVO_BASE = "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA_MENSUAL = "Variación mensual IPC Nacional"
HOJA_INTERANUAL = "Var. interanual IPC Nacional"
HOJA_INDICES = "Índices IPC Cobertura Nacional"
REGION = "Región Noreste"
DIVISIONES = (
    "Alimentos y bebidas no alcohólicas",
    "Bebidas alcohólicas y tabaco",
    "Prendas de vestir y calzado",
    "Vivienda, agua, electricidad, gas y otros combustibles",
    "Equipamiento y mantenimiento del hogar",
    "Salud",
    "Transporte",
    "Comunicación",
    "Recreación y cultura",
    "Educación",
    "Restaurantes y hoteles",
    "Bienes y servicios varios",
)

COLUMNAS_POWER_BI = (
    "periodo",
    "anio",
    "mes_numero",
    "codigo",
    "indicador",
    "territorio",
    "division",
    "tipo_medicion",
    "valor",
    "unidad",
    "fuente",
    "url_planilla",
    "consultado_en",
)

METADATOS_POWER_BI = {
    "PT": ("IPC nivel general", "Región NEA", "Nivel general", "Índice", "índice"),
    "PU": ("IPC nivel general", "Total nacional", "Nivel general", "Índice", "índice"),
    "PV": ("IPC nivel general", "Región NEA", "Nivel general", "Variación mensual", "%"),
    "PW": ("IPC nivel general", "Total nacional", "Nivel general", "Variación mensual", "%"),
    "PX": ("IPC nivel general", "Región NEA", "Nivel general", "Variación acumulada", "%"),
    "PY": ("IPC nivel general", "Total nacional", "Nivel general", "Variación acumulada", "%"),
    "PZ": ("IPC nivel general", "Región NEA", "Nivel general", "Variación interanual", "%"),
    "QA": ("IPC nivel general", "Total nacional", "Nivel general", "Variación interanual", "%"),
}

for codigo, division in zip(
    ("QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM"),
    DIVISIONES,
):
    METADATOS_POWER_BI[codigo] = (
        "IPC por división",
        "Región NEA",
        division,
        "Variación mensual",
        "%",
    )

for codigo, division in zip(
    ("QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY"),
    DIVISIONES,
):
    METADATOS_POWER_BI[codigo] = (
        "IPC por división",
        "Región NEA",
        division,
        "Variación interanual",
        "%",
    )


def raiz_proyecto() -> Path:
    return Path(__file__).resolve().parents[2]


def pedir(url: str, timeout: int = 60) -> bytes:
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (automatizacion-indicadores/1.0)"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as respuesta:
        contenido = respuesta.read()
        tipo = respuesta.headers.get_content_type()
    if not contenido.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        raise ValueError(f"La respuesta no es una planilla XLS ({tipo}): {url}")
    return contenido


def sumar_meses(valor: date, cantidad: int) -> date:
    indice = valor.year * 12 + valor.month - 1 + cantidad
    return date(indice // 12, indice % 12 + 1, 1)


def descubrir_planilla(hoy: date | None = None, meses_atras: int = 18) -> tuple[str, bytes]:
    """Prueba primero la edicion esperada y retrocede hasta hallar la ultima publicada."""
    hoy = hoy or date.today()
    for retroceso in range(meses_atras + 1):
        publicacion = sumar_meses(date(hoy.year, hoy.month, 1), -retroceso)
        url = URL_XLS.format(mes=publicacion.month, anio=publicacion.year % 100)
        try:
            return url, pedir(url)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, ValueError) as error:
            if isinstance(error, urllib.error.HTTPError) and error.code not in (403, 404):
                raise
            continue
    raise FileNotFoundError(f"No se encontro una planilla IPC en los ultimos {meses_atras + 1} meses")


def fecha_xls(libro: xlrd.book.Book, valor) -> date:
    if not isinstance(valor, (int, float)):
        raise ValueError(f"Encabezado de fecha XLS invalido: {valor!r}")
    partes = xlrd.xldate_as_tuple(valor, libro.datemode)
    return date(partes[0], partes[1], 1)


def normalizar(texto) -> str:
    return " ".join(str(texto).split()).casefold()


@dataclass(frozen=True)
class Bloque:
    hoja: xlrd.sheet.Sheet
    fila_fechas: int
    fila_nivel: int

    def fila(self, etiqueta: str) -> int:
        buscada = normalizar(etiqueta)
        for fila in range(self.fila_nivel, min(self.fila_nivel + 20, self.hoja.nrows)):
            if normalizar(self.hoja.cell_value(fila, 0)) == buscada:
                return fila
        raise ValueError(f"No se encontro {etiqueta!r} en {self.hoja.name}, bloque fila {self.fila_nivel + 1}")


def bloque(hoja: xlrd.sheet.Sheet, region: str) -> Bloque:
    buscada = normalizar(region)
    fila_region = next(
        (fila for fila in range(hoja.nrows) if normalizar(hoja.cell_value(fila, 0)) == buscada),
        None,
    )
    if fila_region is None:
        raise ValueError(f"No se encontro el bloque {region!r} en {hoja.name}")
    fila_nivel = next(
        (fila for fila in range(fila_region + 1, min(fila_region + 12, hoja.nrows))
         if normalizar(hoja.cell_value(fila, 0)) in {"nivel general", "nivel general "}),
        None,
    )
    if fila_nivel is None:
        raise ValueError(f"No se encontro Nivel general despues de {region!r} en {hoja.name}")
    return Bloque(hoja, fila_region, fila_nivel)


def columna_periodo(libro: xlrd.book.Book, datos: Bloque, periodo: date) -> int:
    for columna in range(1, datos.hoja.ncols):
        valor = datos.hoja.cell_value(datos.fila_fechas, columna)
        if valor in ("", None):
            continue
        try:
            if fecha_xls(libro, valor) == periodo:
                return columna
        except ValueError:
            continue
    raise ValueError(f"No existe {periodo:%Y-%m} en {datos.hoja.name}, bloque {datos.fila_fechas + 1}")


def ultimo_periodo(libro: xlrd.book.Book, datos: Bloque) -> date:
    fechas = []
    for columna in range(1, datos.hoja.ncols):
        valor = datos.hoja.cell_value(datos.fila_fechas, columna)
        if valor not in ("", None):
            try:
                fechas.append(fecha_xls(libro, valor))
            except ValueError:
                pass
    if not fechas:
        raise ValueError(f"No hay fechas en {datos.hoja.name}")
    return max(fechas)


def valor(libro: xlrd.book.Book, datos: Bloque, periodo: date, etiqueta: str) -> float:
    numero = datos.hoja.cell_value(datos.fila(etiqueta), columna_periodo(libro, datos, periodo))
    if not isinstance(numero, (int, float)) or not math.isfinite(float(numero)):
        raise ValueError(f"Valor invalido para {etiqueta}, {periodo:%Y-%m}, {datos.hoja.name}: {numero!r}")
    return float(numero)


def redondear_un_decimal(numero: float) -> float:
    # Los cuadros publican porcentajes con un decimal.
    return round(numero + 1e-12, 1)


def extraer(contenido: bytes, url: str) -> dict:
    libro = xlrd.open_workbook(file_contents=contenido)
    requeridas = {HOJA_MENSUAL, HOJA_INTERANUAL, HOJA_INDICES}
    faltantes = requeridas.difference(libro.sheet_names())
    if faltantes:
        raise ValueError(f"Faltan hojas requeridas: {sorted(faltantes)}")

    mensual_n = bloque(libro.sheet_by_name(HOJA_MENSUAL), "Total nacional")
    mensual_nea = bloque(libro.sheet_by_name(HOJA_MENSUAL), REGION)
    anual_n = bloque(libro.sheet_by_name(HOJA_INTERANUAL), "Total nacional")
    anual_nea = bloque(libro.sheet_by_name(HOJA_INTERANUAL), REGION)
    indice_n = bloque(libro.sheet_by_name(HOJA_INDICES), "Total nacional")
    indice_nea = bloque(libro.sheet_by_name(HOJA_INDICES), REGION)

    periodo = ultimo_periodo(libro, indice_n)
    periodos = {
        ultimo_periodo(libro, datos)
        for datos in (mensual_n, mensual_nea, anual_n, anual_nea, indice_n, indice_nea)
    }
    if periodos != {periodo}:
        raise ValueError(f"Las hojas o regiones no terminan en el mismo periodo: {sorted(periodos)}")

    diciembre = date(periodo.year - 1, 12, 1)
    indice_actual_n = valor(libro, indice_n, periodo, "Nivel general")
    indice_actual_nea = valor(libro, indice_nea, periodo, "Nivel general")
    acumulada_n = redondear_un_decimal((indice_actual_n / valor(libro, indice_n, diciembre, "Nivel general") - 1) * 100)
    acumulada_nea = redondear_un_decimal((indice_actual_nea / valor(libro, indice_nea, diciembre, "Nivel general") - 1) * 100)

    valores = {
        "PT": indice_actual_nea,
        "PU": indice_actual_n,
        "PV": valor(libro, mensual_nea, periodo, "Nivel general"),
        "PW": valor(libro, mensual_n, periodo, "Nivel general"),
        "PX": acumulada_nea,
        "PY": acumulada_n,
        "PZ": valor(libro, anual_nea, periodo, "Nivel general"),
        "QA": valor(libro, anual_n, periodo, "Nivel general"),
    }
    for columna, division in zip(
        ("QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM"),
        DIVISIONES,
    ):
        valores[columna] = valor(libro, mensual_nea, periodo, division)
    for columna, division in zip(
        ("QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY"),
        DIVISIONES,
    ):
        valores[columna] = valor(libro, anual_nea, periodo, division)

    return {
        "fuente": "INDEC - IPC",
        "periodo": periodo.isoformat(),
        "anio": periodo.year,
        "mes_numero": periodo.month,
        "url_planilla": url,
        "consultado_en": datetime.now().isoformat(timespec="seconds"),
        "valores": valores,
        "validaciones": {
            "hojas_y_regiones_mismo_periodo": True,
            "cantidad_valores": len(valores),
            "acumuladas_calculadas_desde_indices": True,
        },
    }


def localizar_fila(ws, anio: int, mes: int) -> int:
    for fila in range(2, ws.max_row + 1):
        valor_celda = ws.cell(fila, 1).value
        if isinstance(valor_celda, datetime) and valor_celda.year == anio and valor_celda.month == mes:
            return fila
        if isinstance(valor_celda, date) and valor_celda.year == anio and valor_celda.month == mes:
            return fila
    raise ValueError(f"No existe el periodo {anio}-{mes:02d} en la columna A")


def iguales(anterior, nuevo) -> bool:
    if isinstance(anterior, (int, float)) and isinstance(nuevo, (int, float)):
        return math.isclose(float(anterior), float(nuevo), rel_tol=0, abs_tol=1e-9)
    return anterior == nuevo


def actualizar_base(resultado: dict, ruta_base: Path) -> dict:
    if not ruta_base.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta_base}")
    libro = load_workbook(ruta_base, data_only=False)
    if HOJA_BASE not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {HOJA_BASE!r}")
    ws = libro[HOJA_BASE]
    fila = localizar_fila(ws, resultado["anio"], resultado["mes_numero"])
    cambios, coincidencias = [], []
    for columna, nuevo in resultado["valores"].items():
        celda = ws[f"{columna}{fila}"]
        anterior = celda.value
        if iguales(anterior, nuevo):
            coincidencias.append(celda.coordinate)
        else:
            celda.value = nuevo
            cambios.append({"celda": celda.coordinate, "antes": anterior, "despues": nuevo})
    if cambios:
        libro.save(ruta_base)
    return {
        "archivo_excel": str(ruta_base.resolve()),
        "fila": fila,
        "celdas_actualizadas": len(cambios),
        "cambios": cambios,
        "coincidencias": coincidencias,
    }


def guardar_reporte(resultado: dict, carpeta: Path) -> Path:
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / f"ipc_{resultado['anio']}_{resultado['mes_numero']:02d}.json"
    ruta.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    return ruta


def fila_power_bi(resultado: dict, codigo: str, valor_indicador: float) -> dict:
    indicador, territorio, division, tipo_medicion, unidad = METADATOS_POWER_BI[codigo]
    return {
        "periodo": resultado["periodo"],
        "anio": resultado["anio"],
        "mes_numero": resultado["mes_numero"],
        "codigo": codigo,
        "indicador": indicador,
        "territorio": territorio,
        "division": division,
        "tipo_medicion": tipo_medicion,
        "valor": valor_indicador,
        "unidad": unidad,
        "fuente": resultado["fuente"],
        "url_planilla": resultado["url_planilla"],
        "consultado_en": resultado["consultado_en"],
    }


def actualizar_csv_power_bi(resultado: dict, ruta_csv: Path) -> dict:
    """Inserta o reemplaza el período del IPC en una tabla larga para Power BI."""
    filas_existentes: list[dict] = []

    if ruta_csv.is_file():
        with ruta_csv.open("r", encoding="utf-8-sig", newline="") as archivo:
            lector = csv.DictReader(archivo)
            if lector.fieldnames != list(COLUMNAS_POWER_BI):
                raise ValueError(
                    "El CSV de Power BI tiene encabezados inesperados: "
                    f"{lector.fieldnames}"
                )
            filas_existentes = list(lector)

    periodo_actual = resultado["periodo"]
    filas_conservadas = [
        fila
        for fila in filas_existentes
        if not (
            fila.get("fuente") == resultado["fuente"]
            and fila.get("periodo") == periodo_actual
        )
    ]
    filas_nuevas = [
        fila_power_bi(resultado, codigo, valor_indicador)
        for codigo, valor_indicador in resultado["valores"].items()
    ]
    filas = filas_conservadas + filas_nuevas
    filas.sort(key=lambda fila: (fila["periodo"], fila["fuente"], fila["codigo"]))

    ruta_csv.parent.mkdir(parents=True, exist_ok=True)
    ruta_temporal = ruta_csv.with_suffix(".tmp")
    with ruta_temporal.open("w", encoding="utf-8-sig", newline="") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=COLUMNAS_POWER_BI)
        escritor.writeheader()
        escritor.writerows(filas)
    ruta_temporal.replace(ruta_csv)

    return {
        "archivo_csv": str(ruta_csv.resolve()),
        "filas_periodo": len(filas_nuevas),
        "filas_totales": len(filas),
    }


def prueba_parser(ruta_xls: Path) -> None:
    contenido = ruta_xls.read_bytes()
    resultado = extraer(contenido, ruta_xls.as_uri())
    assert resultado["validaciones"]["cantidad_valores"] == 32
    assert set(resultado["valores"]) == {
        "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "QA",
        "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM",
        "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY",
    }
    print(json.dumps(resultado, ensure_ascii=False, indent=2))
    print("PRUEBA_PARSER=OK")


def main() -> int:
    raiz = raiz_proyecto()
    parser = argparse.ArgumentParser(description="Detecta el ultimo IPC publicado y actualiza PT:QY.")
    parser.add_argument("--base", type=Path, default=raiz / "base" / ARCHIVO_BASE)
    parser.add_argument("--solo-extraer", action="store_true")
    parser.add_argument("--xls", type=Path, help="Usa una planilla local en vez de descargarla (pruebas).")
    parser.add_argument("--probar-parser", type=Path, metavar="ARCHIVO_XLS")
    args = parser.parse_args()
    try:
        if args.probar_parser:
            prueba_parser(args.probar_parser)
            return 0
        if args.xls:
            contenido, url = args.xls.read_bytes(), args.xls.resolve().as_uri()
        else:
            url, contenido = descubrir_planilla()
        resultado = extraer(contenido, url)
        if not args.solo_extraer:
            resultado["actualizacion"] = actualizar_base(resultado, args.base)
        reporte = guardar_reporte(resultado, raiz / "scripts" / "reportes" / "ipc")
        resultado["reporte"] = str(reporte.resolve())
        resultado["power_bi"] = actualizar_csv_power_bi(
            resultado,
            raiz / "powerbi" / "datos_observatorio.csv",
        )
        print(json.dumps(resultado, ensure_ascii=False, indent=2))
        print("ESTADO=OK")
        return 0
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())