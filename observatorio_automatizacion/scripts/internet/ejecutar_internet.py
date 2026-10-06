"""Actualiza Accesos a internet (INDEC) en la base del Observatorio.

La publicación es trimestral, pero el archivo web contiene los tres meses
del trimestre. El proceso detecta automáticamente el último trimestre
disponible y actualiza Formosa y total país, incluidos los cálculos i.a.
"""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


RAIZ = Path(__file__).resolve().parents[2]
BASE = RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA = "BASE DE DATOS"
URL_SERIE = "https://www.indec.gob.ar/ftp/cuadros/sociedad/serie22_accesos_internet.xlsx"
URL_INFORME = "https://www.indec.gob.ar/uploads/informesdeprensa/internet_09_2629C1A6F505.pdf"

TITULOS = {
    "total_f": "Cantidad de accesos a internet F",
    "var_total_f": "Cantidad de accesos a internet - Var. Interanual (%) F",
    "total_n": "Cantidad de accesos a internet N",
    "var_total_n": "Cantidad de accesos a internet - Var. Interanual (%) N",
    "fijos_f": "Cantidad de accesos fijos a internet F",
    "var_fijos_f": "Cantidad de accesos fijos a internet - Var. Interanual (%) F",
    "fijos_n": "Cantidad de accesos fijos a internet N",
    "var_fijos_n": "Cantidad de accesos fijos a internet - Var. Interanual (%) N",
    "moviles_f": "Cantidad de accesos moviles a internet F",
    "var_moviles_f": "Cantidad de accesos moviles a internet - Var. Interanual (%) F",
    "moviles_n": "Cantidad de accesos moviles a internet N",
    "var_moviles_n": "Cantidad de accesos moviles a internet - Var. Interanual (%) N",
}

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def texto(valor) -> str:
    return str(valor or "").strip().casefold()


def descargar(destino: Path) -> None:
    solicitud = urllib.request.Request(URL_SERIE, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(solicitud, timeout=90) as respuesta, destino.open("wb") as salida:
        shutil.copyfileobj(respuesta, salida)
    if destino.stat().st_size < 10_000:
        raise RuntimeError("La serie descargada de INDEC está vacía o incompleta.")


def columnas_periodo(hoja, fila_anio: int, fila_mes: int) -> dict[tuple[int, int], int]:
    actual = None
    resultado = {}
    for columna in range(2, hoja.max_column + 1):
        valor_anio = hoja.cell(fila_anio, columna).value
        if isinstance(valor_anio, (int, float)):
            actual = int(valor_anio)
        elif texto(valor_anio).rstrip("*").isdigit():
            actual = int(texto(valor_anio).rstrip("*"))
        mes = MESES.get(texto(hoja.cell(fila_mes, columna).value))
        if actual and mes:
            resultado[(actual, mes)] = columna
    return resultado


def buscar_fila(hoja, etiqueta: str, desde: int, hasta: int | None = None) -> int:
    hasta = hasta or hoja.max_row
    buscada = etiqueta.casefold()
    for fila in range(desde, hasta + 1):
        if texto(hoja.cell(fila, 1).value) == buscada:
            return fila
    raise RuntimeError(f"No se encontró {etiqueta!r} en {hoja.title}.")


def extraer(ruta: Path) -> dict[tuple[int, int], dict[str, int]]:
    libro = load_workbook(ruta, read_only=True, data_only=True)
    try:
        jurisd = libro["Cuadro 4"]
        periodos_total = columnas_periodo(jurisd, 3, 4)
        fila_total_n = buscar_fila(jurisd, "Total del país", 5)
        fila_total_f = buscar_fila(jurisd, "Formosa", 5)

        tipos = libro["Cuadro 5"]
        fila_sec_mov = buscar_fila(tipos, "Accesos móviles", 1)
        fila_sec_fij = buscar_fila(tipos, "Accesos fijos", fila_sec_mov + 1)
        periodos_mov = columnas_periodo(tipos, fila_sec_mov + 1, fila_sec_mov + 2)
        periodos_fij = columnas_periodo(tipos, fila_sec_fij + 1, fila_sec_fij + 2)
        mov_n = buscar_fila(tipos, "Total del país", fila_sec_mov + 3, fila_sec_fij - 1)
        mov_f = buscar_fila(tipos, "Formosa", fila_sec_mov + 3, fila_sec_fij - 1)
        fij_n = buscar_fila(tipos, "Total del país", fila_sec_fij + 3)
        fij_f = buscar_fila(tipos, "Formosa", fila_sec_fij + 3)

        ultimo = max(set(periodos_total) & set(periodos_mov) & set(periodos_fij))
        anio, mes = ultimo
        inicio = ((mes - 1) // 3) * 3 + 1
        trimestre = [(anio, m) for m in range(inicio, inicio + 3)]
        datos = {}
        for periodo in trimestre:
            if periodo not in periodos_total or periodo not in periodos_mov or periodo not in periodos_fij:
                raise RuntimeError(f"INDEC no publicó completo el trimestre {periodo}.")
            ct, cm, cf = periodos_total[periodo], periodos_mov[periodo], periodos_fij[periodo]
            datos[periodo] = {
                "total_f": int(jurisd.cell(fila_total_f, ct).value),
                "total_n": int(jurisd.cell(fila_total_n, ct).value),
                "moviles_f": int(tipos.cell(mov_f, cm).value),
                "moviles_n": int(tipos.cell(mov_n, cm).value),
                "fijos_f": int(tipos.cell(fij_f, cf).value),
                "fijos_n": int(tipos.cell(fij_n, cf).value),
            }
        return datos
    finally:
        libro.close()


def actualizar(datos: dict[tuple[int, int], dict[str, int]]) -> dict:
    libro = load_workbook(BASE, data_only=False)
    hoja = libro[HOJA]
    encabezados = {str(hoja.cell(1, c).value or "").strip(): c for c in range(1, hoja.max_column + 1)}
    faltantes = [titulo for titulo in TITULOS.values() if titulo not in encabezados]
    if faltantes:
        libro.close()
        raise RuntimeError("Faltan encabezados en la base: " + "; ".join(faltantes))

    filas = {}
    for fila in range(2, hoja.max_row + 1):
        valor = hoja.cell(fila, 1).value
        if isinstance(valor, datetime):
            filas[(valor.year, valor.month)] = fila

    cambios = []
    periodos = []
    pares = (
        ("total_f", "var_total_f"), ("total_n", "var_total_n"),
        ("fijos_f", "var_fijos_f"), ("fijos_n", "var_fijos_n"),
        ("moviles_f", "var_moviles_f"), ("moviles_n", "var_moviles_n"),
    )
    for periodo in sorted(datos):
        fila = filas.get(periodo)
        fila_anterior = filas.get((periodo[0] - 1, periodo[1]))
        if not fila or not fila_anterior:
            raise RuntimeError(f"No se encontró la fila actual o interanual para {periodo}.")
        periodos.append(f"{periodo[0]:04d}-{periodo[1]:02d}")
        for clave_nivel, clave_var in pares:
            col_nivel = encabezados[TITULOS[clave_nivel]]
            col_var = encabezados[TITULOS[clave_var]]
            valor = datos[periodo][clave_nivel]
            formula = f"={hoja.cell(fila, col_nivel).coordinate}/{hoja.cell(fila_anterior, col_nivel).coordinate}*100-100"
            for columna, nuevo in ((col_nivel, valor), (col_var, formula)):
                celda = hoja.cell(fila, columna)
                anterior = celda.value
                if anterior != nuevo:
                    celda.value = nuevo
                    cambios.append({"celda": celda.coordinate, "titulo": hoja.cell(1, columna).value, "antes": anterior, "despues": nuevo})

    respaldo = None
    if cambios:
        carpeta = BASE.parent / "respaldos"
        carpeta.mkdir(parents=True, exist_ok=True)
        respaldo = carpeta / f"antes_internet_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(BASE, respaldo)
        libro.save(BASE)
    libro.close()
    return {"celdas_actualizadas": len(cambios), "hubo_cambios": bool(cambios), "cambios": cambios, "periodos": periodos, "respaldo": str(respaldo) if respaldo else None}


def main() -> int:
    try:
        with tempfile.TemporaryDirectory(prefix="indec_internet_") as temporal:
            descarga = Path(temporal) / "serie.xlsx"
            descargar(descarga)
            datos = extraer(descarga)
        actualizacion = actualizar(datos)
        ultimo = max(datos)
        resultado = {
            "fuente": "INDEC - Accesos a internet",
            "url": URL_INFORME,
            "serie_url": URL_SERIE,
            "consultado_en": datetime.now().astimezone().isoformat(timespec="seconds"),
            "periodo": f"{ultimo[0]:04d}-{ultimo[1]:02d}",
            "periodos": actualizacion["periodos"],
            "datos": {f"{a:04d}-{m:02d}": valores for (a, m), valores in datos.items()},
            "actualizacion": actualizacion,
        }
        reportes = RAIZ / "scripts" / "reportes" / "internet"
        reportes.mkdir(parents=True, exist_ok=True)
        reporte = reportes / f"internet_{ultimo[0]:04d}_{ultimo[1]:02d}.json"
        reporte.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
        resultado["reporte"] = str(reporte)
        print("RESULTADO_INTERNET=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
        print("ESTADO=OK")
        return 0
    except Exception as error:
        print(f"ERROR_INTERNET={error}", file=sys.stderr)
        print("ESTADO=ERROR")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
