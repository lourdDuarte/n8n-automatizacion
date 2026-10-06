"""Carga de un bloque mensual en el Excel local: datos + columnas calculadas.

Se usa en las fuentes nuevas (SRT y combustibles). Para cada mes:

* ubica la fila por la columna PERIODO (no por número de fila);
* escribe los datos originales en las columnas indicadas por título;
* en las columnas calculadas del bloque copia la fórmula del último mes que
  la tenga, trasladada a la fila nueva (igual que "arrastrar" en Excel);
* devuelve las cargas para el Sheet (datos + fórmulas) con el número de fila
  local, que luego se traduce a la ubicación real del Sheet.

Nunca reemplaza un dato existente distinto de cero salvo que se pida
explícitamente; en ese caso informa la corrección.
"""
from __future__ import annotations

import shutil
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from columnas_por_titulo import normalizar_titulo

HOJA = 'BASE DE DATOS'
MESES = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')


def _es_formula(v) -> bool:
    return isinstance(v, str) and v.startswith('=')


def _vacio_o_cero(v) -> bool:
    return v in (None, '') or (isinstance(v, (int, float)) and not isinstance(v, bool) and v == 0)


def _periodo_celda(v):
    if isinstance(v, (datetime, date)):
        return (v.year, v.month)
    return None


def _indice_columnas(ws) -> dict[str, list[int]]:
    indice = {}
    for celda in next(ws.iter_rows(min_row=1, max_row=1)):
        t = normalizar_titulo(celda.value)
        if t:
            indice.setdefault(t, []).append(celda.column)
    return indice


def columna(indice, titulo) -> int:
    cols = indice.get(normalizar_titulo(titulo), [])
    if len(cols) != 1:
        raise ValueError(f'Encabezado {titulo!r}: {len(cols)} coincidencias en el Excel local')
    return cols[0]


def filas_por_periodo(ws) -> dict[tuple, int]:
    salida = {}
    for fila in range(2, ws.max_row + 1):
        p = _periodo_celda(ws.cell(fila, 1).value)
        if p:
            if p in salida:
                raise ValueError(f'El período {p[0]}-{p[1]:02d} está repetido en la columna A')
            salida[p] = fila
    return salida


def ultimo_periodo_cargado(ruta: Path, titulo_control: str):
    libro = load_workbook(ruta, read_only=False, data_only=False)
    try:
        ws = libro[HOJA]
        col = columna(_indice_columnas(ws), titulo_control)
        filas = filas_por_periodo(ws)
        cargados = [p for p, f in filas.items() if not _vacio_o_cero(ws.cell(f, col).value)]
        return max(cargados) if cargados else None
    finally:
        libro.close()


def cargar_meses(ruta: Path, meses: dict[tuple, dict[str, float]], columnas_calculadas: list[str],
                 sobrescribir: bool = False, etiqueta_respaldo: str = 'bloque') -> dict:
    """meses: {(año, mes): {titulo_dato: valor}}. Devuelve cambios y cargas para el Sheet."""
    libro = load_workbook(ruta, data_only=False, keep_links=True)
    ws = libro[HOJA]
    indice = _indice_columnas(ws)
    filas = filas_por_periodo(ws)
    cols_calc = [columna(indice, t) for t in columnas_calculadas]
    cambios, cargas = [], []
    for periodo in sorted(meses):
        if periodo not in filas:
            raise ValueError(f'No existe la fila del período {periodo[0]}-{periodo[1]:02d} en el Excel local')
        fila = filas[periodo]
        valores_sheet = {}
        for titulo, valor in meses[periodo].items():
            col = columna(indice, titulo)
            celda = ws.cell(fila, col)
            if _es_formula(celda.value):
                raise ValueError(f'{celda.coordinate} tiene una fórmula; no se reemplaza por un dato')
            if isinstance(celda.value, (int, float)) and abs(float(celda.value) - float(valor)) < 1e-9:
                pass
            elif _vacio_o_cero(celda.value) or sobrescribir:
                cambios.append({'celda': celda.coordinate, 'titulo': str(titulo), 'antes': celda.value, 'despues': valor})
                celda.value = valor
            else:
                raise ValueError(f'{celda.coordinate} ya tiene {celda.value!r} y la fuente dice {valor}; no se pisa sin revisar')
            valores_sheet[str(ws.cell(1, col).value)] = valor
        for col in cols_calc:
            celda = ws.cell(fila, col)
            if not _es_formula(celda.value):
                # Se copia la fórmula del mes más cercano que la tenga (primero hacia atrás).
                cercanas = [f for d in range(1, 40) for f in (fila - d, fila + d)
                            if 1 < f <= ws.max_row and _es_formula(ws.cell(f, col).value)]
                origen = ws.cell(cercanas[0], col) if cercanas else None
                if origen is None:
                    raise ValueError(f'No hay una fórmula cercana para copiar en {celda.coordinate}')
                if not _vacio_o_cero(celda.value) and not sobrescribir:
                    raise ValueError(f'{celda.coordinate} tiene un dato manual; no se reemplaza por fórmula')
                formula = Translator(origen.value, origin=origen.coordinate).translate_formula(celda.coordinate)
                cambios.append({'celda': celda.coordinate, 'titulo': str(ws.cell(1, col).value), 'antes': celda.value, 'despues': formula})
                celda.value = formula
            valores_sheet[str(ws.cell(1, col).value)] = celda.value
        cargas.append({'hoja': HOJA, 'row_number': fila, 'periodo': f'{periodo[0]}-{periodo[1]:02d}', 'valores': valores_sheet})
    respaldo = None
    if cambios:
        carpeta = ruta.parent / 'respaldos'
        carpeta.mkdir(parents=True, exist_ok=True)
        respaldo = carpeta / f'antes_{etiqueta_respaldo}_{datetime.now():%Y%m%d_%H%M%S}.xlsx'
        shutil.copy2(ruta, respaldo)
        libro.calculation.fullCalcOnLoad = True
        libro.save(ruta)
    libro.close()
    return {'celdas_actualizadas': len(cambios), 'hubo_cambios': bool(cambios), 'cambios': cambios,
            'cargas_google': cargas, 'respaldo': str(respaldo) if respaldo else None}


def texto_periodo(p) -> str:
    return f'{MESES[p[1] - 1]} {p[0] % 100:02d}'
