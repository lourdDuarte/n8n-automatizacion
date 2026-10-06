"""Mapea actividad pecuaria por encabezado y conserva cálculos existentes."""

from __future__ import annotations

import math
import re
import shutil
import unicodedata
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook


HOJA = "BASE DE DATOS - Indicadores Pro"
MAPEO = {
    "stock_bovino_formosa": ("stock", "fsa", "bovinos"),
    "stock_bovino_nacion": ("stock", "nacion", "bovinos"),
    "faena_bovina_formosa": ("faena", "fsa", "bovinos"),
    "faena_bovina_nacion": ("faena", "nacion", "bovinos"),
    "consumo_bovino": ("consumo", "capita", "bovinos"),
    "stock_porcino_formosa": ("stock", "fsa", "porcinos"),
    "stock_porcino_nacion": ("stock", "n", "porcinos"),
    "faena_porcina_nacion": ("faena", "total", "porcinos"),
    "consumo_porcino": ("consumo", "capita", "porcinos"),
    "huevos_industria": ("produccion", "industria", "huevos"),
    "produccion_carne_aviar": ("produccion", "carne", "tn"),
    "faena_aviar": ("faena", "miles", "aves"),
    "consumo_aviar": ("consumo", "capita", "aves"),
}


def normalizar(valor) -> str:
    t = unicodedata.normalize("NFKD", str(valor or ""))
    t = "".join(c for c in t if not unicodedata.combining(c)).casefold()
    return " ".join(re.sub(r"[^a-z0-9]+", " ", t).split())


def localizar_columnas(ws) -> tuple[int, dict[str, tuple[int, str]]]:
    fila = max(range(1, min(ws.max_row, 10) + 1), key=lambda r: sum(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)))
    encabezados = {c: str(ws.cell(fila, c).value or "") for c in range(1, ws.max_column + 1)}
    resultado = {}
    for campo, tokens in MAPEO.items():
        def coincide(titulo: str) -> bool:
            limpio = normalizar(titulo); palabras = set(limpio.split())
            if "nacion" in campo and "fsa" in palabras: return False
            return all((token in palabras if len(token) == 1 else token in limpio) for token in tokens)
        hallados = [(c, t) for c, t in encabezados.items() if coincide(t)]
        if len(hallados) == 1:
            resultado[campo] = hallados[0]
        elif campo == "produccion_carne_aviar" and not hallados:
            # Esta variable figura en Libroproductivo.xlsx, pero no estaba en la
            # copia anterior de la hoja del Observatorio. No se altera el esquema.
            continue
        else:
            raise ValueError(f"No se encontró un encabezado único para {campo}: {[x[1] for x in hallados]}")
    return fila, resultado


def periodos(ws, columna_periodo: int) -> dict[tuple[int, int], int]:
    salida = {}
    for fila in range(1, ws.max_row + 1):
        valor = ws.cell(fila, columna_periodo).value
        if isinstance(valor, datetime):
            # La tabla productiva histórica codifica el año en el día:
            # 20/01/2025 se muestra como ene-20; 26/08/2026, como ago-26.
            if 20 <= valor.day <= 99:
                salida[(2000 + valor.day, valor.month)] = fila
            else:
                salida[(valor.year, valor.month)] = fila
        elif isinstance(valor, date):
            salida[(valor.year, valor.month)] = fila
        elif isinstance(valor, str):
            limpio = normalizar(valor)
            meses = {
                "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
                "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10,
                "nov": 11, "dic": 12,
            }
            coincidencia = re.search(r"([a-z]+)\s+(\d{2,4})", limpio)
            if coincidencia:
                mes = meses.get(coincidencia.group(1)[:4], meses.get(coincidencia.group(1)[:3]))
                anio = int(coincidencia.group(2))
                if mes:
                    salida[((2000 + anio) if anio < 100 else anio, mes)] = fila
    return salida


def iguales(a, b) -> bool:
    return isinstance(a, (int, float)) and isinstance(b, (int, float)) and math.isclose(float(a), float(b), abs_tol=1e-6, rel_tol=0) or a == b


def numero_local(valor) -> float:
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    texto = str(valor).strip().replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    return float(texto)


def asegurar_periodo(ws, columna_periodo: int, anio: int, mes: int) -> int:
    """Crea un mes faltante desplazando solo el bloque pecuario."""
    actuales = periodos(ws, columna_periodo)
    if (anio, mes) in actuales:
        return actuales[(anio, mes)]

    posteriores = sorted((clave, fila) for clave, fila in actuales.items() if clave > (anio, mes))
    if posteriores:
        fila_nueva = posteriores[0][1]
        ultima = max(actuales.values())
        for fila in range(ultima, fila_nueva - 1, -1):
            for columna in range(columna_periodo, ws.max_column + 1):
                origen = ws.cell(fila, columna)
                destino = ws.cell(fila + 1, columna)
                destino.value = origen.value
                if origen.has_style:
                    destino._style = copy(origen._style)
                destino.number_format = origen.number_format
    else:
        fila_nueva = max(actuales.values(), default=2) + 1

    modelo = ws.cell(fila_nueva + 1, columna_periodo).value
    if not isinstance(modelo, datetime):
        modelo = ws.cell(fila_nueva - 1, columna_periodo).value
    if isinstance(modelo, datetime) and 20 <= modelo.day <= 99:
        ws.cell(fila_nueva, columna_periodo).value = datetime(modelo.year, mes, anio - 2000)
    else:
        ws.cell(fila_nueva, columna_periodo).value = datetime(anio, mes, 1)
    for columna in range(columna_periodo + 1, ws.max_column + 1):
        ws.cell(fila_nueva, columna).value = None
    return fila_nueva


def actualizar(ruta: Path, registros: list[dict]) -> dict:
    wb = load_workbook(ruta, data_only=False)
    nombre = next((n for n in wb.sheetnames if normalizar(n) == normalizar(HOJA)), None)
    if not nombre:
        raise ValueError(f"No existe la hoja {HOJA}")
    ws = wb[nombre]
    fila_titulos, columnas = localizar_columnas(ws)
    col_periodo = next(c for c in range(1, ws.max_column + 1) if normalizar(ws.cell(fila_titulos, c).value) == "periodo")
    filas = periodos(ws, col_periodo)
    faltantes = sorted({(r["periodo"].year, r["periodo"].month) for r in registros} - set(filas))
    for anio, mes in faltantes:
        asegurar_periodo(ws, col_periodo, anio, mes)
    filas = periodos(ws, col_periodo)
    cambios, filas_tocadas = [], set()
    for registro in registros:
        p = registro["periodo"]
        clave = (p.year, p.month)
        if clave not in filas:
            raise ValueError(f"No existe el período {p:%Y-%m} en la tabla pecuaria")
        fila = filas[clave]
        for campo, valor in registro.items():
            if campo in {"periodo", "faena_bovina_formosa_acumulada"} or campo not in columnas or valor is None:
                continue
            col, titulo = columnas[campo]
            celda = ws.cell(fila, col)
            if iguales(celda.value, valor):
                continue
            cambios.append({"celda": celda.coordinate, "titulo": titulo, "antes": celda.value, "despues": valor})
            celda.value = valor
            filas_tocadas.add(fila)

        # El consumo total de proteína es cálculo: bovino + porcino + aviar.
        col_total = next((c for c in range(1, ws.max_column + 1) if all(t in normalizar(ws.cell(fila_titulos, c).value) for t in ("consumo", "total", "proteina"))), None)
        if (col_total and any(k in registro for k in ("consumo_bovino", "consumo_porcino", "consumo_aviar"))
                and all(k in columnas for k in ("consumo_bovino", "consumo_porcino", "consumo_aviar"))):
            refs = [ws.cell(fila, columnas[k][0]).coordinate for k in ("consumo_bovino", "consumo_porcino", "consumo_aviar")]
            if any(ws[ref].value in (None, "") for ref in refs):
                continue
            formula = "=SUM(" + ",".join(refs) + ")"
            celda = ws.cell(fila, col_total)
            if celda.value != formula:
                anterior = celda.value
                celda.value = formula
                cambios.append({"celda": celda.coordinate, "titulo": ws.cell(fila_titulos, col_total).value, "antes": anterior, "despues": formula})
                filas_tocadas.add(fila)
    respaldo = None
    if cambios:
        carpeta = ruta.parent / "respaldos"; carpeta.mkdir(parents=True, exist_ok=True)
        respaldo = carpeta / f"antes_pecuaria_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(ruta, respaldo)
        wb.calculation.fullCalcOnLoad = True; wb.calculation.forceFullCalc = True; wb.calculation.calcMode = "auto"
        wb.save(ruta)
    periodos_por_fila = {str(f): f"{a:04d}-{m:02d}-01" for (a, m), f in filas.items() if f in filas_tocadas}
    return {"hoja": nombre, "celdas_actualizadas": len(cambios), "hubo_cambios": bool(cambios), "cambios": cambios,
            "filas_localizadas": sorted(filas_tocadas), "fila_localizada": max(filas_tocadas) if filas_tocadas else None,
            "periodos_por_fila": periodos_por_fila,
            "respaldo": str(respaldo) if respaldo else None}
