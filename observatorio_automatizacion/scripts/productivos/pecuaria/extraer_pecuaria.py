"""Lectores tolerantes para las planillas y PDF pecuarios de SAGyP."""

from __future__ import annotations

import io
import math
import re
import unicodedata
from datetime import date, datetime

import fitz
import xlrd
from openpyxl import load_workbook


MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
         "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
         "noviembre": 11, "diciembre": 12}


def normalizar(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c)).casefold()
    return " ".join(re.sub(r"[^a-z0-9]+", " ", texto).split())


def numero(valor) -> float:
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        n = float(valor)
    else:
        texto = str(valor or "").strip().replace("\u00a0", "").replace(" ", "")
        texto = re.sub(r"[^0-9,.-]", "", texto)
        if not texto or texto in {"-", ".", ","}:
            raise ValueError(f"No es número: {valor!r}")
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"-?\d{1,3}\.\d{3}", texto):
            texto = texto.replace(".", "")
        elif texto.count(".") > 1:
            texto = texto.replace(".", "")
        n = float(texto)
    if not math.isfinite(n):
        raise ValueError(f"Número inválido: {valor!r}")
    return n


def periodo_en_texto(texto: str) -> date | None:
    t = normalizar(texto)
    anios = [int(x) for x in re.findall(r"\b20\d{2}\b", t)]
    for nombre, mes in MESES.items():
        if nombre in t and anios:
            return date(max(anios), mes, 1)
    return None


def texto_pdf(contenido: bytes) -> str:
    doc = fitz.open(stream=contenido, filetype="pdf")
    try:
        return "\n".join(pagina.get_text("text") for pagina in doc)
    finally:
        doc.close()


def filas_excel(contenido: bytes):
    wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            for fila in ws.iter_rows(values_only=True):
                yield ws.title, list(fila)
    finally:
        wb.close()


def extraer_stock_bovino(contenido: bytes) -> dict:
    wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    hojas = [s for s in wb.sheetnames if re.fullmatch(r"20\d{2}", str(s))]
    if not hojas: raise ValueError("No se reconoció la hoja anual del stock bovino")
    anio = max(map(int, hojas)); ws = wb[str(anio)]; formosa = nacional = None
    for fila in ws.iter_rows(values_only=True):
        nombre = normalizar(fila[0] if fila else None)
        if nombre == "formosa": formosa = numero(fila[-1])
        if nombre in {"total", "total general", "total pais"}: nacional = numero(fila[-1])
    wb.close()
    if formosa is None or nacional is None:
        raise ValueError(f"Stock bovino incompleto para {anio}: Formosa={formosa}, Nación={nacional}")
    return {"periodo": date(anio, 12, 1), "stock_bovino_formosa": formosa, "stock_bovino_nacion": nacional}


def extraer_faena_bovina(pdf: bytes, xls: bytes) -> dict:
    texto = texto_pdf(pdf)
    periodo = periodo_en_texto(texto[:500])
    if not periodo: raise ValueError("No se reconoció el período del informe bovino")
    libro = xlrd.open_workbook(file_contents=xls)
    tabla = libro.sheet_by_name("Faena por categoria")
    mes_nombre = next(k for k, v in MESES.items() if v == periodo.month)
    fila_acum = next(
        (r for r in range(tabla.nrows)
         if normalizar(tabla.cell_value(r, 0)) == f"acum {periodo.year}"),
        None,
    )
    if fila_acum is None:
        raise ValueError(f"No se encontró el bloque {periodo.year} en la planilla bovina")
    fila_mes = next(
        (r for r in range(fila_acum - 1, -1, -1)
         if normalizar(tabla.cell_value(r, 0)) == mes_nombre),
        None,
    )
    if fila_mes is None:
        raise ValueError(f"No se encontró {mes_nombre} {periodo.year} en la planilla bovina")
    nacional = numero(tabla.cell_value(fila_mes, 16))
    with fitz.open(stream=pdf, filetype="pdf") as documento:
        t = "\n".join(
            pagina.get_text("text") for pagina in documento
            if "FORMOSA" in pagina.get_text("text").upper()
        )
    if not t:
        raise ValueError("No se encontró la matriz provincial bovina")
    m = re.search(r"FORMOSA\s+([\d.]+).*?FORMOSA\s+([\d.]+)", t, re.S | re.I)
    if not m: raise ValueError("No se reconoció el acumulado de Formosa")
    return {"periodo": periodo, "faena_bovina_nacion": round(nacional),
            "faena_bovina_formosa_acumulada": round(numero(m.group(2)), 2)}


def extraer_indicadores_bovinos(contenido: bytes) -> list[dict]:
    """Faena nacional y consumo mensual desde la planilla oficial enlazada."""
    wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    salida = []
    try:
        for ws in wb.worksheets:
            for fila in ws.iter_rows(values_only=True):
                periodo = fila[1] if len(fila) > 1 else None
                if not isinstance(periodo, (date, datetime)) or len(fila) < 14:
                    continue
                if fila[2] is None or fila[12] is None:
                    continue
                salida.append({
                    "periodo": date(periodo.year, periodo.month, 1),
                    "faena_bovina_nacion": round(numero(fila[2]), 2),
                    "consumo_bovino": round(numero(fila[12]), 2),
                })
    finally:
        wb.close()
    if not salida:
        raise ValueError("No se reconocieron indicadores bovinos mensuales")
    return salida


def extraer_huevos_pdf(contenido: bytes) -> list[dict]:
    texto = texto_pdf(contenido)
    m_anios = re.search(r"Año\s+(20\d{2})\s+(20\d{2})", texto, re.I)
    if not m_anios: raise ValueError("No se reconocieron años en Industrialización de Huevos")
    anio = int(m_anios.group(2)); salida = []
    for nombre, mes in MESES.items():
        m = re.search(rf"^{nombre}\s+[\d.]+\s+([\d.]+)", texto, re.I | re.M)
        if m: salida.append({"periodo": date(anio, mes, 1), "huevos_industria": round(numero(m.group(1)))})
    if not salida: raise ValueError("No se reconocieron datos de huevos del año actual")
    return salida


def extraer_consumo_carnes(contenido: bytes) -> dict:
    texto = texto_pdf(contenido)
    m = re.search(r"Al\s+mes\s+de\s+([A-ZÁÉÍÓÚ]+)", texto, re.I)
    anio_m = re.search(r"\b(20\d{2})\b", texto[:300])
    if not m or not anio_m: raise ValueError("No se reconoció el período del consumo de carnes")
    mes = MESES[normalizar(m.group(1))]; anio = int(anio_m.group(1))
    # Primer gráfico: vacuna, porcina, aviar y total, en ese orden.
    tramo = texto[:2500]
    nums = [numero(x) for x in re.findall(r"\b\d{1,3},\d{2}\b", tramo)]
    terna = next((nums[i:i+3] for i in range(len(nums)-2) if 35 <= nums[i] <= 70 and 10 <= nums[i+1] <= 35 and 35 <= nums[i+2] <= 70), None)
    if not terna: raise ValueError("No se reconocieron los tres consumos per cápita")
    return {"periodo": date(anio, mes, 1), "consumo_bovino": terna[0], "consumo_porcino": terna[1], "consumo_aviar": terna[2]}


def extraer_aves_excel(contenido: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    salida = []
    try:
        for ws in wb.worksheets:
            anio = None
            for fila in ws.iter_rows(values_only=True):
                primero = fila[0] if fila else None
                if isinstance(primero, (int, float)) and int(primero) == primero and 2000 <= int(primero) <= 2100:
                    anio = int(primero); continue
                mes = MESES.get(normalizar(primero))
                if not mes or anio is None or len(fila) < 10:
                    continue
                if fila[1] is None or fila[2] is None or fila[8] is None:
                    continue
                salida.append({"periodo": date(anio, mes, 1), "faena_aviar": round(numero(fila[1])),
                               "produccion_carne_aviar": round(numero(fila[2])) * 1000,
                               "consumo_aviar": round(numero(fila[8]), 2)})
    finally:
        wb.close()
    if not salida:
        raise ValueError("No se reconocieron meses completos en Oferta y Demanda aviar")
    return salida


def extraer_porcinos_excel(contenido: bytes) -> list[dict]:
    wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    salida = []
    try:
        for ws in wb.worksheets:
            anio = None
            for fila in ws.iter_rows(values_only=True):
                primero = fila[0] if fila else None
                if isinstance(primero, (int, float)) and int(primero) == primero and 2000 <= int(primero) <= 2100:
                    anio = int(primero); continue
                abreviado = normalizar(primero)[:3]
                mes = next((n for nombre, n in MESES.items() if normalizar(nombre)[:3] == abreviado), None)
                if not mes or anio is None or len(fila) < 10 or fila[1] is None or fila[8] is None:
                    continue
                salida.append({"periodo": date(anio, mes, 1), "faena_porcina_nacion": round(numero(fila[1])),
                               "consumo_porcino": round(numero(fila[8]), 2)})
    finally:
        wb.close()
    if not salida:
        raise ValueError("No se reconocieron meses en Evolución Mensual Porcina")
    return salida


def extraer_stock_porcino(contenido: bytes) -> dict:
    wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    try:
        ws = wb.active
        titulo = " ".join(str(ws.cell(r, c).value or "") for r in range(1, 4) for c in range(1, min(ws.max_column, 15) + 1))
        anios = [int(x) for x in re.findall(r"20\d{2}", titulo)]
        if not anios: raise ValueError("No se encontró el año del stock porcino")
        anio = max(anios); formosa = nacion = None
        for fila in ws.iter_rows(values_only=True):
            nombre = normalizar(fila[0] if fila else None)
            if nombre == "formosa": formosa = numero(fila[-1])
            if nombre in {"total pais", "total nacional"}: nacion = numero(fila[-1])
        if formosa is None or nacion is None:
            raise ValueError(f"Stock porcino incompleto: Formosa={formosa}, Nación={nacion}")
        return {"periodo": date(anio, 3, 1), "stock_porcino_formosa": formosa, "stock_porcino_nacion": nacion}
    finally:
        wb.close()


def extraer_serie_pdf(contenido: bytes, campos: dict[str, tuple[str, ...]]) -> list[dict]:
    texto = texto_pdf(contenido)
    lineas = [" ".join(x.split()) for x in texto.splitlines() if x.strip()]
    resultados = []
    for i, linea in enumerate(lineas):
        periodo = periodo_en_texto(linea)
        if not periodo:
            continue
        entorno = " ".join(lineas[max(0, i - 3):i + 6])
        registro = {"periodo": periodo}
        for campo, rotulos in campos.items():
            if all(normalizar(r) in normalizar(entorno) for r in rotulos):
                nums = re.findall(r"-?\d[\d.]*,?\d*", entorno)
                candidatos = []
                for x in nums:
                    try: candidatos.append(numero(x))
                    except Exception: pass
                if candidatos:
                    registro[campo] = candidatos[-1]
        if len(registro) > 1:
            resultados.append(registro)
    return resultados


def combinar_por_periodo(*series: list[dict]) -> list[dict]:
    salida = {}
    for serie in series:
        for item in serie:
            periodo = item["periodo"]
            salida.setdefault(periodo, {"periodo": periodo}).update(item)
    return [salida[p] for p in sorted(salida)]


URL_TABLERO_FAENA = ("https://www.magyp.gob.ar/sitio/areas/bovinos/informacion_sectorial/_archivos//"
                     "000015_Tablero%20de%20faena%20bovina/000001_Tablero_Faena_Bovino.pdf")


def faena_nacion_tablero(timeout: int = 90) -> dict | None:
    """Faena bovina nacional del mes según el tablero mensual de SAGyP (Información sectorial → FAENA).

    Es el "Total de faena bovina" en cabezas que indica la guía del Observatorio
    (p. ej. agosto 2026: 1.011.334). Si no se puede leer, devuelve None y queda el
    valor de la planilla de faena."""
    import urllib.request
    try:
        req = urllib.request.Request(URL_TABLERO_FAENA, headers={"User-Agent": "Mozilla/5.0 (Observatorio Formosa/1.0)"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            contenido = r.read()
        texto = " ".join(texto_pdf(contenido).split())
        m_valor = re.search(r"Faena \(cabezas\)\s+([\d.]+)", texto, re.I)
        meses = [(int(a), MESES[normalizar(m)]) for m, a in
                 re.findall(r"\b(" + "|".join(MESES) + r")\s+(20\d{2})\b", texto, re.I)]
        if not m_valor or not meses:
            return None
        anio, mes = max(set(meses), key=meses.count)  # el mes del tablero es el que más se repite
        valor = int(m_valor.group(1).replace(".", ""))
        if not 300_000 <= valor <= 2_500_000:
            return None
        return {"periodo": date(anio, mes, 1), "faena_bovina_nacion": valor}
    except Exception:
        return None
