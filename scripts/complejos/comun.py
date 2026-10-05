from __future__ import annotations

import json
import math
import shutil
import unicodedata
import urllib.request
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

HOJA_BASE = "BASE DE DATOS"
ARCHIVO_BASE = "BASE DE DATOS INDICADORES PROPIOS.xlsx"


def raiz_proyecto(desde: Path) -> Path:
    actual = desde.resolve()
    for carpeta in (actual, *actual.parents):
        if (carpeta / "base").is_dir() and (carpeta / "scripts").is_dir():
            return carpeta
    raise FileNotFoundError("No se pudo localizar la carpeta principal del proyecto")


def normalizar(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.casefold().replace("\n", " ").split())


def descargar(url: str, timeout: int = 90) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Observatorio Formosa/1.0)"})
    with urllib.request.urlopen(req, timeout=timeout) as respuesta:
        return respuesta.read()


def numero(valor) -> float:
    if isinstance(valor, (int, float)) and math.isfinite(float(valor)):
        return float(valor)
    texto = str(valor or "").strip().replace("%", "")
    if not texto or normalizar(texto) in {"s", "...", "…", "///", "-"}:
        return 0.0
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    return float(texto)


def abrir_base(ruta: Path):
    if not ruta.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta}")
    libro = load_workbook(ruta, data_only=False)
    if HOJA_BASE not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {HOJA_BASE!r}")
    return libro, libro[HOJA_BASE]


def fila_periodo(ws, anio: int, mes: int) -> int:
    for fila in range(2, ws.max_row + 1):
        valor = ws.cell(fila, 1).value
        if isinstance(valor, (datetime, date)) and valor.year == anio and valor.month == mes:
            return fila
    raise ValueError(f"No existe el período {anio}-{mes:02d} en la columna A")


def columnas_por_titulo(ws) -> dict[str, int]:
    resultado = {}
    for columna in range(1, ws.max_column + 1):
        titulo = ws.cell(1, columna).value
        if titulo not in (None, ""):
            clave = normalizar(titulo)
            if clave in resultado:
                raise ValueError(f"Título duplicado en la base: {titulo!r}")
            resultado[clave] = columna
    return resultado


def columna_titulo(ws, titulo: str) -> int:
    columnas = columnas_por_titulo(ws)
    clave = normalizar(titulo)
    if clave not in columnas:
        raise ValueError(f"No se encontró la columna titulada {titulo!r}")
    return columnas[clave]


def iguales(anterior, nuevo) -> bool:
    if isinstance(anterior, (int, float)) and isinstance(nuevo, (int, float)):
        return math.isclose(float(anterior), float(nuevo), rel_tol=0, abs_tol=1e-7)
    return anterior == nuevo


def escribir(ws, fila: int, titulo: str, nuevo, cambios: list, coincidencias: list) -> str:
    columna = columna_titulo(ws, titulo)
    celda = ws.cell(fila, columna)
    anterior = celda.value
    if iguales(anterior, nuevo):
        coincidencias.append(celda.coordinate)
    else:
        celda.value = nuevo
        cambios.append({"celda": celda.coordinate, "titulo": titulo, "antes": anterior, "despues": nuevo})
    return celda.coordinate


def respaldar_base(ruta: Path, etiqueta: str) -> Path:
    carpeta = ruta.parent / "respaldos"
    carpeta.mkdir(parents=True, exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = carpeta / f"antes_{etiqueta}_{sello}.xlsx"
    shutil.copy2(ruta, destino)
    return destino


def guardar_si_cambia(libro, ruta: Path, cambios: list, etiqueta: str) -> str | None:
    if not cambios:
        libro.close()
        return None
    respaldo = respaldar_base(ruta, etiqueta)
    libro.save(ruta)
    libro.close()
    return str(respaldo.resolve())


def guardar_reporte(raiz: Path, fuente: str, periodo: date, resultado: dict) -> Path:
    carpeta = raiz / "scripts" / "reportes" / fuente
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / f"{fuente}_{periodo.year}_{periodo.month:02d}.json"
    ruta.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    return ruta


def emitir(resultado: dict) -> None:
    print(json.dumps(resultado, ensure_ascii=False, indent=2))
    print("RESULTADO_FUENTE=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
    print("ESTADO=OK")
