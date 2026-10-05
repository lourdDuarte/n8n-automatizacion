"""Ejecutor independiente de indicadores productivos para n8n."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from openpyxl import load_workbook


RAIZ = Path(__file__).resolve().parent
ARCHIVO_EXCEL = RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA_DESTINO = "BASE DE DATOS - Indicadores Pro"


def serializar_json(valor):
    """Convierte tipos de Excel/Python a valores aceptados por JSON y n8n."""
    if isinstance(valor, (datetime, date)):
        return valor.isoformat()
    if isinstance(valor, Path):
        return str(valor)
    raise TypeError(f"Object of type {valor.__class__.__name__} is not JSON serializable")


def capturar_hoja() -> dict:
    libro = load_workbook(ARCHIVO_EXCEL, read_only=True, data_only=False)
    try:
        hoja = libro[HOJA_DESTINO]
        filas = list(hoja.iter_rows(values_only=True))
        if len(filas) < 2:
            return {"encabezados": {}, "valores": {}}
        # En Indicadores Productivos la fila 1 contiene los bloques generales
        # y los encabezados reales de las columnas se encuentran en la fila 2.
        return {
            "encabezados": {c: valor for c, valor in enumerate(filas[1], start=1)},
            "valores": {
                (f, c): valor
                for f, fila in enumerate(filas[2:], start=3)
                for c, valor in enumerate(fila, start=1)
            },
        }
    finally:
        libro.close()


def normalizar_periodo(valor) -> str:
    if isinstance(valor, datetime):
        meses = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
        return f"{meses[valor.month - 1]} {valor.year % 100:02d}"
    return str(valor).strip()


def normalizar_titulo(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c)).casefold()
    return " ".join(re.sub(r"[^a-z0-9]+", " ", texto).split())


ENCABEZADOS_PECUARIA = {
    "periodo",
    "stock fsa cabezas bovinos",
    "stock nacion cabezas bovinos",
    "faena fsa cabezas bovinos",
    "faena nacion cabezas bovinos",
    "consumo per capita n kg hab ano bovinos",
    "stock fsa cabezas porcinos",
    "stock n cabezas porcinos",
    "faena total n cabezas porcinos",
    "consumo per capita n kg hab ano porcinos",
    "produccion destino industria miles de huevos n aves",
    "produccion carne tn",
    "faena miles cab n aves",
    "consumo per capita n kg hab ano aves",
    "consumo total de proteina de origen animal kg n",
}
ENCABEZADO_PECUARIA_OPCIONAL = "produccion carne tn"


def columnas_pecuarias(encabezados: dict) -> list[int]:
    """Localiza el bloque por títulos; no depende de letras ni posiciones."""
    encontradas = {
        normalizar_titulo(titulo): columna
        for columna, titulo in encabezados.items()
        if normalizar_titulo(titulo) in ENCABEZADOS_PECUARIA
    }
    faltantes = (ENCABEZADOS_PECUARIA - {ENCABEZADO_PECUARIA_OPCIONAL}) - set(encontradas)
    if faltantes:
        raise ValueError(
            "Faltan encabezados pecuarios en la hoja: " + ", ".join(sorted(faltantes))
        )
    return [encontradas[titulo] for titulo in ENCABEZADOS_PECUARIA if titulo in encontradas]


def preparar_actualizaciones_google(antes: dict, despues: dict) -> list[dict]:
    encabezados = despues["encabezados"]
    filas = {}
    for f, c in set(antes["valores"]) | set(despues["valores"]):
        previo = antes["valores"].get((f, c))
        nuevo = despues["valores"].get((f, c))
        if previo == nuevo:
            continue
        titulo = encabezados.get(c)
        if titulo in (None, ""):
            continue
        item = filas.setdefault(
            f,
            {
                "hoja": HOJA_DESTINO,
                "row_number": f,
                "periodo": None,
                "valores": {},
            },
        )
        item["valores"][str(titulo)] = nuevo
    return list(filas.values())


def preparar_filas_actuales_google(resultados: list[dict], estado: dict) -> list[dict]:
    """Entrega la fila mensual completa para una sincronizacion idempotente."""
    encabezados = estado["encabezados"]
    valores = estado["valores"]
    salida = []
    filas_incluidas = set()

    for resultado in resultados:
        if resultado.get("estado") != "OK":
            continue
        candidatas = resultado.get("filas_localizadas") or [resultado.get("fila_localizada")]
        for fila in candidatas:
            if not isinstance(fila, int) or fila < 3 or fila in filas_incluidas:
                continue
            datos = {}
            columnas = columnas_pecuarias(encabezados) if resultado.get("fuente") == "pecuaria" else range(3, 19)
            for columna in columnas:
                titulo = encabezados.get(columna)
                valor = valores.get((fila, columna))
                if titulo not in (None, "") and valor is not None:
                    datos[str(titulo)] = valor
            if datos:
                periodo_fila = resultado.get("periodos_por_fila", {}).get(str(fila), resultado.get("periodo"))
                salida.append({"hoja": HOJA_DESTINO, "row_number": fila,
                               "periodo": normalizar_periodo(periodo_fila), "valores": datos})
                filas_incluidas.add(fila)

    return salida


@dataclass(frozen=True)
class Fuente:
    nombre: str
    ruta_relativa: str

    def ruta(self) -> Path:
        return RAIZ / self.ruta_relativa


FUENTES = {
    "comercializacion": Fuente(
        "Comercialización de granos",
        "scripts/productivos/comercializacion/ejecutar_comercializacion.py",
    ),
    "pecuaria": Fuente(
        "Actividad pecuaria",
        "scripts/productivos/pecuaria/ejecutar_pecuaria.py",
    ),
}


def seleccionar(texto: str) -> list[str]:
    texto = texto.strip().casefold()
    if texto in {"todas", "todo", "all"}:
        return list(FUENTES)
    seleccion = []
    for item in texto.split(","):
        clave = item.strip().casefold()
        if clave and clave not in seleccion:
            seleccion.append(clave)
    desconocidas = [x for x in seleccion if x not in FUENTES]
    if desconocidas:
        raise ValueError(
            f"Fuentes desconocidas: {', '.join(desconocidas)}. "
            f"Disponibles: {', '.join(FUENTES)}"
        )
    if not seleccion:
        raise ValueError("No se indicó ninguna fuente")
    return seleccion


def ejecutar(clave: str, fuente: Fuente) -> dict:
    ruta = fuente.ruta()
    if not ruta.is_file():
        return {"fuente": clave, "nombre": fuente.nombre, "estado": "ERROR", "detalle": f"No existe {ruta}"}
    entorno = os.environ.copy()
    entorno["PYTHONIOENCODING"] = "utf-8"
    entorno["PYTHONUTF8"] = "1"
    inicio = time.monotonic()
    try:
        proceso = subprocess.run(
            [sys.executable, "-u", str(ruta)],
            cwd=RAIZ,
            env=entorno,
            text=True,
            capture_output=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            timeout=600,
        )
    except subprocess.TimeoutExpired:
        return {
            "fuente": clave,
            "nombre": fuente.nombre,
            "estado": "ERROR",
            "codigo": 124,
            "duracion_segundos": round(time.monotonic() - inicio, 2),
            "celdas_actualizadas": 0,
            "hubo_cambios": False,
            "salida_final": "",
            "error_final": "La descarga o extracción superó el límite de 10 minutos",
        }
    salida = proceso.stdout.strip()
    error = proceso.stderr.strip()
    cambios = None
    periodo = None
    fila_localizada = None
    filas_localizadas = []
    periodos_por_fila = {}
    for linea in salida.splitlines():
        if linea.startswith(("RESULTADO_COMERCIALIZACION=", "RESULTADO_PECUARIA=")):
            try:
                detalle = json.loads(linea.split("=", 1)[1])
                periodo = detalle.get("periodo")
                actualizacion = detalle.get("actualizacion", {})
                cambios = actualizacion.get("celdas_actualizadas", 0)
                fila_localizada = actualizacion.get("fila_localizada")
                filas_localizadas = actualizacion.get("filas_localizadas", [])
                periodos_por_fila = actualizacion.get("periodos_por_fila", {})
            except json.JSONDecodeError:
                pass
    return {
        "fuente": clave,
        "nombre": fuente.nombre,
        "estado": "OK" if proceso.returncode == 0 else "ERROR",
        "codigo": proceso.returncode,
        "duracion_segundos": round(time.monotonic() - inicio, 2),
        "periodo": periodo,
        "fila_localizada": fila_localizada,
        "filas_localizadas": filas_localizadas,
        "periodos_por_fila": periodos_por_fila,
        "celdas_actualizadas": cambios,
        "hubo_cambios": bool(cambios),
        "salida_final": salida[-6000:],
        "error_final": error[-3000:],
    }


def main() -> int:
    for flujo in (sys.stdout, sys.stderr):
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass
    parser = argparse.ArgumentParser(description="Ejecutor de indicadores productivos")
    parser.add_argument("--fuentes", default="todas")
    parser.add_argument("--listar", action="store_true")
    parser.add_argument("--estricto", action="store_true")
    args = parser.parse_args()
    if args.listar:
        listado = {
            clave: {"nombre": fuente.nombre, "ruta": fuente.ruta_relativa, "existe": fuente.ruta().is_file()}
            for clave, fuente in FUENTES.items()
        }
        print(json.dumps(listado, ensure_ascii=False, indent=2))
        print("ESTADO=" + ("OK" if all(x["existe"] for x in listado.values()) else "ERROR"))
        return 0
    try:
        seleccion = seleccionar(args.fuentes)
    except ValueError as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 2
    inicio = datetime.now().astimezone()
    estado_antes = capturar_hoja()
    resultados = [ejecutar(clave, FUENTES[clave]) for clave in seleccion]
    errores = [x["fuente"] for x in resultados if x["estado"] == "ERROR"]
    total_cambios = sum(x.get("celdas_actualizadas") or 0 for x in resultados)
    estado_despues = capturar_hoja()
    actualizaciones_google = preparar_filas_actuales_google(resultados, estado_despues)
    resumen = {
        "inicio": inicio.isoformat(timespec="seconds"),
        "fin": datetime.now().astimezone().isoformat(timespec="seconds"),
        "estado": "ERROR" if errores else "OK",
        "errores": errores,
        "notificacion": {
            "hubo_cambios": total_cambios > 0,
            "total_celdas_actualizadas": total_cambios,
            "fuentes_modificadas": [x["fuente"] for x in resultados if x.get("hubo_cambios")],
        },
        "actualizaciones_google": actualizaciones_google,
        "resultados": resultados,
    }
    carpeta_n8n = RAIZ / "salidas_n8n"
    carpeta_n8n.mkdir(parents=True, exist_ok=True)
    destino_n8n = carpeta_n8n / "resumen_productivos.json"
    temporal_n8n = carpeta_n8n / "resumen_productivos.tmp"
    temporal_n8n.write_text(
        json.dumps(resumen, ensure_ascii=False, separators=(",", ":"), default=serializar_json),
        encoding="utf-8",
    )
    temporal_n8n.replace(destino_n8n)
    print("RESUMEN_PRODUCTIVOS_N8N=" + json.dumps(
        resumen, ensure_ascii=False, separators=(",", ":"), default=serializar_json
    ))
    print(f"ESTADO={resumen['estado']}")
    return 1 if errores else 0


if __name__ == "__main__":
    raise SystemExit(main())
