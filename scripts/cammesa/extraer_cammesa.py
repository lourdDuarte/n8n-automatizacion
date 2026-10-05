"""Extrae la demanda mensual de Formosa desde la base oficial de CAMMESA."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
import json
import re
import sys
import zipfile

from openpyxl import load_workbook


RAIZ = Path(__file__).resolve().parent.parent
FUENTES = RAIZ / "fuentes" / "cammesa"
REPORTES = RAIZ / "reportes" / "cammesa"

CATEGORIAS = {
    "Comercial": "demanda_comercial",
    "Industrial/Comercial Grande": "demanda_industrial_comercial_grande",
    "Residencial": "demanda_residencial",
}


def configurar_salida() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def periodo_del_zip(ruta: Path) -> str:
    coincidencia = re.search(r"BASE_INFORME_MENSUAL_(20\d{2})-(0[1-9]|1[0-2])", ruta.name, re.I)
    if not coincidencia:
        raise ValueError(f"No se pudo obtener el período desde {ruta.name!r}.")
    return f"{coincidencia.group(1)}-{coincidencia.group(2)}"


def buscar_ultimo_zip() -> Path:
    archivos = [p for p in FUENTES.glob("BASE_INFORME_MENSUAL_????-??.zip") if p.is_file()]
    if not archivos:
        raise FileNotFoundError(f"No hay bases mensuales de CAMMESA en {FUENTES}.")
    return max(archivos, key=periodo_del_zip)


def extraer_demanda_mensual(ruta_zip: Path) -> Path:
    periodo = periodo_del_zip(ruta_zip)
    destino = FUENTES / f"Demanda_Mensual_{periodo}.xlsx"
    with zipfile.ZipFile(ruta_zip) as archivo:
        candidatos = [
            nombre for nombre in archivo.namelist()
            if nombre.replace("\\", "/").endswith(
                "Bases_Demanda_INFORME_MENSUAL/Demanda Mensual.xlsx"
            )
        ]
        if len(candidatos) != 1:
            raise RuntimeError(
                "No se encontró una única planilla "
                "Bases_Demanda_INFORME_MENSUAL/Demanda Mensual.xlsx."
            )
        contenido = archivo.read(candidatos[0])
    if not contenido.startswith(b"PK"):
        raise RuntimeError("La planilla Demanda Mensual.xlsx no es válida.")
    destino.write_bytes(contenido)
    return destino


def encabezados(hoja) -> dict[str, int]:
    resultado = {
        str(celda.value).strip().upper(): celda.column
        for celda in hoja[22]
        if celda.value not in (None, "")
    }
    requeridos = {
        "MES", "PROVINCIA", "CATEGORIA DEMANDA", "CATEGORIA TARIFA",
        "TARIFA", "TIPO DEMANDA", "DEMANDA [MWH]",
    }
    faltantes = sorted(requeridos - set(resultado))
    if faltantes:
        raise RuntimeError("Faltan columnas en Demanda Mensual: " + ", ".join(faltantes))
    return resultado


def sumar_periodos(ruta_excel: Path, fechas: set[tuple[int, int]]) -> tuple[dict, dict]:
    libro = load_workbook(ruta_excel, read_only=True, data_only=True)
    try:
        if "DEMANDA" not in libro.sheetnames:
            raise RuntimeError("La planilla no contiene la hoja DEMANDA.")
        hoja = libro["DEMANDA"]
        columnas = encabezados(hoja)
        totales = defaultdict(float)
        cantidades = defaultdict(int)

        for fila in hoja.iter_rows(min_row=23, values_only=True):
            mes = fila[columnas["MES"] - 1]
            provincia = str(fila[columnas["PROVINCIA"] - 1] or "").strip().upper()
            categoria = str(fila[columnas["CATEGORIA TARIFA"] - 1] or "").strip()
            valor = fila[columnas["DEMANDA [MWH]"] - 1]
            if not isinstance(mes, datetime) or (mes.year, mes.month) not in fechas:
                continue
            if provincia != "FORMOSA" or categoria not in CATEGORIAS:
                continue
            if isinstance(valor, bool) or not isinstance(valor, (int, float)):
                raise ValueError(f"Demanda no numérica para Formosa {mes:%Y-%m} / {categoria}: {valor!r}")
            clave = ((mes.year, mes.month), categoria)
            totales[clave] += float(valor)
            cantidades[clave] += 1

        for fecha in fechas:
            for categoria in CATEGORIAS:
                if cantidades[(fecha, categoria)] == 0:
                    raise RuntimeError(
                        f"No se encontraron registros de Formosa para {fecha[0]:04d}-{fecha[1]:02d} / {categoria}."
                    )
        return dict(totales), dict(cantidades)
    finally:
        libro.close()


def mes_anterior(anio: int, mes: int) -> tuple[int, int]:
    return (anio - 1, 12) if mes == 1 else (anio, mes - 1)


def variacion(actual: float, base: float, descripcion: str) -> float:
    if base == 0:
        raise ZeroDivisionError(f"No se puede calcular {descripcion}: el valor base es cero.")
    return round((actual / base - 1) * 100, 2)


def extraer(ruta_zip: Path) -> dict:
    periodo = periodo_del_zip(ruta_zip)
    anio, mes = map(int, periodo.split("-"))
    actual = (anio, mes)
    anterior = mes_anterior(anio, mes)
    interanual = (anio - 1, mes)
    ruta_excel = extraer_demanda_mensual(ruta_zip)
    totales, cantidades = sumar_periodos(ruta_excel, {actual, anterior, interanual})

    datos = {}
    controles = {}
    for categoria, clave in CATEGORIAS.items():
        valor_actual = round(totales[(actual, categoria)], 2)
        valor_anterior = totales[(anterior, categoria)]
        valor_interanual = totales[(interanual, categoria)]
        datos[clave] = valor_actual
        datos[f"variacion_mensual_{clave.removeprefix('demanda_')}"] = variacion(
            valor_actual, valor_anterior, f"variación mensual de {categoria}"
        )
        datos[f"variacion_anual_{clave.removeprefix('demanda_')}"] = variacion(
            valor_actual, valor_interanual, f"variación interanual de {categoria}"
        )
        controles[categoria] = {
            "registros_actuales": cantidades[(actual, categoria)],
            "actual": valor_actual,
            "mes_anterior": round(valor_anterior, 2),
            "mismo_mes_anio_anterior": round(valor_interanual, 2),
        }

    return {
        "fuente": "CAMMESA - Demanda Mensual",
        "periodo": periodo,
        "unidad": "MWh",
        "archivo_zip": ruta_zip.name,
        "archivo_excel": ruta_excel.name,
        "datos": datos,
        "validaciones": controles,
    }


def main() -> int:
    configurar_salida()
    resultado = extraer(buscar_ultimo_zip())
    REPORTES.mkdir(parents=True, exist_ok=True)
    salida = REPORTES / f"cammesa_{resultado['periodo'].replace('-', '_')}.json"
    salida.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    print("=" * 72)
    print("EXTRACCIÓN — CAMMESA DEMANDA FORMOSA")
    print("=" * 72)
    print(f"Período: {resultado['periodo']}")
    print("Valores validados: 3 demandas + 6 variaciones")
    print(f"CAMMESA_JSON={salida.resolve()}")
    print("RESULTADO: OK")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        configurar_salida()
        print(f"ERROR: {type(error).__name__}: {error}", file=sys.stderr)
        raise SystemExit(1)
