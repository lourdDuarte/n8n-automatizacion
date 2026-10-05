"""Actualiza la tabla COMERCIALIZACION de indicadores productivos.

Fuentes oficiales:
- Precios internos: Rosario ($/tn).
- Precios externos: FOB Puertos Argentinos (USD/tn).

La hoja, las columnas y los periodos se localizan por sus nombres. No utiliza
posiciones fijas de filas o columnas de la base.
"""

from __future__ import annotations

import argparse
import io
import json
import math
import re
import shutil
import sys
import time
import unicodedata
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


URL_INTERNOS = "https://www.magyp.gob.ar/sitio/areas/ss_mercados_agropecuarios/areas/granos/_archivos/000056_Precios%20Locales/000010_Precios%20Internos%20de%20los%20Principales%20Granos/000001_Informe%20Actual%20de%20la%20Evolucion%20Precios%20Internos%20de%20Granos.php"
URL_EXTERNOS = "https://www.magyp.gob.ar/sitio/areas/ss_mercados_agropecuarios/areas/granos/_archivos/000057_Precios%20Internacionales/000030_Evoluci%C3%B3n%20de%20los%20Precios%20Externos/000001_Granos%20%28Mensual%20en%20u%24s-tn%29/000002_Historicos%20de%20la%20Evoluci%C3%B3n%20de%20los%20Precios%20Externos%20de%20Granos.xlsx"
ARCHIVO_BASE = "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA_BASE = "BASE DE DATOS - Indicadores Pro"
PRODUCTOS = ("maiz", "sorgo", "soja", "trigo")
MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def raiz_proyecto() -> Path:
    return Path(__file__).resolve().parents[3]


def normalizar(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.casefold().replace("ñ", "n")
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return " ".join(texto.split())


def numero(valor) -> float:
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        resultado = float(valor)
    else:
        texto = str(valor or "").strip()
        if normalizar(texto) in {"", "s c", "sd", "s d", "nan", "none", "-"}:
            raise ValueError(f"Valor sin cotización: {valor!r}")
        texto = texto.replace("$", "").replace("USD", "").replace("US$", "")
        texto = texto.replace("\u00a0", "").replace(" ", "")
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        resultado = float(texto)
    if not math.isfinite(resultado):
        raise ValueError(f"Valor numérico inválido: {valor!r}")
    return resultado


def descargar_bytes(url: str, intentos: int = 3, timeout: int = 60) -> tuple[bytes, str]:
    ultimo_error = None
    for intento in range(1, intentos + 1):
        solicitud = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml",
                "Accept-Language": "es-AR,es;q=0.9",
            },
        )
        try:
            with urllib.request.urlopen(solicitud, timeout=timeout) as respuesta:
                datos = respuesta.read()
                codificacion = respuesta.headers.get_content_charset() or "utf-8"
            return datos, codificacion
        except (urllib.error.URLError, TimeoutError, OSError) as error:
            ultimo_error = error
            if intento < intentos:
                time.sleep(intento * 2)
    raise ConnectionError(f"No se pudo descargar {url}: {ultimo_error}")


def descargar_texto(url: str) -> str:
    datos, codificacion = descargar_bytes(url)
    return datos.decode(codificacion, errors="replace")


def columnas_planas(tabla: pd.DataFrame) -> list[str]:
    resultado = []
    for columna in tabla.columns:
        partes = columna if isinstance(columna, tuple) else (columna,)
        limpias = []
        for parte in partes:
            texto = normalizar(parte)
            if texto and not texto.startswith("unnamed") and texto not in limpias:
                limpias.append(texto)
        resultado.append(" ".join(limpias))
    return resultado


def mes_de(valor) -> int | None:
    texto = normalizar(valor)
    for nombre, mes in MESES.items():
        if texto == nombre or texto.startswith(nombre + " "):
            return mes
    return None


def anio_del_html(html: str) -> int:
    actual = date.today().year
    candidatos = [int(x) for x in re.findall(r"(?<!\d)(20\d{2})(?!\d)", html)]
    candidatos = [x for x in candidatos if 2000 <= x <= actual + 1]
    if not candidatos:
        raise ValueError("No se pudo identificar el año de la publicación")
    return max(candidatos)


def buscar_columna(encabezados: list[str], producto: str, tipo: str) -> int:
    candidatos = []
    for indice, titulo in enumerate(encabezados):
        if producto not in titulo:
            continue
        if tipo == "interno" and "rosario" in titulo:
            candidatos.append(indice)
        if tipo == "externo" and (
            "ptos arg" in titulo
            or "pto arg" in titulo
            or "puertos arg" in titulo
            or "puerto arg" in titulo
        ):
            candidatos.append(indice)
    if len(candidatos) != 1:
        raise ValueError(
            f"Se esperada una columna {producto}/{tipo} y se encontraron {len(candidatos)}. "
            f"Encabezados: {encabezados}"
        )
    return candidatos[0]


class LectorTablasHTML(HTMLParser):
    """Lector pequeño de tablas HTML, sin dependencias adicionales."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tablas: list[list[list[dict]]] = []
        self.profundidad_tabla = 0
        self.tabla_actual = None
        self.fila_actual = None
        self.celda_actual = None

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.casefold()
        if tag == "table":
            self.profundidad_tabla += 1
            if self.profundidad_tabla == 1:
                self.tabla_actual = []
        elif self.profundidad_tabla == 1 and tag == "tr":
            self.fila_actual = []
        elif self.profundidad_tabla == 1 and tag in {"th", "td"}:
            self.celda_actual = {"attrs": dict(attrs), "texto": []}
        elif self.celda_actual is not None and tag == "br":
            self.celda_actual["texto"].append(" ")

    def handle_data(self, data: str) -> None:
        if self.celda_actual is not None:
            self.celda_actual["texto"].append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.casefold()
        if self.profundidad_tabla == 1 and tag in {"th", "td"}:
            if self.fila_actual is not None and self.celda_actual is not None:
                self.celda_actual["texto"] = " ".join(
                    "".join(self.celda_actual["texto"]).split()
                )
                self.fila_actual.append(self.celda_actual)
            self.celda_actual = None
        elif self.profundidad_tabla == 1 and tag == "tr":
            if self.tabla_actual is not None and self.fila_actual:
                self.tabla_actual.append(self.fila_actual)
            self.fila_actual = None
        elif tag == "table" and self.profundidad_tabla:
            if self.profundidad_tabla == 1 and self.tabla_actual:
                self.tablas.append(self.tabla_actual)
                self.tabla_actual = None
            self.profundidad_tabla -= 1


def expandir_tabla_html(tabla_html) -> list[list[str]]:
    """Expande rowspan/colspan para conservar la relación producto-localidad."""
    filas: list[list[str]] = []
    pendientes: dict[int, tuple[int, str]] = {}
    for tr in tabla_html:
        fila: list[str] = []
        columna = 0

        def incorporar_pendientes() -> None:
            nonlocal columna
            while columna in pendientes:
                restantes, texto = pendientes[columna]
                fila.append(texto)
                if restantes <= 1:
                    del pendientes[columna]
                else:
                    pendientes[columna] = (restantes - 1, texto)
                columna += 1

        incorporar_pendientes()
        for celda in tr:
            incorporar_pendientes()
            texto = celda["texto"]
            try:
                colspan = max(1, int(celda["attrs"].get("colspan", 1)))
            except (TypeError, ValueError):
                colspan = 1
            try:
                rowspan = max(1, int(celda["attrs"].get("rowspan", 1)))
            except (TypeError, ValueError):
                rowspan = 1
            for _ in range(colspan):
                fila.append(texto)
                if rowspan > 1:
                    pendientes[columna] = (rowspan - 1, texto)
                columna += 1
        incorporar_pendientes()
        if any(normalizar(valor) for valor in fila):
            filas.append(fila)
    ancho = max((len(fila) for fila in filas), default=0)
    return [fila + [""] * (ancho - len(fila)) for fila in filas]


def encabezados_desde_matriz(matriz: list[list[str]]) -> tuple[list[str], int]:
    primera_fila_mes = next(
        (indice for indice, fila in enumerate(matriz) if any(mes_de(v) for v in fila)),
        None,
    )
    if primera_fila_mes is None:
        return [], 0
    ancho = max(len(fila) for fila in matriz)
    encabezados = []
    for columna in range(ancho):
        partes = []
        for fila in matriz[:primera_fila_mes]:
            texto = normalizar(fila[columna] if columna < len(fila) else "")
            if not texto or texto in partes or texto in {"ano mes", "mes", "ano"}:
                continue
            if re.fullmatch(r"20\d{2}", texto):
                continue
            partes.append(texto)
        encabezados.append(" ".join(partes))
    return encabezados, primera_fila_mes


def separar_bloques_productos(matriz: list[list[str]]) -> list[list[list[str]]]:
    """Separa bloques repetidos dentro de una misma tabla HTML oficial."""
    inicios = []
    for indice, fila in enumerate(matriz):
        textos = [normalizar(valor) for valor in fila]
        cantidad_productos = sum(
            any(producto == texto or producto in texto.split() for texto in textos)
            for producto in PRODUCTOS
        )
        if cantidad_productos and not any(mes_de(valor) for valor in fila):
            inicios.append(indice)
    if not inicios:
        return [matriz]
    bloques = []
    for posicion, inicio in enumerate(inicios):
        fin = inicios[posicion + 1] if posicion + 1 < len(inicios) else len(matriz)
        bloques.append(matriz[inicio:fin])
    return bloques


def extraer_tabla(html: str, tipo: str) -> dict:
    lector = LectorTablasHTML()
    lector.feed(html)
    tablas = lector.tablas
    if not tablas:
        raise ValueError(f"La página de precios {tipo}s no contiene tablas legibles")
    anio = anio_del_html(html)
    # La fuente oficial puede repartir los productos entre varias tablas
    # (por ejemplo: Trigo/Maíz/Sorgo en una y Soja en otra). Se reúnen por
    # período y recién se acepta un mes cuando están los cuatro productos.
    por_periodo: dict[date, dict[str, float]] = {}
    diagnostico_encabezados = []
    for tabla in tablas:
        matriz_completa = expandir_tabla_html(tabla)
        for matriz in separar_bloques_productos(matriz_completa):
            encabezados, primera_fila_mes = encabezados_desde_matriz(matriz)
            if encabezados:
                diagnostico_encabezados.append(encabezados)
            indices = {}
            for producto in PRODUCTOS:
                try:
                    indices[producto] = buscar_columna(encabezados, producto, tipo)
                except ValueError:
                    continue
            if not indices:
                continue

            for fila in matriz[primera_fila_mes:]:
                meses = [mes_de(valor) for valor in fila]
                meses_validos = [m for m in meses if m]
                if not meses_validos:
                    continue
                mes = meses_validos[0]
                periodo = date(anio, mes, 1)
                valores_periodo = por_periodo.setdefault(periodo, {})
                for producto, indice in indices.items():
                    try:
                        valores_periodo[producto] = numero(fila[indice])
                    except (ValueError, TypeError, IndexError):
                        continue

    completos = [
        {"periodo": periodo, "valores": valores}
        for periodo, valores in por_periodo.items()
        if all(producto in valores for producto in PRODUCTOS)
    ]
    if not completos:
        encontrados = {
            periodo.isoformat(): sorted(valores)
            for periodo, valores in por_periodo.items()
        }
        raise ValueError(
            f"No se reunieron Maíz, Sorgo, Soja y Trigo para precios {tipo}s. "
            f"Productos encontrados por período: {encontrados}. "
            f"Encabezados detectados: {diagnostico_encabezados}"
        )
    return max(completos, key=lambda item: item["periodo"])


def extraer_externos_excel(contenido: bytes) -> dict:
    libro = load_workbook(io.BytesIO(contenido), data_only=True, read_only=True)
    if "PREFOB" not in libro.sheetnames:
        raise ValueError("El Excel de precios externos no contiene la hoja PREFOB")
    ws = libro["PREFOB"]

    fila_productos = None
    for fila in range(1, min(ws.max_row, 20) + 1):
        encontrados = {
            normalizar(ws.cell(fila, columna).value)
            for columna in range(1, ws.max_column + 1)
        }
        if all(producto in encontrados for producto in PRODUCTOS):
            fila_productos = fila
            break
    if fila_productos is None:
        raise ValueError("No se encontró el encabezado de productos en PREFOB")

    columnas = {}
    producto_actual = None
    for columna in range(1, ws.max_column + 1):
        posible = normalizar(ws.cell(fila_productos, columna).value)
        if posible in PRODUCTOS:
            producto_actual = posible
        ubicacion = " ".join(
            normalizar(ws.cell(fila, columna).value)
            for fila in range(fila_productos + 1, min(fila_productos + 4, ws.max_row) + 1)
        )
        if producto_actual and (
            "ptos arg" in ubicacion or "puertos arg" in ubicacion or "pto arg" in ubicacion
        ):
            columnas[producto_actual] = columna
    if set(columnas) != set(PRODUCTOS):
        raise ValueError(f"No se identificaron las cuatro columnas FOB PTOS.ARG.: {columnas}")

    por_periodo = {}
    anio_actual = None
    for fila in range(fila_productos + 1, ws.max_row + 1):
        primero = ws.cell(fila, 1).value
        if isinstance(primero, (int, float)) and int(primero) == primero and 1990 <= int(primero) <= 2100:
            anio_actual = int(primero)
            continue
        mes = mes_de(primero)
        if not mes or anio_actual is None:
            continue
        try:
            valores = {
                producto: numero(ws.cell(fila, columna).value)
                for producto, columna in columnas.items()
            }
        except ValueError:
            continue
        por_periodo[date(anio_actual, mes, 1)] = valores
    if not por_periodo:
        raise ValueError("El Excel PREFOB no contiene períodos completos")
    periodo = max(por_periodo)
    return {"periodo": periodo, "valores": por_periodo[periodo]}


def extraer(html_internos: str, excel_externos: bytes) -> dict:
    internos = extraer_tabla(html_internos, "interno")
    externos = extraer_externos_excel(excel_externos)
    if internos["periodo"] != externos["periodo"]:
        raise ValueError(
            "Las fuentes no tienen el mismo último período: "
            f"internos={internos['periodo']:%Y-%m}, externos={externos['periodo']:%Y-%m}"
        )
    valores = {}
    for producto in PRODUCTOS:
        valores[f"nacional_{producto}"] = internos["valores"][producto]
        valores[f"internacional_{producto}"] = externos["valores"][producto]
    return {
        "fuente": "Agricultura - Comercialización de granos",
        "periodo": internos["periodo"].isoformat(),
        "anio": internos["periodo"].year,
        "mes_numero": internos["periodo"].month,
        "consultado_en": datetime.now().astimezone().isoformat(timespec="seconds"),
        "url_precios_internos": URL_INTERNOS,
        "url_precios_externos": URL_EXTERNOS,
        "valores": valores,
        "validaciones": {
            "precios_internos_rosario": True,
            "precios_externos_fob_puertos_argentinos": True,
            "fuentes_mismo_periodo": True,
            "cantidad_valores_originales": len(valores),
        },
    }


@dataclass(frozen=True)
class ColumnaIndicador:
    producto: str
    clase: str
    datos: int
    variacion: int
    titulo_datos: str
    titulo_variacion: str


def contiene_tokens(texto: str, *tokens: str) -> bool:
    palabras = set(normalizar(texto).split())
    return all(token in palabras for token in tokens)


def buscar_unica(encabezados: dict[int, str], descripcion: str, predicado) -> tuple[int, str]:
    encontrados = [(columna, titulo) for columna, titulo in encabezados.items() if predicado(titulo)]
    if len(encontrados) != 1:
        detalle = [titulo for _, titulo in encontrados]
        raise ValueError(f"Se esperaba una columna para {descripcion}; encontradas={detalle}")
    return encontrados[0]


def localizar_columnas(ws) -> dict[str, ColumnaIndicador]:
    mejor_fila, encabezados = None, None
    for fila in range(1, min(ws.max_row, 12) + 1):
        actuales = {
            columna: str(ws.cell(fila, columna).value or "")
            for columna in range(1, ws.max_column + 1)
        }
        cantidad = sum("precio" in normalizar(v) for v in actuales.values())
        if encabezados is None or cantidad > sum("precio" in normalizar(v) for v in encabezados.values()):
            mejor_fila, encabezados = fila, actuales
    if encabezados is None or mejor_fila is None:
        raise ValueError("No se encontró la fila de títulos")

    resultado = {}
    for producto in PRODUCTOS:
        for clase in ("nacional", "internacional"):
            datos_col, datos_titulo = buscar_unica(
                encabezados,
                f"precio {clase} de {producto}",
                lambda t, p=producto, c=clase: contiene_tokens(t, "precio", c, p)
                and "variacion" not in normalizar(t),
            )
            variacion_col, variacion_titulo = buscar_unica(
                encabezados,
                f"variación del precio {clase} de {producto}",
                lambda t, p=producto, c=clase: contiene_tokens(t, "variacion", p)
                and ((c == "internacional" and "fob" in normalizar(t))
                     or (c == "nacional" and "fob" not in normalizar(t))),
            )
            clave = f"{clase}_{producto}"
            resultado[clave] = ColumnaIndicador(
                producto, clase, datos_col, variacion_col, datos_titulo, variacion_titulo
            )
    return resultado


def periodos_de_hoja(ws) -> dict[tuple[int, int], int]:
    periodos = {}
    anio_actual = None
    for fila in range(1, ws.max_row + 1):
        celdas = [ws.cell(fila, 1).value, ws.cell(fila, 2).value]
        for valor in celdas:
            if isinstance(valor, (int, float)) and int(valor) == valor and 2000 <= int(valor) <= 2100:
                anio_actual = int(valor)
                break
            texto = normalizar(valor)
            if re.fullmatch(r"20\d{2}", texto):
                anio_actual = int(texto)
                break
        if anio_actual is None:
            continue
        mes = next((mes_de(v) for v in celdas if mes_de(v)), None)
        if mes:
            clave = (anio_actual, mes)
            if clave in periodos:
                raise ValueError(f"El período {anio_actual}-{mes:02d} aparece más de una vez")
            periodos[clave] = fila
    return periodos


def iguales(anterior, nuevo) -> bool:
    if isinstance(anterior, (int, float)) and isinstance(nuevo, (int, float)):
        return math.isclose(float(anterior), float(nuevo), rel_tol=0, abs_tol=1e-6)
    return anterior == nuevo


def actualizar_base(resultado: dict, ruta_base: Path) -> dict:
    if not ruta_base.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta_base}")
    libro = load_workbook(ruta_base, data_only=False)
    nombre_hoja = next(
        (n for n in libro.sheetnames if normalizar(n) == normalizar(HOJA_BASE)), None
    )
    if nombre_hoja is None:
        raise ValueError(f"No existe la hoja {HOJA_BASE!r}")
    ws = libro[nombre_hoja]
    columnas = localizar_columnas(ws)
    periodos = periodos_de_hoja(ws)
    periodo = (resultado["anio"], resultado["mes_numero"])
    if periodo not in periodos:
        raise ValueError(f"No existe el período {periodo[0]}-{periodo[1]:02d} en la hoja")
    fila = periodos[periodo]

    mes_anterior_fecha = date(periodo[0], periodo[1], 1)
    if mes_anterior_fecha.month == 1:
        anterior = (mes_anterior_fecha.year - 1, 12)
    else:
        anterior = (mes_anterior_fecha.year, mes_anterior_fecha.month - 1)
    if anterior not in periodos:
        raise ValueError(f"Falta el mes anterior {anterior[0]}-{anterior[1]:02d} para calcular variaciones")
    fila_anterior = periodos[anterior]

    cambios = []
    coincidencias = []
    formulas_conservadas = []
    for clave, nuevo in resultado["valores"].items():
        indicador = columnas[clave]
        celda_dato = ws.cell(fila, indicador.datos)
        antes = celda_dato.value
        if iguales(antes, nuevo):
            coincidencias.append(celda_dato.coordinate)
        else:
            celda_dato.value = nuevo
            cambios.append({
                "celda": celda_dato.coordinate,
                "titulo": indicador.titulo_datos,
                "antes": antes,
                "despues": nuevo,
            })

        celda_variacion = ws.cell(fila, indicador.variacion)
        col_dato = get_column_letter(indicador.datos)
        formula = f"=(({col_dato}{fila}-{col_dato}{fila_anterior})/{col_dato}{fila_anterior})*100"
        antes_formula = celda_variacion.value
        if antes_formula in (None, ""):
            celda_variacion.value = formula
            cambios.append({
                "celda": celda_variacion.coordinate,
                "titulo": indicador.titulo_variacion,
                "antes": antes_formula,
                "despues": formula,
            })
        elif isinstance(antes_formula, str) and antes_formula.startswith("="):
            formulas_conservadas.append(celda_variacion.coordinate)
        else:
            formulas_conservadas.append(celda_variacion.coordinate)

    respaldo = None
    if cambios:
        carpeta = ruta_base.parent / "respaldos"
        carpeta.mkdir(parents=True, exist_ok=True)
        respaldo = carpeta / f"antes_comercializacion_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(ruta_base, respaldo)
        libro.calculation.fullCalcOnLoad = True
        libro.calculation.forceFullCalc = True
        libro.calculation.calcMode = "auto"
        libro.save(ruta_base)

    return {
        "archivo_excel": str(ruta_base.resolve()),
        "hoja": nombre_hoja,
        "periodo": f"{periodo[0]}-{periodo[1]:02d}",
        "fila_localizada": fila,
        "celdas_actualizadas": len(cambios),
        "hubo_cambios": bool(cambios),
        "cambios": cambios,
        "coincidencias": coincidencias,
        "formulas_conservadas": formulas_conservadas,
        "respaldo": str(respaldo.resolve()) if respaldo else None,
    }


def guardar_reporte(resultado: dict, raiz: Path) -> Path:
    carpeta = raiz / "scripts" / "reportes" / "productivos" / "comercializacion"
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / f"comercializacion_{resultado['anio']}_{resultado['mes_numero']:02d}.json"
    ruta.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    return ruta


def main() -> int:
    raiz = raiz_proyecto()
    parser = argparse.ArgumentParser(description="Actualiza Comercialización de granos")
    parser.add_argument("--base", type=Path, default=raiz / "base" / ARCHIVO_BASE)
    parser.add_argument("--solo-extraer", action="store_true")
    parser.add_argument("--html-internos", type=Path)
    parser.add_argument("--xlsx-externos", type=Path)
    args = parser.parse_args()
    try:
        html_internos = (
            args.html_internos.read_text(encoding="utf-8", errors="replace")
            if args.html_internos else descargar_texto(URL_INTERNOS)
        )
        excel_externos = (
            args.xlsx_externos.read_bytes()
            if args.xlsx_externos else descargar_bytes(URL_EXTERNOS)[0]
        )
        resultado = extraer(html_internos, excel_externos)
        if not args.solo_extraer:
            resultado["actualizacion"] = actualizar_base(resultado, args.base)
        reporte = guardar_reporte(resultado, raiz)
        resultado["reporte"] = str(reporte.resolve())
        print("RESULTADO_COMERCIALIZACION=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
        print("ESTADO=OK")
        return 0
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
