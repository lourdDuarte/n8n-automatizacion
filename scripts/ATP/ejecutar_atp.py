"""Consulta ATP Formosa y actualiza la recaudación tributaria mensual.

El destino se resuelve por el encabezado PERIODO y por el título del
indicador. No depende de letras de columnas ni de números de filas.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import urllib.request
from copy import copy
from datetime import date, datetime
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator


RAIZ = Path(__file__).resolve().parents[2]
SCRIPTS = RAIZ / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from comun.columnas_por_titulo import resolver_columnas


URL_FUENTE = "https://www.formosa.gob.ar/atp/recaudaciones"
BASE_PREDETERMINADA = RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA_PREDETERMINADA = "BASE DE DATOS"
TITULOS_ATP = {
    "totales": "Recaudación tributaria F",
    "imp sobre ingresos brutos": "Recaudación de IIBB F",
    "imp de sellos": "Recaudación Imp.\nde Sellos F",
    "tasas retrib de servicios": "Recaudacion tasas Retrib.\nde Servicios F",
    "reg de facil de pago": "Recaudacion Reg. de Facil.\nde Pago\t",
    "inmob rural": "Recaudacion Inmob. Rural",
    "imp a los juegos de azar": "Recaudacion Imp. a los\nJuegos de Azar",
    "transf mej e tier fisc": "Recaudacion Transf. Mej.\ne/Tier.\nFisc.",
}

# El informe de transferencias que ya descarga el módulo IARAF incluye
# explícitamente la inflación supuesta antes de la publicación de INDEC.
CARPETA_IARAF = RAIZ / "fuentes" / "iaraf"
TITULOS_REALES = (
    "Coeficiente ajuste Mensual IPC NEA/IERAF",
    "Coeficiente ajuste Anual IPC NEA",
    "Recaudación tributaria F Real IM",
    "Recaudación tributaria F Real IA",
    "Variación intermensual recaudación tributaria F Real",
    "Variación interanual recaudación tributaria F Real",
    "Recaudación tributaria acumulada F Real",
    "Recaudación de IIBB F Real",
    "Variación intermensual recaudación de IIBB F Real",
    "Recaudación IIBB Real IA",
    "Variación interanual recaudación de IIBB F Real",
)

MESES = {
    "ene": 1, "enero": 1, "feb": 2, "febrero": 2,
    "mar": 3, "marzo": 3, "abr": 4, "abril": 4,
    "may": 5, "mayo": 5, "jun": 6, "junio": 6,
    "jul": 7, "julio": 7, "ago": 8, "agosto": 8,
    "sep": 9, "sept": 9, "septiembre": 9, "setiembre": 9,
    "oct": 10, "octubre": 10, "nov": 11, "noviembre": 11,
    "dic": 12, "diciembre": 12,
}


def configurar_salida() -> None:
    for flujo in (sys.stdout, sys.stderr):
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def normalizar(valor: object) -> str:
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", texto.casefold()).strip()


def numero_argentino(texto: str) -> float:
    limpio = re.sub(r"[^0-9,.-]", "", texto).replace(".", "").replace(",", ".")
    if not limpio or limpio in {"-", ".", "-."}:
        raise ValueError(f"Importe ATP inválido: {texto!r}")
    return float(limpio)


def descargar_html(timeout: int = 60) -> str:
    solicitud = urllib.request.Request(
        URL_FUENTE,
        headers={"User-Agent": "Mozilla/5.0 (automatizacion-indicadores/1.0)"},
    )
    with urllib.request.urlopen(solicitud, timeout=timeout) as respuesta:
        contenido = respuesta.read()
        codificacion = respuesta.headers.get_content_charset() or "utf-8"
    return contenido.decode(codificacion, errors="replace")


class LectorTablas(HTMLParser):
    """Convierte las tablas HTML en matrices de texto sin dependencias externas."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tablas: list[list[list[str]]] = []
        self._tabla: list[list[str]] | None = None
        self._fila: list[str] | None = None
        self._celda: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.casefold()
        if tag == "table":
            self._tabla = []
        elif tag == "tr" and self._tabla is not None:
            self._fila = []
        elif tag in {"td", "th"} and self._fila is not None:
            self._celda = []

    def handle_data(self, data: str) -> None:
        if self._celda is not None:
            self._celda.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.casefold()
        if tag in {"td", "th"} and self._celda is not None and self._fila is not None:
            self._fila.append(" ".join("".join(self._celda).split()))
            self._celda = None
        elif tag == "tr" and self._fila is not None and self._tabla is not None:
            if self._fila:
                self._tabla.append(self._fila)
            self._fila = None
        elif tag == "table" and self._tabla is not None:
            if self._tabla:
                self.tablas.append(self._tabla)
            self._tabla = None


def extraer_ultimo_periodo(html: str, anio: int) -> dict:
    lector = LectorTablas()
    lector.feed(html)
    candidatos: list[dict] = []

    for filas in lector.tablas:
        if not filas:
            continue
        encabezados = [normalizar(celda) for celda in filas[0]]
        if not encabezados or "totales" not in encabezados:
            continue
        columnas = {
            clave: encabezados.index(clave)
            for clave in TITULOS_ATP
            if clave in encabezados
        }
        faltantes = [clave for clave in TITULOS_ATP if clave not in columnas]
        if faltantes:
            continue

        for fila in filas[1:]:
            celdas = fila
            if len(celdas) <= max(columnas.values()):
                continue
            mes_texto = normalizar(celdas[0])
            mes = MESES.get(mes_texto)
            if mes is None:
                continue
            valores = {
                TITULOS_ATP[clave]: numero_argentino(celdas[columna])
                for clave, columna in columnas.items()
            }
            candidatos.append({
                "periodo": f"{anio:04d}-{mes:02d}",
                "anio": anio,
                "mes": mes,
                "mes_fuente": celdas[0],
                "recaudacion": valores[TITULOS_ATP["totales"]],
                "valores": valores,
            })

    if not candidatos:
        raise ValueError("No se encontró la tabla mensual de recaudación de ATP.")
    return max(candidatos, key=lambda item: (item["anio"], item["mes"]))


def periodo_celda(valor: object) -> tuple[int, int] | None:
    if isinstance(valor, (date, datetime)):
        return valor.year, valor.month
    texto = normalizar(valor)
    if not texto:
        return None
    coincidencia = re.fullmatch(r"(20\d{2}) (0?[1-9]|1[0-2])", texto)
    if coincidencia:
        return int(coincidencia.group(1)), int(coincidencia.group(2))
    coincidencia = re.fullmatch(r"([a-z]+) (\d{2}|20\d{2})", texto)
    if coincidencia and coincidencia.group(1) in MESES:
        anio = int(coincidencia.group(2))
        return (anio + 2000 if anio < 100 else anio), MESES[coincidencia.group(1)]
    return None


def buscar_fila(hoja, columna_periodo: int, periodo: str) -> int:
    objetivo = tuple(map(int, periodo.split("-")))
    coincidencias = [
        fila
        for fila in range(2, hoja.max_row + 1)
        if periodo_celda(hoja.cell(fila, columna_periodo).value) == objetivo
    ]
    if not coincidencias:
        raise ValueError(f"No se encontró el período {periodo} en PERIODO.")
    if len(coincidencias) != 1:
        raise ValueError(f"El período {periodo} está repetido en las filas {coincidencias}.")
    return coincidencias[0]


def iguales(actual: object, nuevo: float, tolerancia: float = 0.005) -> bool:
    return (
        isinstance(actual, (int, float))
        and not isinstance(actual, bool)
        and abs(float(actual) - nuevo) <= tolerancia
    )


def guardar_atomico(libro, ruta: Path) -> None:
    descriptor, nombre = tempfile.mkstemp(prefix=f"{ruta.stem}_", suffix=ruta.suffix, dir=ruta.parent)
    os.close(descriptor)
    temporal = Path(nombre)
    try:
        libro.save(temporal)
        os.replace(temporal, ruta)
    finally:
        if temporal.exists():
            temporal.unlink()


def estimacion_iaraf(anio: int, mes: int, informe: Path | None = None) -> tuple[float, str]:
    """Lee el supuesto del PDF local ya descargado por el módulo IARAF."""
    ruta = informe or CARPETA_IARAF / f"iaraf_transferencias_{anio:04d}_{mes:02d}.pdf"
    if not ruta.is_file():
        raise FileNotFoundError(
            f"Falta el informe IARAF de {anio}-{mes:02d}: {ruta}. "
            "Ejecutá primero el módulo IARAF o pasá --informe-iaraf."
        )
    import fitz
    with fitz.open(ruta) as pdf:
        texto = " ".join(pagina.get_text() for pagina in pdf)
    texto = unicodedata.normalize("NFKD", texto.casefold())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    nombre_mes = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
                  "septiembre", "octubre", "noviembre", "diciembre")[mes - 1]
    if not re.search(rf"\b{nombre_mes}\s+(?:de\s+)?{anio}\b", texto):
        raise ValueError(f"El informe {ruta.name} no corresponde a {nombre_mes} de {anio}.")
    patrones = (
        rf"(?:para|en) el mes de {nombre_mes} (?:se considera|se supuso) "
        rf"una inflacion del?\s*([0-9]+[,.][0-9]+)\s*%",
        rf"(?:se considera|se supuso) una inflacion del?\s*"
        rf"([0-9]+[,.][0-9]+)\s*%\s*(?:en|para)\s*"
        rf"(?:el mes de )?{nombre_mes}",
        rf"para la elaboracion del informe se supuso una inflacion del?\s*"
        rf"([0-9]+[,.][0-9]+)\s*%\s*en\s*{nombre_mes}",
    )
    verbo = r"(?:se considera|se considero|se supuso|se supone|se asume|se asumio|se estima|se estimo)"
    patrones = patrones + (
        rf"{verbo} una inflacion del?\s*([0-9]+[,.][0-9]+)\s*%\s*(?:en|para)\s*(?:el mes de\s*)?{nombre_mes}",
        rf"(?:para|en) el mes de {nombre_mes} {verbo} una inflacion del?\s*([0-9]+[,.][0-9]+)\s*%",
    )
    tasas = {float(valor.replace(",", ".")) for patron in patrones
             for valor in re.findall(patron, texto)}
    if len(tasas) != 1 or not 0 <= next(iter(tasas)) < 100:
        raise ValueError(f"El informe {ruta.name} no contiene una única estimación "
                         f"válida para {nombre_mes} de {anio}: {sorted(tasas)}.")
    return tasas.pop(), str(ruta.resolve())


def completar_reales_atp(ruta: Path, hoja_nombre: str, periodo: str,
                         tasa: float, url: str) -> dict:
    libro = load_workbook(ruta, data_only=False, keep_links=True)
    hoja = libro[hoja_nombre]
    pedidos = {"periodo": "PERIODO", "nea": "Indice de precios al consumidor- Nivel general (Región NEA)",
               **{f"real_{i}": titulo for i, titulo in enumerate(TITULOS_REALES)}}
    columnas = resolver_columnas(hoja, pedidos)
    fila = buscar_fila(hoja, columnas["periodo"], periodo)
    anio, mes = map(int, periodo.split("-"))
    anterior = buscar_fila(hoja, columnas["periodo"],
                          f"{anio - 1:04d}-12" if mes == 1
                          else f"{anio:04d}-{mes - 1:02d}")
    interanual = buscar_fila(hoja, columnas["periodo"], f"{anio - 1:04d}-{mes:02d}")
    indice = columnas["nea"]
    if not all(isinstance(hoja.cell(r, indice).value, (int, float)) and
               hoja.cell(r, indice).value > 0 for r in (anterior, interanual)):
        raise ValueError("Falta el IPC NEA del mes anterior o de igual mes del año anterior.")
    cambios = []
    # Los dos coeficientes son los únicos supuestos; los demás son fórmulas existentes.
    mensual = hoja.cell(fila, columnas["real_0"])
    anual = hoja.cell(fila, columnas["real_1"])
    indice_anterior = hoja.cell(anterior, indice).coordinate
    indice_interanual = hoja.cell(interanual, indice).coordinate
    # Tasa usada hasta ahora en el coeficiente (para informar si IARAF la cambió).
    previa = re.search(r"1\+([0-9]+(?:\.[0-9]+)?)/100", str(mensual.value or ""))
    tasa_anterior = float(previa.group(1)) if previa else None
    esperadas = {
        mensual.coordinate: f"=1/(1+{tasa:g}/100)",
        anual.coordinate: f"={indice_interanual}/({indice_anterior}*(1+{tasa:g}/100))",
    }
    for i in range(2, len(TITULOS_REALES)):
        destino = hoja.cell(fila, columnas[f"real_{i}"])
        if isinstance(destino.value, str) and destino.value.startswith("="):
            continue
        origen = hoja.cell(fila - 1, destino.column)
        if not isinstance(origen.value, str) or not origen.value.startswith("="):
            raise ValueError(f"No existe fórmula anterior para {destino.coordinate}.")
        esperadas[destino.coordinate] = Translator(
            origen.value, origin=origen.coordinate).translate_formula(destino.coordinate)
    for direccion, formula in esperadas.items():
        celda = hoja[direccion]
        if celda.value == formula:
            continue
        if celda.value not in (None, 0, "") and not (
            isinstance(celda.value, str) and celda.value.startswith("=")
        ):
            raise ValueError(f"{direccion} contiene un valor manual; no se sobrescribe.")
        # Se preservan fórmulas existentes, salvo los dos supuestos que se actualizan
        # con la fuente del período.
        if celda.value not in (None, 0, "") and direccion not in (
            mensual.coordinate, anual.coordinate):
            continue
        cambios.append((direccion, formula))
    if cambios:
        respaldo = ruta.parent / "respaldos"
        respaldo.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ruta, respaldo / f"antes_iaraf_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        for direccion, formula in cambios:
            hoja[direccion] = formula
        for celda in (mensual, anual):
            celda.comment = __import__("openpyxl").comments.Comment(
                f"Estimación IARAF {periodo}: {tasa:g}% mensual. Fuente: {url}. "
                "Uso provisional hasta IPC NEA oficial.", "Observatorio"
            )
        libro.calculation.fullCalcOnLoad = True
        libro.calculation.forceFullCalc = True
        libro.calculation.calcMode = "auto"
        guardar_atomico(libro, ruta)
    else:
        libro.close()
    # Verificación: el coeficiente guardado debe usar la tasa de IARAF del período.
    control = load_workbook(ruta, data_only=False, read_only=True, keep_links=True)
    try:
        guardado = str(control[hoja_nombre][mensual.coordinate].value or "")
    finally:
        control.close()
    if guardado.replace(" ", "") != f"=1/(1+{tasa:g}/100)":
        raise RuntimeError(f"El coeficiente {mensual.coordinate} no quedó con la inflación IARAF {tasa:g}%: {guardado!r}")
    return {"estado": "OK", "periodo": periodo, "tasa_estimada": tasa, "tasa_anterior": tasa_anterior,
            "fuente_iaraf": url,
            "formulas_completadas": [direccion for direccion, _ in cambios]}


def actualizar_base(ruta: Path, hoja_nombre: str, dato: dict, sobrescribir: bool) -> dict:
    libro = load_workbook(ruta, data_only=False, keep_links=True)
    if hoja_nombre not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {hoja_nombre!r}.")
    hoja = libro[hoja_nombre]
    pedidos = {"periodo": "PERIODO"}
    pedidos.update({f"atp_{n}": titulo for n, titulo in enumerate(TITULOS_ATP.values())})
    columnas = resolver_columnas(hoja, pedidos)
    fila = buscar_fila(hoja, columnas["periodo"], dato["periodo"])
    cambios = []
    coincidencias = []
    asignaciones = [
        (titulo, hoja.cell(fila, columnas[f"atp_{n}"]), float(dato["valores"][titulo]))
        for n, titulo in enumerate(TITULOS_ATP.values())
    ]

    for titulo, celda, nuevo in asignaciones:
        anterior = celda.value
        if isinstance(anterior, str) and anterior.startswith("="):
            raise ValueError(f"Seguridad: {celda.coordinate} contiene una fórmula.")
        if iguales(anterior, nuevo):
            coincidencias.append(celda.coordinate)
            continue
        vacio_o_cero = anterior in (None, "") or (
            isinstance(anterior, (int, float))
            and not isinstance(anterior, bool)
            and float(anterior) == 0.0
        )
        if not vacio_o_cero and not sobrescribir:
            raise ValueError(
                f"{celda.coordinate} ya contiene {anterior!r}, distinto de {nuevo}. "
                "Revisá el dato o usá --sobrescribir."
            )
        cambios.append({
            "celda": celda.coordinate,
            "titulo": titulo,
            "antes": anterior,
            "despues": nuevo,
        })

    if not cambios:
        libro.close()
        return {
            "archivo_excel": str(ruta.resolve()), "hoja": hoja_nombre,
            "periodo": dato["periodo"], "fila_localizada": fila,
            "celdas_actualizadas": 0, "hubo_cambios": False,
            "cambios": [], "coincidencias": coincidencias, "respaldo": None,
        }

    carpeta = ruta.parent / "respaldos"
    carpeta.mkdir(parents=True, exist_ok=True)
    respaldo = carpeta / f"antes_atp_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(ruta, respaldo)

    for cambio in cambios:
        celda = hoja[cambio["celda"]]
        celda.value = cambio["despues"]
        if not celda.number_format or celda.number_format == "General":
            celda.number_format = copy(hoja.cell(fila - 1, celda.column).number_format)
    libro.calculation.fullCalcOnLoad = True
    libro.calculation.forceFullCalc = True
    libro.calculation.calcMode = "auto"
    guardar_atomico(libro, ruta)

    control = load_workbook(ruta, data_only=False, read_only=True, keep_links=True)
    try:
        hoja_control = control[hoja_nombre]
        guardados = {
            cambio["celda"]: hoja_control[cambio["celda"]].value
            for cambio in cambios
        }
    finally:
        control.close()
    for cambio in cambios:
        if not iguales(guardados[cambio["celda"]], float(cambio["despues"])):
            raise RuntimeError(f"Falló la verificación de {cambio['celda']}.")

    return {
        "archivo_excel": str(ruta.resolve()), "hoja": hoja_nombre,
        "periodo": dato["periodo"], "fila_localizada": fila,
        "celdas_actualizadas": len(cambios), "hubo_cambios": True,
        "cambios": cambios, "coincidencias": coincidencias,
        "respaldo": str(respaldo.resolve()),
    }


def main() -> int:
    configurar_salida()
    parser = argparse.ArgumentParser(description="Actualiza la recaudación tributaria de ATP Formosa.")
    parser.add_argument("--excel", type=Path, default=BASE_PREDETERMINADA)
    parser.add_argument("--hoja", default=HOJA_PREDETERMINADA)
    parser.add_argument("--html", type=Path, help="HTML local, únicamente para pruebas.")
    parser.add_argument("--anio", type=int, default=date.today().year)
    parser.add_argument("--sobrescribir", action="store_true")
    parser.add_argument("--informe-iaraf", type=Path,
                        help="PDF local de IARAF; por defecto fuentes/iaraf/"
                             "iaraf_transferencias_AAAA_MM.pdf.")
    args = parser.parse_args()

    ruta = args.excel.expanduser().resolve()
    if not ruta.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta}")
    html = args.html.read_text(encoding="utf-8") if args.html else descargar_html()
    dato = extraer_ultimo_periodo(html, args.anio)
    actualizacion = actualizar_base(ruta, args.hoja, dato, args.sobrescribir)
    # Los valores reales dependen de la inflación supuesta por IARAF. Si todavía
    # no está, se cargan igual los datos nominales y lo real queda pendiente.
    try:
        tasa, url_iaraf = estimacion_iaraf(dato["anio"], dato["mes"], args.informe_iaraf)
        ajuste_real = completar_reales_atp(
            ruta, args.hoja, dato["periodo"], tasa, url_iaraf)
    except Exception as error:
        ajuste_real = {"estado": "PENDIENTE", "detalle": str(error)}

    resultado = {
        "fuente": "ATP - Recaudación Tributaria Provincial",
        "url": URL_FUENTE,
        "consultado_en": datetime.now().astimezone().isoformat(timespec="seconds"),
        "periodo": dato["periodo"],
        "recaudacion": dato["recaudacion"],
        "valores": dato["valores"],
        "actualizacion": actualizacion,
        "ajuste_real": ajuste_real,
    }
    carpeta_reporte = RAIZ / "scripts" / "reportes" / "atp"
    carpeta_reporte.mkdir(parents=True, exist_ok=True)
    reporte = carpeta_reporte / f"atp_{dato['anio']:04d}_{dato['mes']:02d}.json"
    reporte.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    resultado["reporte"] = str(reporte.resolve())

    print("RESULTADO_ATP=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
    print("ESTADO=OK")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        raise SystemExit(1)
