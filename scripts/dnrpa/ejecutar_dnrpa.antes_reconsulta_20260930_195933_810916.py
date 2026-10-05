from __future__ import annotations

import argparse
import html
import json
import re
import sys
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
from comun.columnas_por_titulo import letra_columna, resolver_columnas


URL = "https://www.dnrpa.gov.ar/portal_dnrpa/estadisticas/rrss_tramites/tram_prov.php"
HOJA = "BASE DE DATOS"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
@dataclass(frozen=True)
class Consulta:
    nombre: str
    consulta: str
    tipo: str


CONSULTAS = (
    Consulta("patentamientos_autos", "inscripciones", "A"),
    Consulta("transferencias_autos", "transferencias", "A"),
    Consulta("patentamientos_motos", "inscripciones", "M"),
    Consulta("transferencias_motos", "transferencias", "M"),
)

DESTINOS = {
    "patentamientos_autos.formosa": "Patentamiento de automóviles F",
    "patentamientos_autos.nacion": "Patentamiento de automóviles N",
    "transferencias_autos.formosa": "Transferencias de automoviles F",
    "transferencias_autos.nacion": "Transferencias de automoviles N",
    "patentamientos_motos.formosa": "Patentamiento de Motos F",
    "patentamientos_motos.nacion": "Patentamiento de Motos N",
    "transferencias_motos.formosa": "Transferencias de Motos F",
    "transferencias_motos.nacion": "Transferencias de Motos N",
}


def raiz_proyecto() -> Path:
    return Path(__file__).resolve().parents[2]


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._table: list[list[str]] | None = None
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.lower()
        if tag == "table":
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in {"td", "th"} and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"} and self._cell is not None and self._row is not None:
            value = " ".join("".join(self._cell).split())
            self._row.append(html.unescape(value))
            self._cell = None
        elif tag == "tr" and self._row is not None and self._table is not None:
            if self._row:
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            if self._table:
                self.tables.append(self._table)
            self._table = None


def numero(texto: str) -> int:
    limpio = re.sub(r"[^0-9-]", "", texto)
    if not limpio or limpio == "-":
        raise ValueError(f"Valor numérico inválido: {texto!r}")
    return int(limpio)


def extraer_tabla(contenido: str) -> dict[str, list[int]]:
    parser = TableParser()
    parser.feed(contenido)
    for tabla in parser.tables:
        filas = {}
        for fila in tabla:
            if len(fila) < 14:
                continue
            nombre = re.sub(r"\s+", " ", fila[0]).strip().upper()
            if nombre in {"FORMOSA", "TOTAL", "TOTAL PAIS", "TOTAL PAÍS"}:
                # Doce meses más la columna final "Total" (acumulado anual publicado).
                filas["TOTAL" if nombre.startswith("TOTAL") else "FORMOSA"] = [numero(x) for x in fila[1:14]]
        if {"FORMOSA", "TOTAL"}.issubset(filas):
            # La fila TOTAL es el valor nacional publicado. No exigir que
            # coincida con la suma de filas visibles: inscripciones distingue
            # categorías cuya tabla provincial puede tener otra cobertura.
            return filas
    raise ValueError("No se encontró una tabla con las filas FORMOSA y TOTAL")


def descargar(anio: int, consulta: str, tipo: str, timeout: int = 45) -> str:
    query = urllib.parse.urlencode({"origen": "portal_dnrpa", "tipo_consulta": consulta})
    payload = urllib.parse.urlencode(
        {
            "anio": str(anio),
            "codigo_tipo": tipo,
            "operacion": "1",
            "tipo_consulta": consulta,
            "origen": "portal_dnrpa",
            "boton": "Aceptar",
        }
    ).encode("ascii")
    req = urllib.request.Request(
        f"{URL}?{query}",
        data=payload,
        headers={"User-Agent": "Mozilla/5.0 (automatizacion-indicadores/1.0)", "Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as respuesta:
        raw = respuesta.read()
        charset = respuesta.headers.get_content_charset() or "latin-1"
    return raw.decode(charset, errors="replace")


def ultimo_mes_publicado(series: dict[str, list[int]]) -> int:
    indices = [i for i in range(12) if series["FORMOSA"][i] != 0 or series["TOTAL"][i] != 0]
    if not indices:
        raise ValueError("La tabla no contiene meses publicados")
    return max(indices)


def validar_transferencias(tabla: dict[str, list[int]], indice: int, cfg: Consulta) -> None:
    """Rechaza un salto simultáneo incompatible con la serie publicada.

    La fuente ha publicado un acumulado en la columna mensual. Se intenta
    normalizar por separado; si no cumple la identidad exacta, se detiene.
    """
    if cfg.consulta != "transferencias" or indice < 3:
        return
    anteriores_f = [v for v in tabla["FORMOSA"][max(0, indice - 6):indice] if v > 0]
    anteriores_n = [v for v in tabla["TOTAL"][max(0, indice - 6):indice] if v > 0]
    if len(anteriores_f) < 3 or len(anteriores_n) < 3:
        return
    formosa, nacion = tabla["FORMOSA"][indice], tabla["TOTAL"][indice]
    if formosa > 2 * max(anteriores_f) and nacion > 2 * max(anteriores_n):
        raise ValueError(
            f"DNRPA: respuesta anómala para {cfg.nombre}, {MESES[indice]}: "
            f"Formosa={formosa} (máximo seis meses anteriores={max(anteriores_f)}), "
            f"Nación={nacion} (máximo={max(anteriores_n)}). "
            "No se pudo validar si la columna es mensual o acumulada; "
            "se suspendió la carga para proteger la base."
        )


def normalizar_acumulado_en_mes(tabla: dict[str, list[int]], indice: int, cfg: Consulta) -> bool:
    """Corrige la tabla oficial cuando el último mes contiene el acumulado del año.

    En ese caso el 'Total' publicado suma meses previos y el acumulado una
    segunda vez. Solo se acepta si ambas filas exhiben exactamente ese patrón.
    """
    if cfg.consulta != "transferencias" or indice < 2:
        return False
    previos = {k: sum(tabla[k][:indice]) for k in ("FORMOSA", "TOTAL")}
    anomalia = all(
        tabla[k][indice] > 2 * max(tabla[k][max(0, indice-6):indice])
        and tabla[k][12] == previos[k] + tabla[k][indice]
        and tabla[k][indice] > previos[k]
        for k in previos
    )
    if not anomalia:
        return False
    for k in previos:
        tabla[k][12] = tabla[k][indice]
        tabla[k][indice] -= previos[k]
    return True


def ultimo_mes_cargado(ruta_base: Path, anio: int) -> int:
    if not ruta_base.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta_base}")
    ws = load_workbook(ruta_base, read_only=True, data_only=True)[HOJA]
    columnas = resolver_columnas(ws, DESTINOS)
    meses = []
    for fila in ws.iter_rows(min_row=2):
        fecha = fila[0].value
        if not isinstance(fecha, datetime) or fecha.year != anio:
            continue
        if all(isinstance(fila[col - 1].value, (int, float))
               and fila[col - 1].value > 0 for col in columnas.values()):
            meses.append(fecha.month)
    return max(meses, default=0)


def extraer(anio: int, mes_cargado: int = 0) -> dict:
    resultado = {"fuente": "DNRPA", "anio": anio, "consultado_en": datetime.now().isoformat(timespec="seconds"), "series": {}}
    meses_detectados = set()
    problemas = []
    consultas = []
    for cfg in CONSULTAS:
        try:
            tabla = extraer_tabla(descargar(anio, cfg.consulta, cfg.tipo))
        except ValueError as error:
            raise ValueError(
                f"DNRPA {cfg.consulta}, vehículo {cfg.tipo}, año {anio}: {error}"
            ) from error
        indice = ultimo_mes_publicado(tabla)
        consultas.append((cfg, tabla, indice))
        meses_detectados.add(indice)
    resultado["meses_consultados"] = {cfg.nombre: MESES[indice]
                                     for cfg, _, indice in consultas}
    ultimo = max(meses_detectados)
    resultado["mes"] = MESES[ultimo]
    resultado["mes_numero"] = ultimo + 1
    if ultimo + 1 <= mes_cargado:
        resultado["estado_publicacion"] = "SIN_NOVEDADES"
        return resultado
    if len(meses_detectados) != 1:
        resultado["estado_publicacion"] = "PENDIENTE_CUATRO_CONSULTAS"
        return resultado
    for cfg, tabla, indice in consultas:
        corregido = normalizar_acumulado_en_mes(tabla, indice, cfg)
        try:
            validar_transferencias(tabla, indice, cfg)
        except ValueError as error:
            problemas.append(str(error))
        if corregido:
            # Un acumulado cargado durante el mes puede ser parcial. Se sigue
            # consultando en cada ejecución, pero no se publica hasta que sea
            # comparable con la serie nacional reciente.
            anteriores = [v for v in tabla["TOTAL"][max(0, indice-6):indice] if v > 0]
            if anteriores and tabla["TOTAL"][indice] < min(anteriores) // 2:
                problemas.append(
                    f"DNRPA {cfg.nombre}, {MESES[indice]}: publicación parcial "
                    f"({tabla['TOTAL'][indice]} nacional tras descontar el acumulado). "
                    "Se volverá a consultar en la próxima ejecución; no se escribió la base."
                )
        resultado["series"][cfg.nombre] = {
            "mes": MESES[indice],
            "mes_numero": indice + 1,
            "formosa": tabla["FORMOSA"][indice],
            "nacion": tabla["TOTAL"][indice],
            "acumulado_formosa": tabla["FORMOSA"][12],
            "acumulado_nacion": tabla["TOTAL"][12],
            "columna_publicada_como_acumulado": corregido,
        }

        calculado_formosa = sum(tabla["FORMOSA"][: indice + 1])
        calculado_nacion = sum(tabla["TOTAL"][: indice + 1])
        if calculado_formosa != tabla["FORMOSA"][12] or calculado_nacion != tabla["TOTAL"][12]:
            problemas.append(
                f"El acumulado publicado no coincide con la suma mensual en {cfg.nombre}: "
                f"Formosa {tabla['FORMOSA'][12]} != {calculado_formosa}; "
                f"Nación {tabla['TOTAL'][12]} != {calculado_nacion}"
            )
    if len(meses_detectados) != 1:
        detalle = {nombre: dato["mes"] for nombre, dato in resultado["series"].items()}
        problemas.append(f"Las cuatro consultas no tienen el mismo último mes: {detalle}")
    if problemas:
        raise ValueError("; ".join(problemas))
    indice = meses_detectados.pop()
    resultado["mes"] = MESES[indice]
    resultado["mes_numero"] = indice + 1
    resultado["estado_publicacion"] = "NUEVA"
    return resultado


def localizar_fila(ws, anio: int, mes: int) -> int:
    for fila in range(2, ws.max_row + 1):
        valor = ws.cell(fila, 1).value
        if isinstance(valor, datetime) and valor.year == anio and valor.month == mes:
            return fila
    raise ValueError(f"No existe el período {anio}-{mes:02d} en la columna A")


def actualizar_base(resultado: dict, ruta_base: Path) -> dict:
    if not ruta_base.is_file():
        raise FileNotFoundError(f"No existe la base: {ruta_base}")
    libro = load_workbook(ruta_base, data_only=False)
    if HOJA not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {HOJA!r}")
    ws = libro[HOJA]
    columnas = resolver_columnas(ws, DESTINOS)
    fila = localizar_fila(ws, resultado["anio"], resultado["mes_numero"])
    cambios, coincidencias = [], []
    for cfg in CONSULTAS:
        datos = resultado["series"][cfg.nombre]
        asignaciones = (
            (f"{cfg.nombre}.formosa", datos["formosa"]),
            (f"{cfg.nombre}.nacion", datos["nacion"]),
        )
        for clave, valor in asignaciones:
            columna = columnas[clave]
            celda = ws.cell(fila, columna)
            anterior = celda.value
            if anterior == valor:
                coincidencias.append(celda.coordinate)
            else:
                celda.value = valor
                cambios.append({"celda": celda.coordinate, "titulo": DESTINOS[clave], "antes": anterior, "despues": valor})
    if cambios:
        libro.save(ruta_base)
    return {"archivo_excel": str(ruta_base.resolve()), "fila": fila,
            "celdas_actualizadas": len(cambios), "cambios": cambios,
            "coincidencias": coincidencias}


def prueba_parser() -> None:
    muestra = """<table><tr><th>Provincia / Mes</th><th>Ene</th><th>Feb</th><th>Mar</th><th>Abr</th><th>May</th><th>Jun</th><th>Jul</th><th>Ago</th><th>Sep</th><th>Oct</th><th>Nov</th><th>Dic</th><th>Total</th></tr><tr><td>FORMOSA</td><td>875</td><td>609</td><td>743</td><td>681</td><td>579</td><td>679</td><td>768</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>4.934</td></tr><tr><td>TOTAL</td><td>153.361</td><td>130.979</td><td>154.229</td><td>154.996</td><td>144.218</td><td>155.778</td><td>157.393</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>1.050.954</td></tr></table>"""
    tabla = extraer_tabla(muestra)
    assert tabla["FORMOSA"][6] == 768 and tabla["TOTAL"][6] == 157393
    assert tabla["FORMOSA"][12] == 4934 and tabla["TOTAL"][12] == 1050954
    assert ultimo_mes_publicado(tabla) == 6
    print("PRUEBA_PARSER=OK")


def main() -> int:
    parser = argparse.ArgumentParser(description="Detecta el último período DNRPA y actualiza la base.")
    parser.add_argument("--anio", type=int, default=datetime.now().year)
    parser.add_argument("--probar-parser", action="store_true")
    parser.add_argument("--solo-extraer", action="store_true")
    parser.add_argument("--base", type=Path, default=raiz_proyecto() / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx")
    args = parser.parse_args()
    if args.probar_parser:
        prueba_parser()
        return 0
    try:
        mes_cargado = ultimo_mes_cargado(args.base, args.anio)
        resultado = extraer(args.anio, mes_cargado)
        if not args.solo_extraer and resultado["estado_publicacion"] == "NUEVA":
            resultado["actualizacion"] = actualizar_base(resultado, args.base)
        elif not args.solo_extraer:
            resultado["actualizacion"] = {"celdas_actualizadas": 0, "cambios": []}
        print(json.dumps(resultado, ensure_ascii=False, indent=2))
        print("ESTADO=OK")
        return 0
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 1


if __name__ == "__main__":
    sys.exit(main())
