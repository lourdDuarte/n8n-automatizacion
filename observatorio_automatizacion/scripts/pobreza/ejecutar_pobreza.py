from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path

import xlrd
from openpyxl.utils import get_column_letter


RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "complejos"))

from comun import (  # noqa: E402
    ARCHIVO_BASE,
    abrir_base,
    descargar,
    emitir,
    fila_periodo,
    guardar_reporte,
    guardar_si_cambia,
    normalizar,
)


URL_CUADROS = "https://www.indec.gob.ar/ftp/cuadros/sociedad/cuadros_informe_pobreza_{mes:02d}_{anio:02d}.xls"


TITULOS_ORIGINALES = {
    "pobreza_personas_f": "Pobreza (cant. personas) F*",
    "pobreza_personas_n": "Pobreza (cant. personas) N*",
    "pobreza_porcentaje_personas_f": "Pobreza (% sobre el total de personas) F*",
    "pobreza_porcentaje_personas_n": "Pobreza (% sobre el total de personas) N*",
    "pobreza_hogares_f": "Pobreza (cant. hogares) F*",
    "pobreza_hogares_n": "Pobreza (cant. hogares) N*",
    "pobreza_porcentaje_hogares_f": "Pobreza (% sobre el total de hogares) F*",
    "pobreza_porcentaje_hogares_n": "Pobreza (% sobre el total de hogares) N*",
    "indigencia_personas_f": "Indigencia (cant. personas) F*",
    "indigencia_personas_n": "Indigencia (cant. personas) N*",
    "indigencia_porcentaje_personas_f": "Indigencia (% sobre el total de personas) F*",
    "indigencia_porcentaje_personas_n": "Indigencia (% sobre el total de personas) N*",
    "indigencia_hogares_f": "Indigencia (cant. hogares) F*",
    "indigencia_hogares_n": "Indigencia (cant. hogares) N*",
    "indigencia_porcentaje_hogares_f": "Indigencia (% sobre el total de hogares) F*",
    "indigencia_porcentaje_hogares_n": "Indigencia (% sobre el total de hogares) N*",
}


FORMULAS = {
    "Var pp. Pobreza (% sobre el total de personas) F*":
        "Pobreza (% sobre el total de personas) F*",
    "Var pp. Pobreza (% sobre el total de personas) N*":
        "Pobreza (% sobre el total de personas) N*",
    "Var pp. Pobreza (% sobre el total de hogares) F*":
        "Pobreza (% sobre el total de hogares) F*",
    "Var pp. Indigencia (% sobre el total de personas) F*":
        "Indigencia (% sobre el total de personas) F*",
    "Var pp. Indigencia (% sobre el total de personas) N*":
        "Indigencia (% sobre el total de personas) N*",
    "Var pp.Indigencia (% sobre el total de hogares) F*":
        "Indigencia (% sobre el total de hogares) F*",
}

# Columnas de variación agregadas por el equipo en el Sheet: se completan si
# existen y se omiten (sin error) si todavía no están.
FORMULAS_OPCIONALES = {
    "Var pp. Pobreza (% sobre el total de hogares) N*":
        "Pobreza (% sobre el total de hogares) N*",
    "Var pp. Indigencia (% sobre el total de hogares) N*":
        "Indigencia (% sobre el total de hogares) N*",
}


def columna_objetivo(ws, titulo: str) -> int:
    coincidencias = [
        columna
        for columna in range(1, ws.max_column + 1)
        if normalizar(ws.cell(1, columna).value) == normalizar(titulo)
    ]
    if not coincidencias:
        raise ValueError(f"No se encontró la columna titulada {titulo!r}")
    if len(coincidencias) > 1:
        raise ValueError(f"El encabezado objetivo está duplicado: {titulo!r}")
    return coincidencias[0]


def escribir_objetivo(ws, fila: int, titulo: str, nuevo, cambios: list, coincidencias: list) -> str:
    columna = columna_objetivo(ws, titulo)
    celda = ws.cell(fila, columna)
    anterior = celda.value
    iguales = anterior == nuevo
    if isinstance(anterior, (int, float)) and isinstance(nuevo, (int, float)):
        iguales = abs(float(anterior) - float(nuevo)) <= 1e-7
    if iguales:
        coincidencias.append(celda.coordinate)
    else:
        celda.value = nuevo
        cambios.append({
            "celda": celda.coordinate,
            "titulo": titulo,
            "antes": anterior,
            "despues": nuevo,
        })
    return celda.coordinate


def periodo_cuadros(libro) -> tuple[date, str]:
    titulos = " ".join(str(libro.sheet_by_name(n).cell_value(0, 0)) for n in ("Cuadro 4.1", "Cuadro 4.2"))
    hallados = re.findall(r"(primer|segundo)\s+semestre\s+de\s+(20\d{2})", normalizar(titulos))
    if len(hallados) != 2 or hallados[0] != hallados[1]:
        raise ValueError("Los cuadros oficiales no identifican el mismo semestre")
    semestre = 1 if hallados[0][0] == "primer" else 2
    anio = int(hallados[0][1])
    return date(anio, 6 if semestre == 1 else 12, 1), f"{semestre}S {anio}"


def abrir_cuadros(contenido: bytes):
    if not contenido.startswith(b"\xd0\xcf\x11\xe0"):
        raise ValueError("La descarga no es un archivo XLS del INDEC")
    return xlrd.open_workbook(file_contents=contenido)


def descubrir() -> tuple[str, bytes, date, str]:
    # Una vez iniciado el mes habitual de publicación, nunca reutilizar el
    # archivo del semestre anterior como si fuera el informe nuevo.
    hoy = date.today()
    if hoy.month >= 9:
        mes, anio_archivo, esperado = 9, hoy.year, date(hoy.year, 6, 1)
    elif hoy.month >= 3:
        mes, anio_archivo, esperado = 3, hoy.year, date(hoy.year - 1, 12, 1)
    else:
        mes, anio_archivo, esperado = 9, hoy.year - 1, date(hoy.year - 1, 6, 1)
    url = URL_CUADROS.format(mes=mes, anio=anio_archivo % 100)
    # Además del nombre conocido se prueban los cuadros de la página de pobreza del INDEC,
    # por si cambian el nombre del archivo; se usa el del semestre más reciente.
    sys.path.insert(0, str(RAIZ / "scripts" / "comun"))
    from fuentes_web import PAGINAS_INDEC, candidatos, elegir_mas_reciente

    def procesar(contenido_, url_):
        periodo_, etiqueta_ = periodo_cuadros(abrir_cuadros(contenido_))
        return periodo_, (url_, contenido_, periodo_, etiqueta_)
    try:
        urls = candidatos([url], PAGINAS_INDEC["pobreza"], incluir=("pobreza",), excluir=("coeficiente",), extensiones=(".xls",))
        (url, contenido, periodo, etiqueta), _, _ = elegir_mas_reciente(urls, procesar)
    except Exception as exc:
        raise FileNotFoundError(
            f"No está disponible el cuadro oficial del período {esperado}: {url}. "
            "No se reutiliza un informe anterior."
        ) from exc
    if periodo != esperado:
        raise ValueError(f"El archivo anuncia {etiqueta}, se esperaba {esperado}; no se carga")
    return url, contenido, periodo, etiqueta


def extraer_de_documento(contenido: bytes) -> dict:
    libro = abrir_cuadros(contenido)
    periodo_cuadros(libro)
    valores = {}
    for concepto, nombre in (("pobreza", "Cuadro 4.1"), ("indigencia", "Cuadro 4.2")):
        hoja = libro.sheet_by_name(nombre)
        encabezados = [normalizar(str(hoja.cell_value(r, c))) for r in range(min(7, hoja.nrows)) for c in range(hoja.ncols)]
        if not {"hogares", "personas", "incidencia"}.issubset(set(encabezados)) or not any("en miles" in h for h in encabezados):
            raise ValueError(f"Cambió la estructura de {nombre}; no se cargarán datos dudosos")
        for sufijo, rotulo in (("f", "formosa"), ("n", "total 31 aglomerados urbanos")):
            filas = [r for r in range(hoja.nrows) if normalizar(str(hoja.cell_value(r, 0))).strip() == rotulo]
            if len(filas) != 1:
                raise ValueError(f"No se encontró una única fila {rotulo} en {nombre}")
            r = filas[0]
            # B/C: hogares; F/G: personas. Cantidades publicadas en miles.
            for unidad, col_cantidad, col_tasa in (("hogares", 1, 2), ("personas", 5, 6)):
                cantidad = float(hoja.cell_value(r, col_cantidad))
                tasa = float(hoja.cell_value(r, col_tasa))
                if cantidad < 0 or not 0 <= tasa <= 100:
                    raise ValueError(f"Valores inválidos en {nombre}, {rotulo}")
                valores[f"{concepto}_{unidad}_{sufijo}"] = round(cantidad * 1000)
                valores[f"{concepto}_porcentaje_{unidad}_{sufijo}"] = tasa
    return valores


def validar(valores: dict) -> None:
    faltantes = [clave for clave in TITULOS_ORIGINALES if clave not in valores]
    if faltantes:
        raise ValueError("Faltan valores: " + ", ".join(faltantes))
    for clave, valor in valores.items():
        if "porcentaje" in clave and not 0 <= float(valor) <= 100:
            raise ValueError(f"Porcentaje fuera de rango: {clave}={valor}")
        if "porcentaje" not in clave and int(valor) < 0:
            raise ValueError(f"Cantidad negativa: {clave}={valor}")


def main() -> int:
    ap = argparse.ArgumentParser(description="Actualiza pobreza e indigencia semestral de INDEC")
    ap.add_argument("--base", type=Path, default=RAIZ / "base" / ARCHIVO_BASE)
    ap.add_argument("--archivo", type=Path, help="XLS oficial local, solo para prueba")
    ap.add_argument("--solo-extraer", action="store_true")
    args = ap.parse_args()

    if args.archivo:
        contenido = args.archivo.read_bytes()
        periodo, etiqueta = periodo_cuadros(abrir_cuadros(contenido))
        url = str(args.archivo.resolve())
    else:
        url, contenido, periodo, etiqueta = descubrir()

    valores = extraer_de_documento(contenido)
    validar(valores)
    valores_por_titulo = {
        TITULOS_ORIGINALES[clave]: valor for clave, valor in valores.items()
    }
    resultado = {
        "fuente": "Pobreza e indigencia",
        "periodo": periodo.isoformat(),
        "semestre": etiqueta,
        "url_informe": url,
        "nota_cantidades": "INDEC publica las cantidades en miles redondeadas; se multiplicaron por 1000.",
        "valores": valores_por_titulo,
    }

    if not args.solo_extraer:
        libro, ws = abrir_base(args.base)
        fila = fila_periodo(ws, periodo.year, periodo.month)
        fila_anterior = fila_periodo(ws, periodo.year - 1, 12) if periodo.month == 6 else fila_periodo(ws, periodo.year, 6)
        cambios = []
        coincidencias = []

        for titulo, valor in valores_por_titulo.items():
            escribir_objetivo(ws, fila, titulo, valor, cambios, coincidencias)

        formulas_escritas = {}
        todas = dict(FORMULAS)
        for titulo_formula, titulo_dato in FORMULAS_OPCIONALES.items():
            if any(normalizar(ws.cell(1, c).value) == normalizar(titulo_formula) for c in range(1, ws.max_column + 1)):
                todas[titulo_formula] = titulo_dato
        for titulo_formula, titulo_dato in todas.items():
            columna_dato = columna_objetivo(ws, titulo_dato)
            letra = get_column_letter(columna_dato)
            formula = f"={letra}{fila}-{letra}{fila_anterior}"
            escribir_objetivo(ws, fila, titulo_formula, formula, cambios, coincidencias)
            formulas_escritas[titulo_formula] = formula

        # El ejecutor general utiliza resultado["valores"] para construir el mapeo
        # destinado a Google Sheets. Se incorporan aquí las fórmulas con sus nombres
        # de encabezado para que el Sheet reciba las 22 columnas, no solo los 16
        # valores publicados por INDEC.
        resultado["valores"].update(formulas_escritas)

        respaldo = guardar_si_cambia(libro, args.base, cambios, "pobreza")
        resultado["actualizacion"] = {
            "archivo_excel": str(args.base.resolve()),
            "fila": fila,
            "fila_semestre_anterior": fila_anterior,
            "celdas_actualizadas": len(cambios),
            "cambios": cambios,
            "coincidencias": coincidencias,
            "formulas": formulas_escritas,
            "respaldo": respaldo,
        }

    resultado["reporte"] = str(
        guardar_reporte(RAIZ, "pobreza", periodo, resultado).resolve()
    )
    emitir(resultado)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR={exc}")
        print("ESTADO=ERROR")
        raise SystemExit(1)
