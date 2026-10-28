"""Tablas para el tablero de Power BI de seguridad social (BESS y ANSES) - Formosa.

Arma, en la carpeta powerbi\\ de la automatización, tablas en formato "largo" (una fila por dato)
para que Power BI no dependa de los encabezados de la hoja BASE DE DATOS:

    seguridad_social.csv          BESS y ANSES de la BASE DE DATOS del Sheet (Formosa y Nación),
                                  una fila por período, indicador y territorio.
    bess_departamentos.csv        BESS, beneficios por departamento de Formosa (cuadros 6.1 y 6.2
                                  del boletín provincial por departamento, Secretaría de Seguridad
                                  Social). Es la única fuente oficial con apertura por departamento.
    poblacion_departamentos.csv   INDEC, Censo 2022, población por departamento (fija).

El mapa de los departamentos (formosa_departamentos.topojson, IGN) se copia una sola vez.

Se ejecuta solo al final de cada revisión que incluye BESS o ANSES (ejecutar_general.py).
Uso manual:  python exportar_seguridad_social.py
"""
from __future__ import annotations

import csv
import io
import json
import re
import sys
import unicodedata
import urllib.request
from datetime import date, datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))

CARPETA = RAIZ / 'powerbi'
URL_DEPTOS = 'https://www.argentina.gob.ar/sites/default/files/boletin_provincias_departamentos_{anio}.xlsx'

# INDEC, Censo Nacional de Población, Hogares y Viviendas 2022, resultados definitivos,
# Cuadro 1.9 Provincia de Formosa (c2022_formosa_est_c1_9.xlsx). Código = código INDEC/IGN.
POBLACION_2022 = [
    ('34007', 'Bermejo', 15644), ('34014', 'Formosa', 269589), ('34021', 'Laishi', 19303),
    ('34028', 'Matacos', 17032), ('34035', 'Patiño', 81819), ('34042', 'Pilagás', 19116),
    ('34049', 'Pilcomayo', 94383), ('34056', 'Pirané', 72804), ('34063', 'Ramón Lista', 17729),
]
FUENTE_POBLACION = 'INDEC, Censo 2022, resultados definitivos (cuadro 1.9 Formosa)'
MESES = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6, 'JULIO': 7,
         'AGOSTO': 8, 'SEPTIEMBRE': 9, 'SETIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}


def norm(x) -> str:
    s = unicodedata.normalize('NFD', str(x or ''))
    return ' '.join(''.join(c for c in s if unicodedata.category(c) != 'Mn').split()).upper()


CODIGOS = {norm(n): c for c, n, _ in POBLACION_2022}


def escribir_csv(nombre: str, encabezado: list[str], filas: list[list]) -> int:
    """Escribe en UTF-8 con BOM (Power BI y Excel lo leen bien) y reemplaza el archivo de una vez."""
    CARPETA.mkdir(parents=True, exist_ok=True)
    destino = CARPETA / nombre
    temporal = destino.with_suffix('.tmp')
    with open(temporal, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(encabezado)
        w.writerows(filas)
    temporal.replace(destino)
    return len(filas)


def separar_unidad(titulo: str) -> tuple[str, str]:
    """'Haber medio previsional SIPA (en pesos)' -> ('Haber medio previsional SIPA', 'pesos')."""
    m = re.search(r'\((?:en )?([^()]*)\)\s*$', titulo)
    if not m:
        return titulo.strip(), ''
    return titulo[:m.start()].strip(), m.group(1).strip()


# ----------------------------------------------------------------- BESS y ANSES del Sheet
def tabla_sheet() -> list[list]:
    from sheet_directo import ClienteSheets, periodo_de
    cliente = ClienteSheets()
    valores = cliente.leer('A1:BZZ2000', render='UNFORMATTED_VALUE')
    encabezados = valores[0] if valores else []
    columnas = []
    for i, t in enumerate(encabezados):
        texto = ' '.join(str(t or '').split())
        m = re.match(r'^(BESS|ANSES) (.+?)\s+([FN])$', texto)
        if m:
            columnas.append((i, m.group(1), m.group(2), 'Formosa' if m.group(3) == 'F' else 'Nación'))
    filas = []
    for fila in valores[1:]:
        if not fila:
            continue
        p = periodo_de(fila[0])
        if not p:
            continue
        for i, fuente, indicador, territorio in columnas:
            v = fila[i] if i < len(fila) else None
            if isinstance(v, bool) or not isinstance(v, (int, float)) or v == 0:
                continue  # en la base, vacío o 0 = sin dato
            nombre, unidad = separar_unidad(indicador)
            iso = f'{p[0]}-{p[1]:02d}-01'
            filas.append([iso, fuente, nombre, territorio, v, unidad,
                          'trimestral' if fuente == 'BESS' else 'mensual'])
    return filas


# ------------------------------------------------------- BESS por departamento (cuadro 6.1/6.2)
def leer_xlsx(contenido: bytes) -> dict:
    import openpyxl
    libro = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    hojas = {h.title: [list(r) for r in h.iter_rows(values_only=True)] for h in libro.worksheets}
    libro.close()
    return hojas


def periodo_cuadro(titulo: str) -> str | None:
    m = re.search(r'([A-ZÁÉÍÓÚ]+)\s+(20\d{2})\s*$', norm(titulo))
    if m and m.group(1) in MESES:
        return f'{m.group(2)}-{MESES[m.group(1)]:02d}-01'
    return None


def bloque_formosa(filas: list[list]) -> list[list]:
    inicio = next((i for i, f in enumerate(filas) if len(f) > 1 and norm(f[1]) == 'TOTAL FORMOSA'), None)
    if inicio is None:
        raise RuntimeError('el cuadro no trae el bloque de Formosa')
    salida = []
    for f in filas[inicio + 1:]:
        if len(f) < 3 or (f[0] not in (None, '') and norm(f[0]) != 'FORMOSA'):
            break  # empieza otra provincia
        if f[1] in (None, ''):
            break
        salida.append(f)
    return salida


def tabla_departamentos() -> tuple[list[list], str]:
    anio = date.today().year
    contenido, url = None, ''
    for a in (anio, anio - 1):
        url = URL_DEPTOS.format(anio=a)
        try:
            pedido = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Observatorio Formosa)'})
            with urllib.request.urlopen(pedido, timeout=120) as r:
                contenido = r.read()
            if contenido.startswith(b'PK'):
                break
        except Exception:  # noqa: BLE001
            contenido = None
    if not contenido:
        raise RuntimeError('no se pudo descargar el boletín por departamento del BESS')
    hojas = leer_xlsx(contenido)
    filas = []
    # Cuadro 6.1: Total, Jubilaciones (total, con y sin moratoria), Pensiones (total, con y sin
    # moratoria), PNC y PUAM.
    c61 = hojas.get('6.1')
    periodo = periodo_cuadro(str(c61[0][0])) if c61 else None
    tipos = [(2, 'Total beneficios', 'Total'), (3, 'Jubilaciones', 'Total'), (4, 'Jubilaciones', 'Con moratoria'),
             (5, 'Jubilaciones', 'Sin moratoria'), (6, 'Pensiones', 'Total'), (7, 'Pensiones', 'Con moratoria'),
             (8, 'Pensiones', 'Sin moratoria'), (9, 'Pensiones no contributivas', 'Total'), (10, 'PUAM', 'Total')]
    for f in bloque_formosa(c61):
        nombre = str(f[1]).strip()
        codigo = CODIGOS.get(norm(nombre), '')
        if not codigo:
            continue  # "No informado": beneficios sin departamento; se dejan fuera del tablero
        for col, tipo, detalle in tipos:
            v = f[col] if col < len(f) else None
            if isinstance(v, (int, float)):
                filas.append([periodo, codigo, nombre, tipo, detalle, v])
    c62 = hojas.get('6.2')
    if c62:
        p62 = periodo_cuadro(str(c62[0][0])) or periodo
        for f in bloque_formosa(c62):
            nombre = str(f[1]).strip()
            v = f[2] if len(f) > 2 else None
            if isinstance(v, (int, float)) and CODIGOS.get(norm(nombre)):
                filas.append([p62, CODIGOS[norm(nombre)], nombre, 'Prestación por desempleo', 'Total', v])
    return filas, url


def main() -> int:
    resumen = {'inicio': datetime.now().isoformat(timespec='seconds')}
    errores = []
    try:
        filas = tabla_sheet()
        resumen['seguridad_social'] = escribir_csv(
            'seguridad_social.csv', ['periodo', 'fuente', 'indicador', 'territorio', 'valor', 'unidad', 'frecuencia'], filas)
    except Exception as error:  # noqa: BLE001
        errores.append(f'Sheet: {error}')
    try:
        filas, url = tabla_departamentos()
        resumen['bess_departamentos'] = escribir_csv(
            'bess_departamentos.csv', ['periodo', 'codigo', 'departamento', 'tipo_beneficio', 'detalle', 'beneficios'], filas)
        resumen['url_departamentos'] = url
    except Exception as error:  # noqa: BLE001
        errores.append(f'Departamentos: {error}')
    resumen['poblacion'] = escribir_csv(
        'poblacion_departamentos.csv', ['codigo', 'departamento', 'poblacion_2022', 'fuente'],
        [[c, n, p, FUENTE_POBLACION] for c, n, p in POBLACION_2022])
    mapa = CARPETA / 'formosa_departamentos.topojson'
    origen = Path(__file__).with_name('formosa_departamentos.topojson')
    if origen.is_file() and not mapa.is_file():
        mapa.write_bytes(origen.read_bytes())
    resumen['errores'] = errores
    print('POWERBI=' + json.dumps(resumen, ensure_ascii=False))
    return 1 if errores else 0


if __name__ == '__main__':
    raise SystemExit(main())
