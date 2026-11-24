"""BESS - Boletín Estadístico de la Seguridad Social (Formosa y Total del país).

Fuente oficial: Secretaría de Trabajo, Empleo y Seguridad Social, sobre datos de
ANSES y ARCA. https://www.argentina.gob.ar/trabajo/seguridadsocial/bess

Cómo se publican los datos
* El BESS es TRIMESTRAL: cada edición (datos a marzo, junio, septiembre y
  diciembre) trae los cuadros por provincia de ese mes. Por eso el indicador se
  carga en las filas de marzo, junio, septiembre y diciembre de la base.
* El "BESS por jurisdicción" (BESSJ) es anual (datos a mayo) y sus series por
  provincia son promedios anuales: no se usa para la carga trimestral.

Cuadros que se leen de cada edición (se ubican por título, no por número):
* Pasivos 2.3.10  Beneficios en vigor y haber medio del SIPA por provincia.
* Pasivos 2.3.11  Beneficios en vigor del SIPA según tipo (jubilaciones / pensiones).
* PNC     3.3     Pensiones no contributivas según tipo de beneficio por provincia.
* PNC     3.9     Pensión Universal para el Adulto Mayor según jurisdicción.
* Desempleo 5.4   Beneficiarios de la prestación por desempleo según jurisdicción.
* AAFF    6.7     Asignaciones familiares: beneficios, masa y prestación promedio.

Uso
    python ejecutar_bess.py              # última edición publicada (revisión diaria)
    python ejecutar_bess.py --historico  # todas las ediciones desde marzo 2023
    python ejecutar_bess.py --solo-extraer  # no escribe en el Sheet

Escribe directo en la hoja BASE DE DATOS del Google Sheet por encabezado y
período (scripts/comun/sheet_directo.py). Si faltan las columnas del BESS, las
agrega al final de la hoja con sus encabezados.
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
import unicodedata
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))

PAGINA = 'https://www.argentina.gob.ar/trabajo/seguridadsocial/bess'
BASE_URL = 'https://www.argentina.gob.ar/sites/default/files/'
CARPETA_REPORTES = RAIZ / 'scripts' / 'reportes' / 'bess'
FUENTE = 'bess'
DESDE = (2023, 3)

# Ediciones desde marzo 2023 (los nombres de archivo no siguen un patrón fijo).
HISTORICO = [
    'estadisticas_de_seguridad_social_pasivos_03-2023.xlsx', 'estadisticas_de_seguridad_social_pnc_03-2023.xlsx',
    'estadisticas_de_seguridad_social_desempleo_03-2023.xlsx',
    'estadisticas_de_seguridad_social_pasivos_06-2023.xlsx', 'estadisticas_de_seguridad_social_pnc_06-2023_actualizado150923.xlsx',
    'estadisticas_de_seguridad_social_desempleo_06-2023.xlsx', 'mteyss-sss-estadisticas_de_seguridad_social_aaff_06-2023.xlsx',
    'mteyss-sss-bess-estadisticas_de_seguridad_social_pasivos_09-2023.xlsx', 'mteyss-sss-bess-estadisticas_de_seguridad_social_pnc_09-2023.xlsx',
    'mteyss-sss-prestaciones-por-desempleo_09-2023.xlsx', 'estadisticas_de_seguridad_social_aaff_09-2023.xlsx',
    'estadisticas_de_seguridad_social_pasivos_12-2023.xlsx', 'estadisticas_de_seguridad_social_pnc_12-2023.xlsx',
    'estadisticas_de_seguridad_social_desempleo_12-2023.xlsx', 'estadisticas_de_seguridad_social_aaff_12-2023.xlsx',
    'estadisticas_de_seguridad_social_pasivos_03-2024.xlsx', 'estadisticas_de_seguridad_social_pnc_03-2024.xlsx',
    'estadisticas_de_seguridad_social_desempleo_03-2024.xlsx', 'estadisticas_de_seguridad_social_aaff_03-2024.xlsx',
    'estadisticas_de_seguridad_social_pasivos_06-2024.xlsx', 'estadisticas_de_seguridad_social_pnc_06-2024.xlsx',
    'estadisticas_de_seguridad_social_desempleo_06-2024.xlsx', 'estadisticas_de_seguridad_social_aaff_06-2024.xlsx',
    'estadisticas_de_seguridad_social_pasivos_09-2024.xlsx', 'estadisticas_de_seguridad_social_pnc_09-2024.xlsx',
    'estadisticas_de_seguridad_social_desempleo_09-2024.xlsx', 'estadisticas_de_seguridad_social_aaff_09-2024.xlsx',
    'estadisticas_de_seguridad_social_pasivos_12-2024.xlsx', 'estadisticas_de_seguridad_social_pnc_12-2024.xlsx',
    'estadisticas_de_seguridad_social_desempleo_12-2024.xlsx', 'estadisticas_de_seguridad_social_aaff_12-2024.xlsx',
    'estadisticas_de_seguridad_social_pasivos_03-2025.xlsx', 'estadisticas_de_seguridad_social_pnc_03-2025.xlsx',
    'estadisticas_de_seguridad_social_desempleo_03-2025.xlsx', 'estadisticas_de_seguridad_social_aaff_03-2025.xlsx',
    'estadisticas_de_seguridad_social_pasivos_06-2025.xlsx', 'estadisticas_de_seguridad_social_pnc_06-2025.xlsx',
    'estadisticas_de_seguridad_social_desempleo_06-2025.xlsx', 'estadisticas_de_seguridad_social_aaff_06-2025.xlsx',
    'estadisticas_de_seguridad_social_pasivos_09-2025.xlsx', 'estadisticas_de_seguridad_social_pnc_09-2025.xlsx',
    'estadisticas_de_seguridad_social_desempleo_09-2025.xlsx', 'estadisticas_de_seguridad_social_aaff_09-2025.xlsx',
    'estadisticas_de_seguridad_social_pasivos_12-2025.xlsx', 'estadisticas_de_seguridad_social_pnc_12-2025.xlsx',
    'estadisticas_de_seguridad_social_desempleo_12-2025.xlsx', 'estadisticas_de_seguridad_social_aaff_12-2025.xlsx',
    'estadisticas_de_seguridad_social_pasivos_03-2026.xlsx', 'estadisticas_de_seguridad_social_pnc_03-2026.xlsx',
    'estadisticas_de_seguridad_social_desempleo_03-2026.xlsx', 'estadisticas_de_seguridad_social_aaff_03-2026.xlsx',
]

# (clave interna, encabezado en el Sheet sin el sufijo F/N, decimales)
VARIABLES = [
    ('sipa_beneficios', 'BESS Beneficios previsionales SIPA (cantidad)', 0),
    ('sipa_haber', 'BESS Haber medio previsional SIPA (en pesos)', 0),
    ('jubilaciones', 'BESS Jubilaciones SIPA (cantidad)', 0),
    ('pensiones', 'BESS Pensiones SIPA (cantidad)', 0),
    ('pnc_total', 'BESS Pensiones no contributivas - Total (cantidad)', 0),
    ('pnc_invalidez', 'BESS Pensiones no contributivas - Invalidez (cantidad)', 0),
    ('pnc_vejez', 'BESS Pensiones no contributivas - Vejez (cantidad)', 0),
    ('pnc_madre7', 'BESS Pensiones no contributivas - Madre de 7 o más hijos (cantidad)', 0),
    ('pnc_leyes_especiales', 'BESS Pensiones no contributivas - Leyes especiales (cantidad)', 0),
    ('pnc_graciables', 'BESS Pensiones no contributivas - Graciables (cantidad)', 0),
    ('pnc_ex_combatientes', 'BESS Pensiones no contributivas - Ex combatientes Malvinas (cantidad)', 0),
    ('puam_casos', 'BESS PUAM - Beneficios (cantidad)', 0),
    ('puam_promedio', 'BESS PUAM - Beneficio promedio (en pesos)', 2),
    ('desempleo', 'BESS Prestación por desempleo - Beneficiarios (cantidad)', 0),
    ('aaff_beneficios', 'BESS Asignaciones familiares - Beneficios total (cantidad)', 0),
    ('aaff_hijo', 'BESS Asignaciones familiares - Por hijo (cantidad)', 0),
    ('aaff_hijo_discapacidad', 'BESS Asignaciones familiares - Por hijo con discapacidad (cantidad)', 0),
    ('aaff_masa', 'BESS Asignaciones familiares - Masa (en millones de pesos)', 2),
    ('aaff_promedio', 'BESS Asignaciones familiares - Prestación promedio (en pesos)', 2),
]
BLOQUES = [('sipa_beneficios', 'pensiones'), ('pnc_total', 'puam_promedio'), ('desempleo', 'desempleo'),
           ('aaff_beneficios', 'aaff_promedio')]


def encabezados_en_orden() -> list[str]:
    """Orden de las columnas: por bloque, primero Formosa (F) y después Total del país (N)."""
    claves = [v[0] for v in VARIABLES]
    salida = []
    for desde, hasta in BLOQUES:
        grupo = VARIABLES[claves.index(desde):claves.index(hasta) + 1]
        for suf in ('F', 'N'):
            salida += [f'{titulo} {suf}' for _, titulo, _ in grupo]
    return salida


TITULO = {k: t for k, t, _ in VARIABLES}
DECIMALES = {k: d for k, _, d in VARIABLES}
MESES = {'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6, 'JULIO': 7, 'AGOSTO': 8,
         'SEPTIEMBRE': 9, 'SETIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}
TOTAL = re.compile(r'^TOTAL( DEL PAIS)?$')
FORMOSA = re.compile(r'^FORMOSA$')


def norm(x) -> str:
    s = unicodedata.normalize('NFD', str(x if x is not None else ''))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return ' '.join(s.split()).upper()


def descargar(nombre_o_url: str) -> bytes:
    url = nombre_o_url if nombre_o_url.startswith('http') else BASE_URL + nombre_o_url
    pedido = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Observatorio Formosa)'})
    with urllib.request.urlopen(pedido, timeout=120) as r:
        return r.read()


def ultimas_ediciones() -> list[str]:
    """Links de la última edición de cada cuadernillo ('Últimas publicaciones del BESS')."""
    html = descargar(PAGINA).decode('utf-8', errors='replace')
    patron = re.compile(r'/sites/default/files/([^"\'\s]*estadisticas_de_seguridad_social_'
                        r'(pasivos|pnc|desempleo|aaff)_(\d{2})-(\d{4})[^"\'\s]*\.xlsx)')
    mejores = {}
    for m in patron.finditer(html):
        tipo, clave = m.group(2), (int(m.group(4)), int(m.group(3)))
        if tipo not in mejores or clave > mejores[tipo][0]:
            mejores[tipo] = (clave, m.group(1))
    if len(mejores) < 4:
        raise RuntimeError(f'La página del BESS no muestra los 4 cuadernillos (encontrados: {sorted(mejores)})')
    ultimas_ediciones.periodos = {t: v[0] for t, v in mejores.items()}
    return [v[1] for v in mejores.values()]


NOMBRES_CUADERNILLO = {'pasivos': 'Jubilaciones y pensiones', 'pnc': 'Pensiones no contributivas y PUAM',
                       'desempleo': 'Desempleo', 'aaff': 'Asignaciones familiares'}


def aviso_cuadernillos_atrasados() -> str | None:
    """Los cuadernillos no siempre salen juntos: avisa cuáles todavía no tienen el último trimestre."""
    periodos = getattr(ultimas_ediciones, 'periodos', {})
    if not periodos:
        return None
    ultimo = max(periodos.values())
    atrasados = sorted(t for t, p in periodos.items() if p < ultimo)
    if not atrasados:
        return None
    meses = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
    partes = [f"{NOMBRES_CUADERNILLO[t]} (último publicado: {meses[periodos[t][1] - 1]} {periodos[t][0] % 100:02d})" for t in atrasados]
    return (f"El BESS de {meses[ultimo[1] - 1]} {ultimo[0] % 100:02d} todavía no publicó: " + '; '.join(partes)
            + '. Se completa solo cuando salga.')


# ------------------------------------------------------------- lectura Excel
def hojas(contenido: bytes):
    from openpyxl import load_workbook
    libro = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    for ws in libro.worksheets:
        filas = [list(f) for f in ws.iter_rows(values_only=True)]
        filas = [f for f in filas if any(v not in (None, '') for v in f)]
        titulo = ' | '.join(norm(v) for f in filas[:4] for v in f if v not in (None, ''))
        hojas.ultima_pestana = norm(ws.title)
        yield titulo, filas


def buscar(contenido: bytes, patron: str):
    rx = re.compile(patron)
    for titulo, filas in hojas(contenido):
        if rx.search(titulo):
            return titulo, filas
    return None, None


def fila(filas, rx):
    for f in filas:
        if any(rx.match(norm(x)) for x in f[:3]) and any(isinstance(x, (int, float)) for x in f):
            return f
    raise ValueError(f'no se encontró la fila {rx.pattern}')


def numeros(f):
    return [x for x in f if isinstance(x, (int, float)) and not isinstance(x, bool)]


def periodo_titulo(titulo: str, ancla: str) -> tuple[int, int]:
    m = re.search(ancla + r'\. ([A-Z]+) (\d{4})', titulo)
    if not m or m.group(1) not in MESES:
        raise ValueError(f'no se pudo leer el mes del título: {titulo[:140]}')
    return int(m.group(2)), MESES[m.group(1)]


def a_periodo(x):
    if isinstance(x, datetime):
        return (x.year, x.month)
    if isinstance(x, date):
        return (x.year, x.month)
    if isinstance(x, (int, float)) and 40000 < x < 60000:
        d = date(1899, 12, 30) + timedelta(days=int(x))
        return (d.year, d.month)
    return None


def extraer(nombre: str, contenido: bytes) -> dict:
    """{(año, mes): {clave_F/N: valor}} de un cuadernillo."""
    datos: dict = {}

    def poner(p, clave, valor):
        datos.setdefault(p, {})[clave] = valor

    if 'pasivos' in nombre:
        titulo, filas = buscar(contenido, r'BENEFICIOS EN VIGOR Y HABER MEDIO.*POR PROVINCIA')
        if not filas:
            raise ValueError('falta el cuadro 2.3.10 (beneficios y haber medio por provincia)')
        cab = next(f for f in filas if sum(1 for x in f if a_periodo(x)) > 3)
        c = max(i for i, x in enumerate(cab) if a_periodo(x))
        p = a_periodo(cab[c])
        for rx, suf in ((TOTAL, 'N'), (FORMOSA, 'F')):
            f = fila(filas, rx)
            poner(p, f'sipa_beneficios_{suf}', f[c])
            poner(p, f'sipa_haber_{suf}', f[c + 1])
        titulo, filas = buscar(contenido, r'BENEFICIOS EN VIGOR DEL SIPA SEGUN TIPO DE BENEFICIO, SEXO Y JURISDICCION')
        if not filas:
            raise ValueError('falta el cuadro 2.3.11 (jubilaciones y pensiones por jurisdicción)')
        p = periodo_titulo(titulo, 'JURISDICCION')
        cab_j = next(f for f in filas if any(norm(x) == 'JURISDICCION' for x in f))
        limite = next((i for i, x in enumerate(cab_j) if i > 0 and norm(x) == 'JURISDICCION'), len(cab_j))
        tipos = next(f for f in filas if any(norm(x) == 'JUBILACIONES' for x in f))
        for rx, suf in ((TOTAL, 'N'), (FORMOSA, 'F')):
            f = fila(filas, rx)
            jub = sum(f[i] or 0 for i, x in enumerate(tipos[:limite]) if norm(x).startswith('JUBILACIONES') and isinstance(f[i], (int, float)))
            pen = sum(f[i] or 0 for i, x in enumerate(tipos[:limite]) if norm(x).startswith('PENSIONES') and isinstance(f[i], (int, float)))
            total = numeros(f)[0]
            if abs(total - jub - pen) > 0.5:
                raise ValueError(f'2.3.11 {suf}: jubilaciones + pensiones ({jub}+{pen}) no suman el total {total}')
            poner(p, f'jubilaciones_{suf}', jub)
            poner(p, f'pensiones_{suf}', pen)
    elif 'pnc' in nombre:
        titulo, filas = buscar(contenido, r'PENSIONES NO CONTRIBUTIVAS SEGUN TIPO DE BENEFICIO POR PROVINCIA')
        if not filas:
            raise ValueError('falta el cuadro 3.3 (PNC por tipo y provincia)')
        p = periodo_titulo(titulo, 'PROVINCIA')
        claves = ['pnc_total', 'pnc_asistenciales', 'pnc_invalidez', 'pnc_vejez', 'pnc_madre7',
                  'pnc_leyes_especiales', 'pnc_graciables', 'pnc_ex_combatientes']
        for rx, suf in ((TOTAL, 'N'), (FORMOSA, 'F')):
            v = numeros(fila(filas, rx))
            if abs(v[0] - v[1] - v[5] - v[6] - v[7]) > 0.5 or abs(v[1] - v[2] - v[3] - v[4]) > 0.5:
                raise ValueError(f'3.3 {suf}: los tipos de PNC no suman el total ({v[:8]})')
            for k, x in zip(claves, v):
                if k != 'pnc_asistenciales':
                    poner(p, f'{k}_{suf}', x)
        titulo, filas = buscar(contenido, r'PENSION UNIVERSAL PARA EL ADULTO MAYOR SEGUN JURISDICCION\. [A-Z]+ \d{4}')
        if not filas:
            raise ValueError('falta el cuadro 3.9 (PUAM por jurisdicción)')
        p = periodo_titulo(titulo, 'JURISDICCION')
        for rx, suf in ((TOTAL, 'N'), (FORMOSA, 'F')):
            v = numeros(fila(filas, rx))
            poner(p, f'puam_casos_{suf}', v[0])
            poner(p, f'puam_promedio_{suf}', v[2])
    elif 'desempleo' in nombre:
        titulo, filas = buscar(contenido, r'PRESTACION POR DESEMPLEO SEGUN SEXO Y JURISDICCION')
        if not filas:
            raise ValueError('falta el cuadro 5.4 (desempleo por jurisdicción)')
        p = periodo_titulo(titulo, 'JURISDICCION')
        for rx, suf in ((TOTAL, 'N'), (FORMOSA, 'F')):
            poner(p, f'desempleo_{suf}', numeros(fila(filas, rx))[0])
    elif 'aaff' in nombre:
        titulo, filas = buscar(contenido, r'ASIGNACIONES FAMILIARES.*SEGUN TIPO DE BENEFICIO Y JURISDICCION')
        if not filas:
            # Ediciones nuevas pueden cambiar la redacción del título: se busca la pestaña 6.7
            # o un cuadro de asignaciones familiares por jurisdicción/provincia con fila FORMOSA.
            for t, f in hojas(contenido):
                pest = getattr(hojas, 'ultima_pestana', '')
                if (re.match(r'^(CUADRO )?6\.7\b', pest) or re.search(r'\b6\.7\b', t[:40])
                        or re.search(r'ASIGNACIONES FAMILIARES.*(JURISDICCION|PROVINCIA)', t)):
                    if any(FORMOSA.match(norm(x)) for r in f for x in r[:3]):
                        titulo, filas = t, f
                        break
        if not filas:
            raise ValueError('falta el cuadro 6.7 (asignaciones familiares por jurisdicción)')
        try:
            p = periodo_titulo(titulo, 'JURISDICCION')
        except ValueError:
            m = re.search(r'aaff_(\d{2})-(\d{4})', nombre)
            if not m:
                raise
            p = (int(m.group(2)), int(m.group(1)))
        for rx, suf in ((TOTAL, 'N'), (FORMOSA, 'F')):
            v = numeros(fila(filas, rx))
            if abs(v[0] - v[1] - v[2]) > 0.5:
                raise ValueError(f'6.7 {suf}: hijo + discapacidad no suman el total ({v[:3]})')
            poner(p, f'aaff_beneficios_{suf}', v[0])
            poner(p, f'aaff_hijo_{suf}', v[1])
            poner(p, f'aaff_hijo_discapacidad_{suf}', v[2])
            # El cuadro dice "en millones de pesos" pero publica pesos: se pasa a millones.
            poner(p, f'aaff_masa_{suf}', v[3] / 1_000_000)
            poner(p, f'aaff_promedio_{suf}', v[6])
    return datos


def a_cargas(datos: dict) -> list[dict]:
    cargas = []
    for p in sorted(datos):
        if p < DESDE:
            continue
        valores = {}
        for clave, valor in sorted(datos[p].items()):
            base, suf = clave.rsplit('_', 1)
            if base not in TITULO or valor is None:
                continue
            valores[f'{TITULO[base]} {suf}'] = round(float(valor), DECIMALES[base]) if DECIMALES[base] else int(round(valor))
        if valores:
            cargas.append({'fuente': FUENTE, 'hoja': 'BASE DE DATOS', 'periodo': f'{p[0]}-{p[1]:02d}-01', 'valores': valores})
    return cargas


# ------------------------------------------------------------- Sheet
def asegurar_columnas(cliente) -> list[str]:
    """Agrega al final de BASE DE DATOS las columnas del BESS que falten."""
    from sheet_directo import API, HOJA, Mapa, clave, letra
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    # Títulos anteriores (con "en vigor", sin "(cantidad)") o con espacios de más: se dejan con el nombre actual.
    nuevos = {clave(n): n for n in encabezados_en_orden()}
    viejos = {}
    for n in encabezados_en_orden():
        sin_cant = n.replace(' (cantidad)', '')
        viejos[clave(sin_cant)] = n
        if ' SIPA ' in sin_cant:
            viejos[clave(sin_cant.replace(' SIPA ', ' SIPA en vigor '))] = n
    renombrar = []
    for i, titulo in enumerate(encabezados, start=1):
        if titulo in (None, ''):
            continue
        destino = nuevos.get(clave(titulo)) or viejos.get(clave(titulo))
        if destino and titulo != destino:
            renombrar.append({'range': f"'{HOJA}'!{letra(i)}1", 'values': [[destino]]})
            encabezados[i - 1] = destino
    if renombrar:
        cliente.sesion.post(f'{API}/values:batchUpdate', json={'valueInputOption': 'RAW', 'data': renombrar},
                            timeout=60).raise_for_status()
    mapa = Mapa(encabezados, [])
    faltan = [t for t in encabezados_en_orden() if mapa.columna(t)[1] == 0]
    if not faltan:
        return []
    ultima = max((i for i, t in enumerate(encabezados, start=1) if t not in (None, '')), default=1)
    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(sheetId,title,gridProperties))'}, timeout=60)
    r.raise_for_status()
    props = next(s['properties'] for s in r.json()['sheets'] if s['properties']['title'] == HOJA)
    necesarias = ultima + len(faltan) - props['gridProperties']['columnCount']
    if necesarias > 0:
        r = cliente.sesion.post(f'{API}:batchUpdate', json={'requests': [{'appendDimension': {
            'sheetId': props['sheetId'], 'dimension': 'COLUMNS', 'length': necesarias}}]}, timeout=60)
        r.raise_for_status()
    desde = letra(ultima + 1)
    hasta = letra(ultima + len(faltan))
    cliente.sesion.post(f'{API}/values:batchUpdate', json={'valueInputOption': 'RAW', 'data': [
        {'range': f"'{HOJA}'!{desde}1:{hasta}1", 'values': [faltan]}]}, timeout=60).raise_for_status()
    return faltan


def main() -> int:
    ap = argparse.ArgumentParser(description='BESS - Seguridad social (Formosa y Total del país), trimestral')
    ap.add_argument('--historico', action='store_true', help='Carga todas las ediciones desde marzo 2023.')
    ap.add_argument('--solo-extraer', action='store_true', help='No escribe en el Sheet.')
    args = ap.parse_args()

    archivos = ultimas_ediciones()
    aviso = aviso_cuadernillos_atrasados()
    if aviso:
        print('NOTA_TELEGRAM=' + aviso)
    if args.historico:
        archivos = HISTORICO + [a for a in archivos if a not in HISTORICO]
    datos: dict = {}
    fuentes_usadas = []
    fallidos = []
    for nombre in archivos:
        contenido = descargar(nombre)
        try:
            parcial = extraer(nombre, contenido)
        except Exception as error:  # noqa: BLE001
            if args.historico:
                raise
            copia = RAIZ / 'fuentes' / 'bess' / Path(nombre).name
            try:
                copia.parent.mkdir(parents=True, exist_ok=True)
                copia.write_bytes(contenido)
            except OSError:
                pass
            fallidos.append(f'{Path(nombre).name}: {error} (copia guardada en fuentes\\bess)')
            continue
        for p, valores in parcial.items():
            datos.setdefault(p, {}).update(valores)
        fuentes_usadas.append(BASE_URL + nombre)
    cargas = a_cargas(datos)
    if not cargas:
        raise RuntimeError('No se obtuvo ningún dato del BESS.')
    periodos = [c['periodo'][:7] for c in cargas]
    resultado = {'fuente': 'BESS - Seguridad social (ANSES)', 'periodo': periodos[-1], 'periodos': periodos,
                 'frecuencia': 'trimestral (marzo, junio, septiembre, diciembre)', 'archivos': fuentes_usadas,
                 'datos': {c['periodo'][:7]: c['valores'] for c in cargas}}
    CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
    (CARPETA_REPORTES / f'bess_{periodos[-1]}.json').write_text(json.dumps(resultado, ensure_ascii=False, indent=1), encoding='utf-8')

    if not args.solo_extraer:
        from sheet_directo import ClienteSheets, escribir_actualizaciones
        cliente = ClienteSheets()
        agregadas = asegurar_columnas(cliente)
        escritura = escribir_actualizaciones(cargas, None, cliente)
        info = escritura.get('por_fuente', {}).get(FUENTE, {})
        resultado['columnas_agregadas'] = agregadas
        resultado['actualizacion'] = {'celdas_actualizadas': info.get('escritas', 0),
                                      'cambios': info.get('cambios', [])[:60], 'errores': info.get('errores', [])}
        resultado.pop('datos', None)
        print(json.dumps(resultado, ensure_ascii=False, indent=1, default=str))
        resumen = {k: info.get(k) for k in ('periodos', 'escritas', 'ya_estaban', 'esperadas', 'conservadas',
                                            'no_encontradas', 'errores')}
        print('SHEET_DIRECTO=' + json.dumps(resumen, ensure_ascii=False, default=str))
        if info.get('errores'):
            print('ERROR=' + '; '.join(info['errores'])[:500])
            print('ESTADO=ERROR')
            return 1
    else:
        print(json.dumps(resultado, ensure_ascii=False, indent=1, default=str))
    if fallidos:
        print('ERROR=' + '; '.join(fallidos)[:500])
        print('ESTADO=ERROR')
        return 1
    print('ESTADO=OK')
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as error:  # noqa: BLE001
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        raise SystemExit(1)
