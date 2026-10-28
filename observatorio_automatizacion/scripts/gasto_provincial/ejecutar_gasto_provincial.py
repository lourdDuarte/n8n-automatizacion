"""Gasto provincial por finalidad y función - Formosa (F) y total de provincias (N).

Fuente oficial: Ministerio de Economía de la Nación, Secretaría de Hacienda, Subsecretaría de
Coordinación Fiscal Provincial (ex DNAP), "Gastos por finalidad y función" de la Administración
Pública No Financiera de cada provincia.
    https://www.argentina.gob.ar/economia/sechacienda/coordinacion-fiscal-provincial/ejecucion-presupuestaria-provincial/gastos-por
Cada trimestre se publica un Excel (gfyf_apnf_AAAA_-_trim._X.xls/xlsx) con el gasto ACUMULADO
del año hasta ese trimestre, en millones de pesos corrientes, por provincia.

Frecuencia: TRIMESTRAL. Se carga en la fila del último mes del trimestre (marzo, junio,
septiembre, diciembre), igual que el BESS. Diciembre = año completo.
    * Montos: acumulado del año, en millones de pesos.
    * Participación (%): parte de cada función en el gasto total de la provincia (la misma
      "Estructura %" que publica el Ministerio en la segunda hoja del Excel).
    * N = total de provincias que publica el Ministerio (en 2026 excluye a La Pampa, que no
      informó; lo aclara la Nota 3 del archivo).
El Ministerio reemplaza los archivos trimestrales cuando publica el cierre del año: de algunos
años solo queda el IV trimestre. Se carga lo que esté publicado.

Uso
    python ejecutar_gasto_provincial.py               # año en curso y anterior (revisión diaria)
    python ejecutar_gasto_provincial.py --historico   # desde 2023
    python ejecutar_gasto_provincial.py --solo-extraer
"""
from __future__ import annotations

import argparse
import html
import json
import re
import sys
import time
import unicodedata
import urllib.request
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))

FUENTE = 'gasto_provincial'
PAGINA = ('https://www.argentina.gob.ar/economia/sechacienda/coordinacion-fiscal-provincial/'
          'ejecucion-presupuestaria-provincial/gastos-por')
CARPETA_FUENTES = RAIZ / 'fuentes' / 'gasto_provincial'
CARPETA_REPORTES = RAIZ / 'scripts' / 'reportes' / 'gasto_provincial'
DESDE = 2023
PREFIJO = 'Gasto provincial MECON - '
TRIM = {'i': 3, 'ii': 6, 'iii': 9, 'iv': 12}
PATRON = re.compile(r'gfyf_apnf_(\d{4})_-_trim\.?_(iv|iii|ii|i)(?:_\d+)?\.xlsx?$', re.I)
# Respaldo para el GASTO TOTAL cuando falta el trimestre de finalidad y función (pasa con I-III de
# 2025, que el Ministerio retiró al publicar el cierre del año): el mismo Ministerio publica
# "Ejecuciones presupuestarias" (esquema Ahorro-Inversión-Financiamiento) y su fila
# "Gastos Totales" es el mismo número que el TOTAL de finalidad y función (verificado en
# jun-24, dic-25 y jun-26). Las funciones (seguridad social, salud, etc.) no están en ese
# esquema, así que en esos trimestres solo se carga el gasto total.
PAGINA_AIF = ('https://www.argentina.gob.ar/economia/sechacienda/coordinacion-fiscal-provincial/'
              'ejecucion-presupuestaria-provincial/ejecuciones')
PATRON_AIF = re.compile(r'ejecuciones_presupuestarias_apnf_(\d{4})_-_trim\.?_(iv|iii|ii|i)(?:_\d+)?\.xlsx?$', re.I)

# (clave, rótulo del Excel normalizado, título de la columna)
MONTOS = [
    ('total', 'TOTAL', 'Gasto total, acumulado del año (en millones de pesos)'),
    ('sociales', 'SERVICIOS SOCIALES', 'Servicios sociales, acumulado del año (en millones de pesos)'),
    ('seguridad', 'SEGURIDAD SOCIAL', 'Seguridad social (jubilaciones y pensiones provinciales), acumulado del año (en millones de pesos)'),
    ('promocion', 'PROMOCION Y ASISTENCIA SOCIAL', 'Promoción y asistencia social, acumulado del año (en millones de pesos)'),
    ('salud', 'SALUD', 'Salud, acumulado del año (en millones de pesos)'),
    ('educacion', 'EDUCACION Y CULTURA', 'Educación y cultura, acumulado del año (en millones de pesos)'),
]
PARTICIPACION = [
    ('sociales', 'Servicios sociales, participación en el gasto total (%)'),
    ('seguridad', 'Seguridad social, participación en el gasto total (%)'),
    ('promocion', 'Promoción y asistencia social, participación en el gasto total (%)'),
]


def encabezados_en_orden() -> list[str]:
    salida = []
    for suf in ('F', 'N'):
        salida += [f'{PREFIJO}{t} {suf}' for _, _, t in MONTOS]
        salida += [f'{PREFIJO}{t} {suf}' for _, t in PARTICIPACION]
    return salida


def norm(x) -> str:
    s = unicodedata.normalize('NFD', str(x or ''))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = re.sub(r'^\s*[IVX]+\.\s*', '', ' '.join(s.split()))   # "III. Servicios sociales" -> "Servicios sociales"
    return s.upper().strip()


def descargar(url: str, timeout: int = 120) -> bytes:
    pedido = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Observatorio Formosa)'})
    with urllib.request.urlopen(pedido, timeout=timeout) as r:
        return r.read()


def enlaces_publicados(pagina: str = PAGINA, patron=PATRON) -> dict[tuple[int, int], str]:
    """{(año, mes de cierre del trimestre): url} según la página del Ministerio."""
    texto = descargar(pagina, 90).decode('utf-8', errors='replace')
    enlaces = {}
    for href in re.findall(r'href="([^"]+)"', texto):
        url = html.unescape(href)
        m = patron.search(url.split('?')[0])
        if not m:
            continue
        if url.startswith('/'):
            url = 'https://www.argentina.gob.ar' + url
        # Algunos enlaces de la página apuntan al servidor interno "back.argentina.gob.ar",
        # que rechaza las descargas (403). El mismo archivo está en www.argentina.gob.ar.
        url = re.sub(r'^https?://back\.argentina\.gob\.ar', 'https://www.argentina.gob.ar', url)
        enlaces[(int(m.group(1)), TRIM[m.group(2).lower()])] = url
    if not enlaces:
        raise RuntimeError('La página del Ministerio de Economía no muestra archivos de gasto por finalidad y función')
    return enlaces


def archivo(url: str, anio: int) -> Path:
    """Guarda el Excel en fuentes/gasto_provincial. Los del año en curso y el anterior se renuevan
    cada 20 horas (el Ministerio puede corregirlos); los más viejos quedan guardados."""
    CARPETA_FUENTES.mkdir(parents=True, exist_ok=True)
    ruta = CARPETA_FUENTES / url.rsplit('/', 1)[-1].split('?')[0]
    vigente = anio >= date.today().year - 1
    if ruta.is_file() and (not vigente or time.time() - ruta.stat().st_mtime < 20 * 3600):
        return ruta
    try:
        contenido = descargar(url)
    except Exception:  # noqa: BLE001
        # Si falla, se prueba el mismo nombre de archivo en la carpeta pública del sitio.
        alternativa = 'https://www.argentina.gob.ar/sites/default/files/' + ruta.name
        if alternativa == url:
            raise
        contenido = descargar(alternativa)
    if not (contenido.startswith(b'PK') or contenido.startswith(b'\xd0\xcf\x11\xe0')):
        raise RuntimeError(f'{ruta.name}: la descarga no es un Excel')
    temporal = ruta.with_suffix(ruta.suffix + '.tmp')
    temporal.write_bytes(contenido)
    temporal.replace(ruta)
    return ruta


def filas_excel(ruta: Path) -> list[list]:
    """Filas de la primera hoja (la de millones de pesos)."""
    if ruta.suffix.lower() == '.xls':
        import xlrd
        hoja = xlrd.open_workbook(str(ruta)).sheet_by_index(0)
        return [hoja.row_values(i) for i in range(hoja.nrows)]
    import openpyxl
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hoja = libro.worksheets[0]
    filas = [list(f) for f in hoja.iter_rows(values_only=True)]
    libro.close()
    return filas


def numero(x) -> float | None:
    if x is None or x == '':
        return None
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def leer_cuadro(filas: list[list]) -> dict:
    """{'F': {clave: monto}, 'N': {clave: monto}} a partir de las filas del Excel."""
    fila_enc = next((i for i, f in enumerate(filas) if any(norm(c) == 'FORMOSA' for c in f)), None)
    if fila_enc is None:
        raise RuntimeError('no se encontró la columna Formosa')
    enc = [norm(c) for c in filas[fila_enc]]
    col_f = enc.index('FORMOSA')
    col_n = max(i for i, c in enumerate(enc) if c == 'TOTAL')
    rotulos = {r for _, r, _ in MONTOS}
    salida = {'F': {}, 'N': {}}
    for f in filas[fila_enc + 1:]:
        if not f:
            continue
        r = norm(f[0])
        if r not in rotulos:
            continue
        clave = next(k for k, rot, _ in MONTOS if rot == r)
        if clave in salida['F']:
            continue   # se toma la primera aparición
        vf, vn = numero(f[col_f]), numero(f[col_n])
        if vf is not None:
            salida['F'][clave] = vf
        if vn is not None:
            salida['N'][clave] = vn
    faltan = [k for k, _, _ in MONTOS if k not in salida['F']]
    if faltan:
        raise RuntimeError('faltan filas en el cuadro: ' + ', '.join(faltan))
    return salida


def leer_gasto_total_aif(filas: list[list]) -> dict:
    """Fila 'Gastos Totales' del esquema Ahorro-Inversión-Financiamiento: {'F': x, 'N': y o None}."""
    fila_enc = next((i for i, f in enumerate(filas) if any(norm(c) == 'FORMOSA' for c in f)), None)
    if fila_enc is None:
        raise RuntimeError('no se encontró la columna Formosa')
    enc = [norm(c) for c in filas[fila_enc]]
    col_f = enc.index('FORMOSA')
    col_n = max(i for i, c in enumerate(enc) if c == 'TOTAL')
    fila = next((f for f in filas[fila_enc + 1:] if f and 'GASTOS TOTALES' in norm(f[0])), None)
    if fila is None:
        raise RuntimeError('no se encontró la fila Gastos Totales')
    vf, vn = numero(fila[col_f]), numero(fila[col_n])
    if vf is None:
        raise RuntimeError('Formosa sin gasto total')
    # Si falta más de una jurisdicción (no solo La Pampa), el total de provincias no es comparable.
    sin_dato = [enc[j] for j in range(1, col_n) if enc[j] and not numero(fila[j])]
    return {'F': vf, 'N': vn if vn and len(sin_dato) <= 1 else None}


def valores_fila(cuadro: dict) -> dict:
    valores = {}
    for suf in ('F', 'N'):
        d = cuadro[suf]
        if suf == 'N' and any(k not in d for k, _, _ in MONTOS):
            # El Ministerio no publicó el total de provincias de todas las funciones (pasa en el
            # III trimestre 2024, que además excluye a Buenos Aires, La Pampa y Neuquén): un total
            # parcial no es comparable con los demás trimestres, así que el bloque N queda vacío.
            for _, _, titulo in MONTOS:
                valores[f'{PREFIJO}{titulo} N'] = ''
            for _, titulo in PARTICIPACION:
                valores[f'{PREFIJO}{titulo} N'] = ''
            continue
        for clave, _, titulo in MONTOS:
            if clave in d:
                valores[f'{PREFIJO}{titulo} {suf}'] = round(d[clave], 2)
        total = d.get('total')
        if total:
            for clave, titulo in PARTICIPACION:
                if clave in d:
                    valores[f'{PREFIJO}{titulo} {suf}'] = round(d[clave] / total * 100, 2)
    return valores


def main() -> int:
    ap = argparse.ArgumentParser(description='Gasto provincial por finalidad y función (Ministerio de Economía)')
    ap.add_argument('--historico', action='store_true', help=f'Carga desde {DESDE}.')
    ap.add_argument('--solo-extraer', action='store_true', help='No escribe en el Sheet.')
    args = ap.parse_args()

    hoy = date.today()
    marca_historico = CARPETA_REPORTES / 'historico_cargado.txt'
    if not args.historico and not args.solo_extraer and not marca_historico.is_file():
        args.historico = True
    desde = DESDE if args.historico else hoy.year - 1

    enlaces = enlaces_publicados()
    cargas, fallos = [], []
    for (anio, mes) in sorted(enlaces):
        if anio < desde:
            continue
        url = enlaces[(anio, mes)]
        try:
            cuadro = leer_cuadro(filas_excel(archivo(url, anio)))
        except Exception as error:  # noqa: BLE001
            fallos.append(f'{anio}-{mes:02d}: {error}')
            continue
        cargas.append({'fuente': FUENTE, 'hoja': 'BASE DE DATOS', 'periodo': f'{anio}-{mes:02d}-01',
                       'valores': valores_fila(cuadro), 'permitir_reemplazos': True})

    # Trimestres sin archivo de finalidad y función: gasto total desde "Ejecuciones presupuestarias".
    solo_total = []
    try:
        enlaces_aif = enlaces_publicados(PAGINA_AIF, PATRON_AIF)
    except Exception as error:  # noqa: BLE001
        enlaces_aif = {}
        fallos.append(f'ejecuciones presupuestarias: {error}')
    for (anio, mes) in sorted(enlaces_aif):
        if anio < desde or (anio, mes) in enlaces:
            continue
        try:
            total = leer_gasto_total_aif(filas_excel(archivo(enlaces_aif[(anio, mes)], anio)))
        except Exception as error:  # noqa: BLE001
            fallos.append(f'{anio}-{mes:02d} (gasto total): {error}')
            continue
        titulo = f'{PREFIJO}{MONTOS[0][2]}'
        valores = {f'{titulo} F': round(total['F'], 2), f'{titulo} N': round(total['N'], 2) if total['N'] else ''}
        cargas.append({'fuente': FUENTE, 'hoja': 'BASE DE DATOS', 'periodo': f'{anio}-{mes:02d}-01',
                       'valores': valores, 'permitir_reemplazos': True})
        solo_total.append(f'{anio}-{mes:02d}')
    cargas.sort(key=lambda c: c['periodo'])
    if not cargas:
        raise RuntimeError('No se pudo leer ningún trimestre: ' + '; '.join(fallos))

    periodos = [c['periodo'][:7] for c in cargas]
    resultado = {'fuente': 'Gasto provincial por finalidad y función (Ministerio de Economía)',
                 'periodo': periodos[-1], 'periodos': periodos, 'frecuencia': 'trimestral (acumulado del año)',
                 'fallos': fallos, 'solo_gasto_total': solo_total, 'datos': {c['periodo'][:7]: c['valores'] for c in cargas}}
    CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
    (CARPETA_REPORTES / f'gasto_provincial_{periodos[-1]}.json').write_text(
        json.dumps(resultado, ensure_ascii=False, indent=1), encoding='utf-8')

    if args.solo_extraer:
        print(json.dumps(resultado, ensure_ascii=False, indent=1, default=str))
        print('ESTADO=OK')
        return 0

    from sheet_directo import ClienteSheets, escribir_actualizaciones
    cliente = ClienteSheets()
    agregadas = asegurar_columnas(cliente)
    escritura = escribir_actualizaciones(cargas, None, cliente)
    info = escritura.get('por_fuente', {}).get(FUENTE, {})
    resultado.pop('datos', None)
    resultado['columnas_agregadas'] = agregadas
    resultado['actualizacion'] = {'celdas_actualizadas': info.get('escritas', 0),
                                  'cambios': info.get('cambios', [])[:60], 'errores': info.get('errores', [])}
    print(json.dumps(resultado, ensure_ascii=False, indent=1, default=str))
    resumen = {k: info.get(k) for k in ('periodos', 'escritas', 'ya_estaban', 'esperadas', 'conservadas',
                                        'no_encontradas', 'errores')}
    print('SHEET_DIRECTO=' + json.dumps(resumen, ensure_ascii=False, default=str))
    if fallos:
        print('NOTA_TELEGRAM=⚠️ Gasto provincial: no se pudo leer ' + '; '.join(fallos)[:250])
    if info.get('errores'):
        print('ERROR=' + '; '.join(info['errores'])[:500])
        print('ESTADO=ERROR')
        return 1
    if args.historico and not fallos:
        marca_historico.write_text(f'Carga histórica completa: {date.today().isoformat()}\n', encoding='utf-8')
    print('ESTADO=OK')
    return 0


def asegurar_columnas(cliente) -> list[str]:
    """Agrega las columnas que falten, a continuación del bloque ANSES (o BESS)."""
    from sheet_directo import API, HOJA, Mapa, letra
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    mapa = Mapa(encabezados, [])
    faltan = [t for t in encabezados_en_orden() if mapa.columna(t)[1] == 0]
    if not faltan:
        return []
    ocupadas = [i for i, t in enumerate(encabezados, start=1) if t not in (None, '')]
    ultima = max(ocupadas, default=1)
    propias = [i for i in ocupadas if str(encabezados[i - 1]).startswith(PREFIJO)]
    previas = propias or [i for i in ocupadas if str(encabezados[i - 1]).startswith('ANSES ')] \
        or [i for i in ocupadas if str(encabezados[i - 1]).startswith('BESS ')]
    despues_de = max(previas) if previas else ultima   # índice (1..n) de la última columna del bloque
    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(sheetId,title,gridProperties))'}, timeout=60)
    r.raise_for_status()
    props = next(s['properties'] for s in r.json()['sheets'] if s['properties']['title'] == HOJA)
    if despues_de < ultima:
        # Hay columnas después de ANSES (p. ej. Becas): se insertan columnas nuevas en ese lugar.
        cliente.sesion.post(f'{API}:batchUpdate', json={'requests': [{'insertDimension': {
            'range': {'sheetId': props['sheetId'], 'dimension': 'COLUMNS', 'startIndex': despues_de,
                      'endIndex': despues_de + len(faltan)}, 'inheritFromBefore': True}}]}, timeout=60).raise_for_status()
    else:
        necesarias = ultima + len(faltan) - props['gridProperties']['columnCount']
        if necesarias > 0:
            cliente.sesion.post(f'{API}:batchUpdate', json={'requests': [{'appendDimension': {
                'sheetId': props['sheetId'], 'dimension': 'COLUMNS', 'length': necesarias}}]}, timeout=60).raise_for_status()
    cliente.sesion.post(f'{API}/values:batchUpdate', json={'valueInputOption': 'RAW', 'data': [
        {'range': f"'{HOJA}'!{letra(despues_de + 1)}1:{letra(despues_de + len(faltan))}1", 'values': [faltan]}]},
        timeout=60).raise_for_status()
    return faltan


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as error:  # noqa: BLE001
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        raise SystemExit(1)
