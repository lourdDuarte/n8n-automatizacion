"""Becas para estudiantes - Formosa (F) y Total del país (N).

Gasto mensual (automático)
    Fuente: Presupuesto Abierto, Ministerio de Economía (sistema eSidif), datos abiertos
    "Presupuesto de gastos y su ejecución detallada - agrupación mensual".
    https://www.presupuestoabierto.gob.ar/sici/datos-abiertos
    Se toma el crédito DEVENGADO (en millones de pesos) de las partidas "Becas"
    (inciso Transferencias) del programa de becas a estudiantes, por ubicación
    geográfica: Formosa y total del país (todas las ubicaciones).
      * Progresar ................ actividad "Otorgamiento de Becas Educativas: Progresar"
      * Becas socioeducativas .... actividad "Otorgamiento de Becas Socioeducativas"
      * Becas Manuel Belgrano .... actividad "Promoción de Carreras Estratégicas"
    Solo se cargan meses cerrados (el mes en curso todavía está incompleto).
    El devengado es el momento en que el Estado registra el gasto: los meses saltan por
    el calendario de convocatorias, los retroactivos y la liberación del 20 % retenido;
    por eso se agregan el acumulado del año y la participación de Formosa.

Cantidad de becarios (anual, carga fija)
    Fuente: Informes del Jefe de Gabinete de Ministros al Congreso (respuestas del
    Ministerio de Capital Humano). Se cargan en la fila de diciembre de cada año.

Uso
    python ejecutar_becas.py               # año en curso (revisión diaria)
    python ejecutar_becas.py --historico   # desde enero 2023 + becarios anuales
    python ejecutar_becas.py --solo-extraer
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import sys
import time
import unicodedata
import urllib.request
import zipfile
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))

FUENTE = 'becas'
URL = 'https://dgsiaf-repo.mecon.gob.ar/repository/pa/datasets/{anio}/credito-mensual-{anio}.zip'
CARPETA_FUENTES = RAIZ / 'fuentes' / 'becas'
CARPETA_REPORTES = RAIZ / 'scripts' / 'reportes' / 'becas'
DESDE = 2023
FORMOSA = 'PROVINCIA DE FORMOSA'

GASTO = [
    ('progresar', 'Progresar - Gasto devengado (en millones de pesos)'),
    ('progresar_acum', 'Progresar - Gasto devengado acumulado del año (en millones de pesos)'),
    ('socio', 'Becas socioeducativas - Gasto devengado (en millones de pesos)'),
    ('belgrano', 'Becas Manuel Belgrano (carreras estratégicas) - Gasto devengado (en millones de pesos)'),
]
PARTICIPACION = 'Progresar - Participación de Formosa en el gasto del país, acumulado del año (%) F'
BECARIOS = [
    ('total', 'Progresar - Becarios total, anual (cantidad)'),
    ('obligatorio', 'Progresar - Becarios nivel obligatorio, anual (cantidad)'),
    ('superior', 'Progresar - Becarios nivel superior, anual (cantidad)'),
    ('trabajo', 'Progresar - Becarios Progresar Trabajo, anual (cantidad)'),
    ('fortalecimiento', 'Becas de fortalecimiento educativo - Becarios, anual (cantidad)'),
]
# Becarios activos a diciembre (o del año) según los Informes del Jefe de Gabinete al Congreso.
# 2023: Informe 140 (Diputados, 2024), preguntas 34: nivel obligatorio y superior a dic-2023.
# 2024: Informe 143 (Senado, 2025), tabla de becarios activos 2024 por línea; Informe 142
#       (Diputados, 2025) para becas de fortalecimiento educativo 2024.
# 2025: Informe 145 (Diputados, 2026), pregunta 335: becarios Progresar 2025.
BECARIOS_ANUALES = {
    2023: {'obligatorio': (22181, 995288), 'superior': (12191, 682735)},
    2024: {'total': (36817, 1359274), 'obligatorio': (19554, 711946), 'superior': (10079, 559540),
           'trabajo': (7184, 87788), 'fortalecimiento': (4231, 33426)},
    2025: {'total': (25260, 953169)},
}


def encabezados_en_orden() -> list[str]:
    salida = []
    for suf in ('F', 'N'):
        salida += [f'{t} {suf}' for _, t in GASTO]
    salida.append(PARTICIPACION)
    for suf in ('F', 'N'):
        salida += [f'{t} {suf}' for _, t in BECARIOS]
    return salida


def norm(x) -> str:
    s = unicodedata.normalize('NFD', str(x or ''))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return ' '.join(s.split()).upper()


def categoria(actividad: str) -> str | None:
    a = norm(actividad)
    if 'PROGRESAR' in a and 'OTORGAMIENTO' in a:
        return 'progresar'
    if 'SOCIOEDUCATIV' in a and 'BECA' in a:
        return 'socio'
    if 'CARRERAS ESTRATEGICAS' in a or 'BELGRANO' in a:
        return 'belgrano'
    return None


def numero(x) -> float:
    try:
        return float(str(x).replace(',', '.'))
    except ValueError:
        return 0.0


def archivo_anio(anio: int) -> Path:
    """Descarga el zip del año (el del año en curso se renueva cada 20 horas)."""
    CARPETA_FUENTES.mkdir(parents=True, exist_ok=True)
    ruta = CARPETA_FUENTES / f'credito-mensual-{anio}.zip'
    vigente = anio == date.today().year
    if ruta.is_file() and (not vigente or time.time() - ruta.stat().st_mtime < 20 * 3600):
        return ruta
    pedido = urllib.request.Request(URL.format(anio=anio), headers={'User-Agent': 'Mozilla/5.0 (Observatorio Formosa)'})
    with urllib.request.urlopen(pedido, timeout=180) as r:
        contenido = r.read()
    if not contenido.startswith(b'PK'):
        raise RuntimeError(f'Presupuesto Abierto {anio}: la descarga no es un zip')
    temporal = ruta.with_suffix('.tmp')
    temporal.write_bytes(contenido)
    temporal.replace(ruta)
    return ruta


def leer_anio(ruta: Path) -> tuple[dict, dict]:
    """{(anio, mes): {cat_F, cat_N}} y las actividades encontradas con su total."""
    datos: dict = {}
    actividades: dict = {}
    with zipfile.ZipFile(ruta) as z:
        nombre = next(n for n in z.namelist() if n.lower().endswith('.csv'))
        with z.open(nombre) as f:
            texto = io.TextIOWrapper(f, encoding='utf-8-sig', newline='')
            encabezado = next(csv.reader([texto.readline()]))
            idx = {c: i for i, c in enumerate(encabezado)}
            for linea in texto:
                if 'Becas' not in linea:      # filtro rápido: solo las partidas de becas
                    continue
                fila = next(csv.reader([linea]))
                if norm(fila[idx['parcial_desc']]) != 'BECAS':
                    continue
                cat = categoria(fila[idx['actividad_desc']])
                if not cat:
                    continue
                anio = int(fila[idx['impacto_presupuestario_anio']])
                mes = int(fila[idx['impacto_presupuestario_mes']])
                monto = numero(fila[idx['credito_devengado']])
                v = datos.setdefault((anio, mes), {})
                v[f'{cat}_N'] = v.get(f'{cat}_N', 0.0) + monto
                if norm(fila[idx['ubicacion_geografica_desc']]) == FORMOSA:
                    v[f'{cat}_F'] = v.get(f'{cat}_F', 0.0) + monto
                clave = f"{fila[idx['jurisdiccion_desc']]} | prog {fila[idx['programa_id']]} | {fila[idx['actividad_desc']]}"
                actividades[clave] = round(actividades.get(clave, 0.0) + monto, 1)
    return datos, actividades


def main() -> int:
    ap = argparse.ArgumentParser(description='Becas para estudiantes (Formosa y total del país)')
    ap.add_argument('--historico', action='store_true', help='Carga desde enero 2023 y los becarios anuales.')
    ap.add_argument('--solo-extraer', action='store_true', help='No escribe en el Sheet.')
    args = ap.parse_args()

    hoy = date.today()
    # La primera vez que corre (revisión diaria incluida) hace la carga histórica completa.
    marca_historico = CARPETA_REPORTES / 'historico_cargado.txt'
    if not args.historico and not args.solo_extraer and not marca_historico.is_file():
        args.historico = True
    anios = list(range(DESDE, hoy.year + 1)) if args.historico else [hoy.year]
    datos: dict = {}
    actividades: dict = {}
    fallos = []
    for anio in anios:
        try:
            d, a = leer_anio(archivo_anio(anio))
        except Exception as error:  # noqa: BLE001
            fallos.append(f'{anio}: {error}')
            continue
        datos.update(d)
        actividades.update({f'{anio} | {k}': v for k, v in a.items()})
    if not datos:
        raise RuntimeError('No se pudo leer Presupuesto Abierto: ' + '; '.join(fallos))

    cargas = []
    acum = {}
    for (anio, mes) in sorted(datos):
        if anio < DESDE or (anio, mes) >= (hoy.year, hoy.month):
            continue  # mes en curso: incompleto
        v = datos[(anio, mes)]
        if anio not in acum:
            acum[anio] = {'F': 0.0, 'N': 0.0}
        valores = {}
        for suf in ('F', 'N'):
            acum[anio][suf] += v.get(f'progresar_{suf}', 0.0)
            valores[f'{GASTO[0][1]} {suf}'] = round(v.get(f'progresar_{suf}', 0.0), 2)
            valores[f'{GASTO[1][1]} {suf}'] = round(acum[anio][suf], 2)
            valores[f'{GASTO[2][1]} {suf}'] = round(v.get(f'socio_{suf}', 0.0), 2)
            valores[f'{GASTO[3][1]} {suf}'] = round(v.get(f'belgrano_{suf}', 0.0), 2)
        if acum[anio]['N'] > 0:
            valores[PARTICIPACION] = round(acum[anio]['F'] / acum[anio]['N'] * 100, 2)
        if args.historico and mes == 12 and anio in BECARIOS_ANUALES:
            for clave, (f, n) in BECARIOS_ANUALES[anio].items():
                titulo = dict(BECARIOS)[clave]
                valores[f'{titulo} F'] = f
                valores[f'{titulo} N'] = n
        cargas.append({'fuente': FUENTE, 'hoja': 'BASE DE DATOS', 'periodo': f'{anio}-{mes:02d}-01',
                       'valores': valores, 'permitir_reemplazos': True})
    if not cargas:
        raise RuntimeError('Presupuesto Abierto no trae meses cerrados con datos de becas.')
    periodos = [c['periodo'][:7] for c in cargas]
    resultado = {'fuente': 'Becas para estudiantes (Presupuesto Abierto)', 'periodo': periodos[-1], 'periodos': periodos,
                 'frecuencia': 'mensual (gasto devengado); becarios anual', 'fallos': fallos,
                 'actividades_encontradas': actividades,
                 'datos': {c['periodo'][:7]: c['valores'] for c in cargas}}
    CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
    (CARPETA_REPORTES / f'becas_{periodos[-1]}.json').write_text(json.dumps(resultado, ensure_ascii=False, indent=1), encoding='utf-8')

    if args.solo_extraer:
        resultado.pop('datos', None)
        resultado['ultimo_mes'] = cargas[-1]['valores']
        resultado['acumulado_por_anio'] = {a: {k: round(x, 1) for k, x in v.items()} for a, v in acum.items()}
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
        print('NOTA_TELEGRAM=⚠️ Becas: no se pudo leer ' + '; '.join(fallos)[:250])
    if info.get('errores'):
        print('ERROR=' + '; '.join(info['errores'])[:500])
        print('ESTADO=ERROR')
        return 1
    if args.historico and not fallos:
        marca_historico.write_text(f'Carga histórica completa: {date.today().isoformat()}\n', encoding='utf-8')
    print('ESTADO=OK')
    return 0


def asegurar_columnas(cliente) -> list[str]:
    """Agrega las columnas de becas que falten, a continuación del bloque ANSES (o BESS)."""
    from sheet_directo import API, HOJA, Mapa, letra
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    mapa = Mapa(encabezados, [])
    faltan = [t for t in encabezados_en_orden() if mapa.columna(t)[1] == 0]
    if not faltan:
        return []
    ocupadas = [i for i, t in enumerate(encabezados, start=1) if t not in (None, '')]
    ultima = max(ocupadas, default=1)
    previas = [i for i in ocupadas if str(encabezados[i - 1]).startswith(('ANSES ', 'BESS ', 'Progresar ', 'Becas '))]
    despues_de = max(previas) if previas else ultima   # índice (1..n) de la última columna del bloque
    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(sheetId,title,gridProperties))'}, timeout=60)
    r.raise_for_status()
    props = next(s['properties'] for s in r.json()['sheets'] if s['properties']['title'] == HOJA)
    if despues_de < ultima:
        # Hay columnas después de ANSES: se insertan columnas nuevas en ese lugar.
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
