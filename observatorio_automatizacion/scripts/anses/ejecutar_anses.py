"""ANSES - Estadísticas de la Seguridad Social: series MENSUALES por provincia.

Fuente oficial: ANSES, "Estadísticas de la Seguridad Social" (informe trimestral
con anexo Excel). https://www.anses.gob.ar/estudiosyestadisticas

Cada edición trae, por provincia, todos los meses del año en curso (y la edición
de diciembre, los doce meses del año). Las tablas mensuales por provincia
empiezan en enero de 2024: para 2023 ANSES no publica datos mensuales por
provincia (para 2023 la base tiene el BESS trimestral).

Diferencia con el BESS: ANSES cuenta todos los beneficios liquidados en el mes
(incluye suspensiones preventivas y pensiones coparticipadas) y asigna la
provincia por el banco de pago; por eso los números no coinciden exactamente con
el BESS. Se cargan en columnas propias "ANSES …" para no mezclar fuentes.

Cuadros (se ubican por título): P.5.1-P.5.6 (jubilaciones, pensiones y no
contributivos, beneficios y haber medio), D.4.1 (desempleo), F.5.1 (asignación
por hijo de trabajadores y monotributistas), H.4.1-H.4.2 (AUH), E.2.1 (AUE).

Actualización: el sitio de ANSES bloquea las descargas automáticas (protección
anti-robots). Cada trimestre, cuando ANSES publica un informe nuevo, alguien baja su
Excel (ESS_*.xlsx) y se lo manda al bot de Telegram (en el grupo, respondiendo a un
mensaje del bot, o por privado). La revisión lo recibe, lo guarda en fuentes/anses/
y lo carga. También se puede dejar el Excel directamente en fuentes/anses/.
El token del bot va en config/telegram.env (TELEGRAM_BOT_TOKEN=...).
La serie inicial (ene-2024 a mar-2026) está en anses_2024_2026.json.

Uso
    python ejecutar_anses.py              # procesa los Excel de fuentes/anses/
    python ejecutar_anses.py --historico  # además carga la serie inicial
    python ejecutar_anses.py --solo-extraer
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
import unicodedata
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))

SITIO = 'https://www.anses.gob.ar'
PAGINA = SITIO + '/estudiosyestadisticas'
CARPETA_REPORTES = RAIZ / 'scripts' / 'reportes' / 'anses'
FUENTE = 'anses'
DESDE = (2024, 1)

# clave, encabezado sin sufijo F/N, decimales
VARIABLES = [
    ('jub', 'ANSES Jubilaciones SIPA (cantidad)', 0),
    ('jub_mor', 'ANSES Jubilaciones SIPA por moratoria (cantidad)', 0),
    ('jub_haber', 'ANSES Haber medio jubilaciones SIPA (en pesos)', 0),
    ('pen', 'ANSES Pensiones SIPA (cantidad)', 0),
    ('pen_haber', 'ANSES Haber medio pensiones SIPA (en pesos)', 0),
    ('nc', 'ANSES Pensiones no contributivas (cantidad)', 0),
    ('nc_haber', 'ANSES Haber medio pensiones no contributivas (en pesos)', 0),
    ('desempleo', 'ANSES Prestación por desempleo - Beneficios (cantidad)', 0),
    ('aaff_hijo', 'ANSES Asignación por hijo (trabajadores y monotributistas) - Hijo (cantidad)', 0),
    ('aaff_disc', 'ANSES Asignación por hijo (trabajadores y monotributistas) - Hijo con discapacidad (cantidad)', 0),
    ('auh_hijo', 'ANSES AUH - Hijo (cantidad)', 0),
    ('auh_disc', 'ANSES AUH - Hijo con discapacidad (cantidad)', 0),
    ('auh_monto', 'ANSES AUH - Monto liquidado (en millones de pesos)', 2),
    ('aue', 'ANSES Asignación por embarazo - Beneficiarias (cantidad)', 0),
]
TITULO = {k: t for k, t, _ in VARIABLES}
DECIMALES = {k: d for k, _, d in VARIABLES}


def encabezados_en_orden() -> list[str]:
    return [f'{t} {s}' for s in ('F', 'N') for _, t, _ in VARIABLES]


def norm(x) -> str:
    s = unicodedata.normalize('NFD', str(x if x is not None else ''))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return ' '.join(s.split()).upper()


def a_periodo(x):
    if isinstance(x, (datetime, date)):
        return (x.year, x.month)
    if isinstance(x, (int, float)) and not isinstance(x, bool) and 40000 < x < 60000:
        d = date(1899, 12, 30) + timedelta(days=int(x))
        return (d.year, d.month)
    return None


def tablas(contenido: bytes) -> dict:
    from openpyxl import load_workbook
    libro = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    salida = {}
    for ws in libro.worksheets:
        codigo = norm(ws.title).rstrip('.').strip()
        filas = [list(f) for f in ws.iter_rows(values_only=True)]
        salida[codigo] = [f for f in filas if any(v not in (None, '') for v in f)]
    return salida


def cuadro(filas: list):
    """Encabezado de provincias y filas mensuales de un cuadro por provincia."""
    i_cab = next(i for i, f in enumerate(filas) if any(norm(x) == 'FORMOSA' for x in f))
    cab = [norm(x) for x in filas[i_cab]]
    meses = {}
    for f in filas[i_cab + 1:]:
        p = a_periodo(f[0]) if f else None
        if p:
            meses[p] = f
    if not meses:
        raise ValueError('sin filas mensuales')
    provincias = [j for j, x in enumerate(cab) if x and x not in ('PERIODO', 'TOTAL')]
    return cab, provincias, meses


def num(f, j):
    v = f[j] if j < len(f) else None
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def serie_par(filas: list) -> dict:
    """Cuadros con dos subcolumnas por provincia (p. ej. No moratoria / Moratoria).
    Devuelve {p: {'F': (a, b), 'N': (a, b), 'total': t}}; Nación = suma de provincias
    (se controla contra la columna Total del cuadro)."""
    cab, provincias, meses = cuadro(filas)
    fi, ti = cab.index('FORMOSA'), cab.index('TOTAL')
    salida = {}
    for p, f in meses.items():
        a = sum(num(f, j) for j in provincias)
        b = sum(num(f, j + 1) for j in provincias)
        salida[p] = {'F': (num(f, fi), num(f, fi + 1)), 'N': (a, b), 'total': num(f, ti)}
    return salida


def serie_simple(filas: list) -> dict:
    cab, _, meses = cuadro(filas)
    fi, ti = cab.index('FORMOSA'), cab.index('TOTAL')
    return {p: {'F': num(f, fi), 'N': num(f, ti)} for p, f in meses.items()}


# Desde el III trimestre 2026 ANSES publica tres informes separados (Pasivos;
# Asignaciones Familiares y Universales; Activos y Desempleo), cada uno con su
# Excel. Se carga lo que traiga cada archivo: no hace falta que estén todos.
GRUPOS = {
    'Pasivos': ('P.5.1', 'P.5.2', 'P.5.3', 'P.5.4', 'P.5.5', 'P.5.6'),
    'Desempleo': ('D.4.1',),
    'Asignaciones familiares y AUH': ('F.5.1', 'H.4.1', 'H.4.2', 'E.2.1'),
}


def grupos_presentes(t: dict) -> list[str]:
    return [g for g, cuadros in GRUPOS.items() if all(c in t for c in cuadros)]


def extraer(contenido: bytes, informar: list | None = None) -> dict:
    t = tablas(contenido)
    presentes = grupos_presentes(t)
    if not presentes:
        raise ValueError('no trae los cuadros de ANSES por provincia (P.5.x, D.4.1, F.5.1, H.4.x, E.2.1)')
    if informar is not None:
        informar.extend(presentes)
    datos: dict = {}
    if 'Pasivos' in presentes:
        J, HJ, P, HP = (serie_par(t[c]) for c in ('P.5.1', 'P.5.2', 'P.5.3', 'P.5.4'))
        NC, HNC = (serie_simple(t[c]) for c in ('P.5.5', 'P.5.6'))
        for p in J:
            if abs(sum(J[p]['N']) - J[p]['total']) > 1:
                raise ValueError(f'P.5.1 {p}: la suma de provincias no da el total')
            v = datos.setdefault(p, {})
            jf, pf = J[p]['F'], P[p]['F']
            v['jub_F'] = sum(jf)
            v['jub_mor_F'] = jf[1]
            v['jub_haber_F'] = (jf[0] * HJ[p]['F'][0] + jf[1] * HJ[p]['F'][1]) / sum(jf)
            v['pen_F'] = sum(pf)
            v['pen_haber_F'] = (pf[0] * HP[p]['F'][0] + pf[1] * HP[p]['F'][1]) / sum(pf)
            v['jub_N'] = J[p]['total']
            v['jub_mor_N'] = J[p]['N'][1]
            v['jub_haber_N'] = HJ[p]['total']
            v['pen_N'] = P[p]['total']
            v['pen_haber_N'] = HP[p]['total']
            for suf in ('F', 'N'):
                v[f'nc_{suf}'] = NC[p][suf]
                v[f'nc_haber_{suf}'] = HNC[p][suf]
    if 'Desempleo' in presentes:
        D = serie_simple(t['D.4.1'])
        for p in D:
            v = datos.setdefault(p, {})
            for suf in ('F', 'N'):
                v[f'desempleo_{suf}'] = D[p][suf]
    if 'Asignaciones familiares y AUH' in presentes:
        A, U, M, E = (serie_par(t[c]) for c in ('F.5.1', 'H.4.1', 'H.4.2', 'E.2.1'))
        for p in A:
            v = datos.setdefault(p, {})
            for suf in ('F', 'N'):
                v[f'aaff_hijo_{suf}'] = A[p][suf][0]
                v[f'aaff_disc_{suf}'] = A[p][suf][1]
                if p in U:
                    v[f'auh_hijo_{suf}'] = U[p][suf][0]
                    v[f'auh_disc_{suf}'] = U[p][suf][1]
                if p in E:
                    v[f'aue_{suf}'] = E[p][suf][0]
            if p in M:
                v['auh_monto_F'] = sum(M[p]['F']) / 1000   # miles de pesos -> millones
                v['auh_monto_N'] = M[p]['total'] / 1000
    return datos


def a_cargas(datos: dict) -> list[dict]:
    cargas = []
    for p in sorted(datos):
        if p < DESDE:
            continue
        valores = {}
        for clave, valor in datos[p].items():
            base, suf = clave.rsplit('_', 1)
            valores[f'{TITULO[base]} {suf}'] = round(valor, DECIMALES[base]) if DECIMALES[base] else int(round(valor))
        cargas.append({'fuente': FUENTE, 'hoja': 'BASE DE DATOS', 'periodo': f'{p[0]}-{p[1]:02d}-01', 'valores': valores})
    return cargas


def asegurar_columnas(cliente) -> list[str]:
    from sheet_directo import API, HOJA, Mapa, clave, letra
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    # Títulos anteriores (sin "(cantidad)") o con espacios de más: se dejan con el nombre actual.
    nuevos = {clave(n): n for n in encabezados_en_orden()}
    viejos = {clave(n.replace(' (cantidad)', '')): n for n in encabezados_en_orden()}
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
        cliente.sesion.post(f'{API}:batchUpdate', json={'requests': [{'appendDimension': {
            'sheetId': props['sheetId'], 'dimension': 'COLUMNS', 'length': necesarias}}]}, timeout=60).raise_for_status()
    cliente.sesion.post(f'{API}/values:batchUpdate', json={'valueInputOption': 'RAW', 'data': [
        {'range': f"'{HOJA}'!{letra(ultima + 1)}1:{letra(ultima + len(faltan))}1", 'values': [faltan]}]},
        timeout=60).raise_for_status()
    return faltan


CARPETA_FUENTES = RAIZ / 'fuentes' / 'anses'
CONFIG_TELEGRAM = RAIZ / 'config' / 'telegram.env'
CHAT_GRUPO = '-5396908101'
ESTADO_TELEGRAM = RAIZ / 'scripts' / 'reportes' / 'anses' / 'telegram_estado.json'


def token_telegram() -> str | None:
    """Token del bot: variable TELEGRAM_BOT_TOKEN o archivo config/telegram.env (lo carga la responsable)."""
    import os
    token = os.environ.get('TELEGRAM_BOT_TOKEN')
    if not token and CONFIG_TELEGRAM.is_file():
        for linea in CONFIG_TELEGRAM.read_text(encoding='utf-8').splitlines():
            if linea.strip().startswith('TELEGRAM_BOT_TOKEN='):
                token = linea.split('=', 1)[1].strip()
    # Se toma solo el token (números:letras), aunque se haya pegado con
    # "HTTP API:" adelante, comillas o espacios.
    encontrado = re.search(r'\d{6,}:[A-Za-z0-9_-]{30,}', token or '')
    return encontrado.group(0) if encontrado else None


def _telegram(token: str, metodo: str, **params):
    url = f'https://api.telegram.org/bot{token}/{metodo}'
    datos = urllib.parse.urlencode(params).encode() if params else None
    with urllib.request.urlopen(urllib.request.Request(url, data=datos), timeout=60) as r:
        respuesta = json.loads(r.read().decode('utf-8'))
    if not respuesta.get('ok'):
        raise RuntimeError(respuesta.get('description', 'error de Telegram'))
    return respuesta['result']


def _estado() -> dict:
    try:
        return json.loads(ESTADO_TELEGRAM.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return {}


def recibir_por_telegram() -> tuple[list[str], str | None]:
    """Baja a fuentes/anses los Excel enviados al bot @Obs_formosa345_bot (por privado, o en el
    grupo si el bot está en el grupo con la privacidad desactivada). Devuelve (archivos guardados, aviso)."""
    token = token_telegram()
    if not token:
        return [], 'falta el token del bot en config/telegram.env'
    estado = _estado()
    try:
        novedades = _telegram(token, 'getUpdates', offset=estado.get('offset', 0), timeout=0,
                              allowed_updates=json.dumps(['message']))
    except Exception as error:  # noqa: BLE001
        return [], f'no se pudieron leer los mensajes del bot ({error})'
    guardados = []
    for u in novedades:
        estado['offset'] = u['update_id'] + 1
        m = u.get('message') or {}
        doc = m.get('document') or {}
        chat = m.get('chat') or {}
        nombre = doc.get('file_name') or ''
        if not nombre.lower().endswith(('.xlsx', '.xls')):
            continue
        # El bot solo está en los grupos donde lo agregó el equipo: se aceptan
        # archivos de esos grupos (aunque Telegram cambie el número del grupo) y por privado.
        if chat.get('type') not in ('private', 'group', 'supergroup'):
            continue
        try:
            ruta_tg = _telegram(token, 'getFile', file_id=doc['file_id'])['file_path']
            with urllib.request.urlopen(f'https://api.telegram.org/file/bot{token}/{ruta_tg}', timeout=120) as r:
                contenido = r.read()
            partes = []
            extraer(contenido, partes)  # valida que sea un Excel de Estadísticas de la Seguridad Social
        except Exception as error:  # noqa: BLE001
            try:
                _telegram(token, 'sendMessage', chat_id=chat['id'],
                          text=f'⚠️ {nombre}: no es el Excel de "Estadísticas de la Seguridad Social" de ANSES '
                               f'o no se pudo leer ({str(error)[:150]}). No se cargó.')
            except Exception:  # noqa: BLE001
                pass
            continue
        CARPETA_FUENTES.mkdir(parents=True, exist_ok=True)
        destino = CARPETA_FUENTES / f"{datetime.now():%Y%m%d_%H%M%S}_{re.sub(r'[^A-Za-z0-9._-]', '_', nombre)}"
        destino.write_bytes(contenido)
        guardados.append(destino.name)
        try:
            _telegram(token, 'sendMessage', chat_id=chat['id'],
                      text=f'✅ Recibí {nombre} ({", ".join(partes)}). Lo cargo en el Sheet ahora; te llega el resumen en el aviso de la revisión.')
        except Exception:  # noqa: BLE001
            pass
    ESTADO_TELEGRAM.parent.mkdir(parents=True, exist_ok=True)
    ESTADO_TELEGRAM.write_text(json.dumps(estado), encoding='utf-8')
    return guardados, None


def recordatorio(ultimo: str) -> str | None:
    """Una vez por día, si ya debería haber salido el informe siguiente, lo recuerda en el aviso."""
    anio, mes = int(ultimo[:4]), int(ultimo[5:7])
    mes_sig = mes + 3
    anio_sig = anio + (mes_sig - 1) // 12
    mes_sig = (mes_sig - 1) % 12 + 1
    # ANSES publica el informe unos 3 meses después de terminado el trimestre
    # (ej.: "ESS_Junio2026" = I trimestre 2026, datos a marzo).
    esperado = date(anio_sig + (mes_sig + 3 - 1) // 12, (mes_sig + 3 - 1) % 12 + 1, 1)
    hoy = date.today()
    estado = _estado()
    if hoy < esperado or estado.get('recordado') == hoy.isoformat():
        return None
    estado['recordado'] = hoy.isoformat()
    ESTADO_TELEGRAM.parent.mkdir(parents=True, exist_ok=True)
    ESTADO_TELEGRAM.write_text(json.dumps(estado), encoding='utf-8')
    meses = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
    return (f'📥 Ya debería estar el informe de ANSES con datos a {meses[mes_sig - 1]} {anio_sig % 100:02d}: '
            'bajá el Excel de anses.gob.ar/estudiosyestadisticas y mandalo a @Obs_formosa345_bot (por privado o en el grupo). '
            'Desde sep-26 ANSES publica tres informes separados (Pasivos; Asignaciones Familiares y Universales; '
            'Activos y Desempleo): mandá el Excel de cada uno cuando salga.')


def main() -> int:
    ap = argparse.ArgumentParser(description='ANSES - Seguridad social mensual por provincia (Formosa y Total)')
    ap.add_argument('--historico', action='store_true', help='Carga también la serie inicial guardada (ene-2024 a mar-2026).')
    ap.add_argument('--solo-extraer', action='store_true')
    args = ap.parse_args()

    # El sitio de ANSES bloquea las descargas automáticas: cada trimestre se baja a mano
    # el Excel del informe "Estadísticas de la Seguridad Social" a fuentes/anses/.
    datos: dict = {}
    usadas, fallos = [], []
    if args.historico:
        inicial = json.loads((Path(__file__).with_name('anses_2024_2026.json')).read_text(encoding='utf-8'))
        for p, valores in inicial['datos'].items():
            datos[(int(p[:4]), int(p[5:7]))] = {'__titulos__': valores}
        usadas.append('serie inicial anses_2024_2026.json')
    recibidos, aviso_telegram = recibir_por_telegram()
    CARPETA_FUENTES.mkdir(parents=True, exist_ok=True)
    for ruta in sorted(CARPETA_FUENTES.glob('*.xls*'), key=lambda r: r.stat().st_mtime):
        try:
            parcial = extraer(ruta.read_bytes())
        except Exception as error:  # noqa: BLE001
            fallos.append(f'{ruta.name}: {error}')
            continue
        for p, v in parcial.items():  # el archivo más nuevo corrige al anterior (solo lo que trae)
            datos.setdefault(p, {}).update(v)
        usadas.append(ruta.name)
    inicial_json = json.loads((Path(__file__).with_name('anses_2024_2026.json')).read_text(encoding='utf-8'))
    ultimo_cargado = max([*inicial_json['datos'], *(f'{p[0]}-{p[1]:02d}' for p in datos)])
    nota = recordatorio(ultimo_cargado)
    if nota:
        print('NOTA_TELEGRAM=' + nota)
    if aviso_telegram:
        estado = _estado()
        if estado.get('avisado') != date.today().isoformat():  # una vez por día
            estado['avisado'] = date.today().isoformat()
            ESTADO_TELEGRAM.parent.mkdir(parents=True, exist_ok=True)
            ESTADO_TELEGRAM.write_text(json.dumps(estado), encoding='utf-8')
            print('NOTA_TELEGRAM=⚠️ Bot de ANSES: ' + aviso_telegram)
    if not datos:
        inicial = json.loads((Path(__file__).with_name('anses_2024_2026.json')).read_text(encoding='utf-8'))
        ultimo = max(inicial['datos'])
        resultado = {'fuente': 'ANSES - Seguridad social mensual por provincia', 'periodo': ultimo, 'periodos': [ultimo],
                     'detalle': 'Sin Excel nuevo en fuentes/anses: cuando ANSES publique el próximo informe trimestral, '
                                'descargarlo en esa carpeta.', 'fallos': fallos}
        print(json.dumps(resultado, ensure_ascii=False))
        print('ESTADO=OK')
        return 0
    cargas = []
    for p in sorted(datos):
        if p < DESDE:
            continue
        v = datos[p]
        if '__titulos__' in v:
            valores = v['__titulos__']
        else:
            valores = {}
            for clave, valor in v.items():
                base, suf = clave.rsplit('_', 1)
                valores[f'{TITULO[base]} {suf}'] = round(valor, DECIMALES[base]) if DECIMALES[base] else int(round(valor))
        cargas.append({'fuente': FUENTE, 'hoja': 'BASE DE DATOS', 'periodo': f'{p[0]}-{p[1]:02d}-01', 'valores': valores})
    periodos = [c['periodo'][:7] for c in cargas]
    resultado = {'fuente': 'ANSES - Seguridad social mensual por provincia', 'periodo': periodos[-1],
                 'periodos': periodos, 'frecuencia': 'mensual (publicación trimestral)', 'archivos': usadas, 'fallos': fallos,
                 'recibidos_por_telegram': recibidos,
                 'datos': {c['periodo'][:7]: c['valores'] for c in cargas}}
    CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
    (CARPETA_REPORTES / f'anses_{periodos[-1]}.json').write_text(json.dumps(resultado, ensure_ascii=False, indent=1), encoding='utf-8')
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
    resultado['actualizacion'] = {'celdas_actualizadas': info.get('escritas', 0), 'cambios': info.get('cambios', [])[:60],
                                  'errores': info.get('errores', [])}
    print(json.dumps(resultado, ensure_ascii=False, indent=1, default=str))
    print('SHEET_DIRECTO=' + json.dumps({k: info.get(k) for k in ('periodos', 'escritas', 'ya_estaban', 'esperadas',
                                                                  'conservadas', 'no_encontradas', 'errores')},
                                        ensure_ascii=False, default=str))
    if info.get('errores'):
        print('ERROR=' + '; '.join(info['errores'])[:500])
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
