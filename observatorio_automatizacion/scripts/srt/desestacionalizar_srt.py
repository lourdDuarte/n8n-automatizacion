"""SRT Formosa - Trabajadores registrados, serie desestacionalizada (estimación DAEE).

La SRT publica sus datos tal como salen (serie original). Todos los años hay meses que suben o
bajan por la época del año (más trabajadores en diciembre y julio, menos en febrero y marzo), y
eso confunde la comparación con el mes anterior. Este script estima la serie sin ese efecto,
con el mismo método de medias móviles que usa el SIPA (X-11, versión simplificada), y la
escribe en el Sheet en tres columnas propias, a continuación del bloque SRT:

* SRT Trabajadores desestacionalizado F (estimación DAEE)  -> NÚMERO que calcula este programa.
  No puede ser una fórmula de fila: para cada mes el método usa toda la serie (medias móviles
  de varios meses y el comportamiento del mismo mes en todos los años), y cada mes nuevo
  corrige un poco los anteriores.
* SRT Var % intermensual trabajadores desestacionalizado F (estimación DAEE)  -> FÓRMULA
  (=nivel del mes / nivel del mes anterior * 100 - 100), igual que las demás columnas SRT.
* SRT Dif intermensual trabajadores desestacionalizado F (estimación DAEE)    -> FÓRMULA
  (=nivel del mes - nivel del mes anterior).

Los encabezados llevan una nota (se ve al pasar el mouse) que explica de dónde salen.

Qué se desestacionaliza y qué no:
* Trabajadores (total): sí. Tiene un patrón estacional claro (cerca de 4 puntos entre el mes
  más alto y el más bajo) y una serie continua desde enero de 2021.
* Empleadores: no. El efecto de la época del año es muy chico (alrededor de 1 %).
* Sectores: no. La serie por sector empieza en 2023 y es muy irregular (por ejemplo, la caída
  de la construcción en 2024); además los sectores se informan en variación interanual, que
  no necesita desestacionalizar.

Es una estimación propia: cada mes nuevo puede corregir un poco los últimos meses (pasa lo
mismo con la serie oficial del SIPA). Por eso se reescribe la serie completa en cada corrida.

Datos atípicos conocidos (atipicos_desestacionalizacion.json): meses con un pico de un solo
mes en un sector (p. ej. Administración pública, junio 2025). Ese pico se descuenta solo para
calcular los factores estacionales; el dato publicado no se toca.

Uso
    python desestacionalizar_srt.py               # calcula y escribe en el Sheet
    python desestacionalizar_srt.py --solo-calcular
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))

FUENTE = 'srt_desestacionalizada'
ORIGEN = 'SRT Evolución Cantidad de Trabajadores F'
SUFIJO = ' F (estimación DAEE)'
COL_NIVEL = 'SRT Trabajadores desestacionalizado' + SUFIJO
COL_VAR = 'SRT Var % intermensual trabajadores desestacionalizado' + SUFIJO
COL_DIF = 'SRT Dif intermensual trabajadores desestacionalizado' + SUFIJO
COLUMNAS = [COL_NIVEL, COL_VAR, COL_DIF]
ATIPICOS = Path(__file__).with_name('atipicos_desestacionalizacion.json')
MINIMO_MESES = 36

# ------------------------------------------------------------------ método (X-11 simplificado)
H13 = [-0.01935, -0.02786, 0.0, 0.06549, 0.14736, 0.21434, 0.24006, 0.21434, 0.14736, 0.06549, 0.0, -0.02786, -0.01935]
M2X12 = [1 / 24] + [1 / 12] * 11 + [1 / 24]
M3X3 = [1 / 9, 2 / 9, 3 / 9, 2 / 9, 1 / 9]
M3X5 = [1 / 15, 2 / 15, 3 / 15, 3 / 15, 3 / 15, 2 / 15, 1 / 15]


def _espejo(x, k):
    return x[k:0:-1] + x + x[-2:-k - 2:-1]


def _mm(x, pesos):
    k = len(pesos) // 2
    if len(x) <= k:
        return [sum(x) / len(x)] * len(x)
    e = _espejo(x, k)
    return [sum(p * e[i + j] for j, p in enumerate(pesos)) for i in range(len(x))]


def _por_mes(si, pesos):
    n = len(si)
    s = [0.0] * n
    for m in range(12):
        idx = list(range(m, n, 12))
        vals = [si[i] for i in idx]
        suav = _mm(vals, pesos) if len(vals) >= 3 else [sum(vals) / len(vals)] * len(vals)
        for i, v in zip(idx, suav):
            s[i] = v
    return s


def _extremos(si, limite=2.0):
    n = len(si)
    out = si[:]
    for m in range(12):
        idx = list(range(m, n, 12))
        vals = [si[i] for i in idx]
        if len(vals) < 3:
            continue
        media = sum(vals) / len(vals)
        desv = math.sqrt(sum((v - media) ** 2 for v in vals) / len(vals)) or 1e-9
        for i in idx:
            if abs(si[i] - media) > limite * desv:
                otros = [si[j] for j in idx if j != i]
                out[i] = sum(otros) / len(otros)
    return out


def _centrar(s):
    t = _mm(s, M2X12)
    return [a - b for a, b in zip(s, t)]


def factores_estacionales(valores: list[float]) -> list[float]:
    """Factores multiplicativos de cada mes (1,02 = ese mes suele estar 2 % arriba)."""
    if len(valores) < MINIMO_MESES or any(v is None or v <= 0 for v in valores):
        raise ValueError(f'hacen falta al menos {MINIMO_MESES} meses seguidos con datos')
    y = [math.log(v) for v in valores]
    t1 = _mm(y, M2X12)
    s1 = _centrar(_por_mes(_extremos([a - b for a, b in zip(y, t1)]), M3X3))
    t2 = _mm([a - b for a, b in zip(y, s1)], H13)
    s2 = _centrar(_por_mes(_extremos([a - b for a, b in zip(y, t2)]), M3X5))
    return [math.exp(b) for b in s2]


# ------------------------------------------------------------------ datos
def _numero(v):
    from sheet_directo import a_numero
    return a_numero(v)


def leer_series(cliente):
    from sheet_directo import Mapa, letra
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    periodos = [f[0] if f else None for f in cliente.leer('A:A')]
    mapa = Mapa(encabezados, periodos)
    atipicos = json.loads(ATIPICOS.read_text(encoding='utf-8')) if ATIPICOS.is_file() else []
    titulos = [ORIGEN] + sorted({a['columna'] for a in atipicos})
    columnas = {}
    for t in titulos:
        col, n = mapa.columna(t)
        if col is None:
            raise ValueError(f'No se encontró la columna "{t}" en el Sheet ({n} coincidencias)')
        datos = cliente.leer(f'{letra(col)}:{letra(col)}')
        columnas[t] = {f: (datos[f - 1][0] if f - 1 < len(datos) and datos[f - 1] else None) for f in mapa.periodo_por_fila}
    meses = sorted(mapa.periodo_por_fila.items(), key=lambda x: x[1])   # [(fila, (año, mes))]
    serie = [(p, _numero(columnas[ORIGEN][f])) for f, p in meses]
    # Último tramo continuo con datos (> 0)
    con_dato = [i for i, (_, v) in enumerate(serie) if v and v > 0]
    if not con_dato:
        raise ValueError('La columna de trabajadores SRT no tiene datos')
    fin = con_dato[-1]
    ini = fin
    while ini - 1 >= 0 and serie[ini - 1][1] and serie[ini - 1][1] > 0:
        ini -= 1
    tramo = serie[ini:fin + 1]
    # Descuento de picos atípicos conocidos (solo para los factores)
    limpia = [v for _, v in tramo]
    aplicados = []
    pos = {p: i for i, (p, _) in enumerate(tramo)}
    fila_de = {p: f for f, p in meses}
    for a in atipicos:
        p = tuple(int(x) for x in a['periodo'].split('-'))
        if p not in pos or pos[p] in (0, len(tramo) - 1):
            continue
        col = columnas[a['columna']]
        ant = (p[0], p[1] - 1) if p[1] > 1 else (p[0] - 1, 12)
        sig = (p[0], p[1] + 1) if p[1] < 12 else (p[0] + 1, 1)
        v, va, vs = (_numero(col.get(fila_de.get(x))) for x in (p, ant, sig))
        if None in (v, va, vs):
            continue
        exceso = v - (va + vs) / 2
        limpia[pos[p]] -= exceso
        aplicados.append(f"{a['periodo']}: {round(exceso)} ({a.get('motivo', '')})")
    return tramo, limpia, aplicados


def calcular(tramo, limpia):
    factores = factores_estacionales(limpia)
    filas = []
    previo = None
    for (p, v), f in zip(tramo, factores):
        sa = v / f
        fila = {'periodo': p, 'original': v, 'factor': f, COL_NIVEL: round(sa)}
        if previo is not None:
            fila[COL_VAR] = round((sa / previo - 1) * 100, 2)
            fila[COL_DIF] = round(sa - previo)
        else:
            fila[COL_VAR] = ''
            fila[COL_DIF] = ''
        filas.append(fila)
        previo = sa
    return filas


NOTAS = {
    COL_NIVEL: ('Estimación DAEE, no es un dato oficial. Trabajadores registrados en Formosa (SRT) sin el efecto '
                'de la época del año. La SRT publica solo la serie original; esta columna la calcula la automatización '
                '(scripts/srt/desestacionalizar_srt.py) con el método de medias móviles X-11 simplificado, usando toda la '
                'serie desde enero 2021. Por eso es un número y no una fórmula: cada mes nuevo puede corregir un poco los '
                'meses anteriores. Se recalcula sola en cada revisión de la SRT. No editar a mano.'),
    COL_VAR: 'Fórmula: variación % contra el mes anterior de la serie desestacionalizada (estimación DAEE).',
    COL_DIF: 'Fórmula: diferencia en personas contra el mes anterior de la serie desestacionalizada (estimación DAEE).',
}


def poner_notas(cliente) -> None:
    """Nota explicativa en el encabezado de cada columna (idempotente)."""
    from sheet_directo import API, HOJA, Mapa
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    mapa = Mapa(encabezados, [])
    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(sheetId,title))'}, timeout=60)
    r.raise_for_status()
    hoja_id = next(s['properties']['sheetId'] for s in r.json()['sheets'] if s['properties']['title'] == HOJA)
    pedidos = []
    for titulo, nota in NOTAS.items():
        col, _ = mapa.columna(titulo)
        if col:
            pedidos.append({'updateCells': {'rows': [{'values': [{'note': nota}]}], 'fields': 'note',
                                            'range': {'sheetId': hoja_id, 'startRowIndex': 0, 'endRowIndex': 1,
                                                      'startColumnIndex': col - 1, 'endColumnIndex': col}}})
    if pedidos:
        cliente.sesion.post(f'{API}:batchUpdate', json={'requests': pedidos}, timeout=60).raise_for_status()


def escribir_formulas(cliente, filas) -> int:
    """Var % y Dif intermensual como fórmulas que miran la fila anterior (mismo estilo que las columnas SRT)."""
    from sheet_directo import HOJA, Mapa, es_formula, letra
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    periodos = [f[0] if f else None for f in cliente.leer('A:A')]
    mapa = Mapa(encabezados, periodos)
    c_niv, c_var, c_dif = (mapa.columna(t)[0] for t in COLUMNAS)
    if None in (c_niv, c_var, c_dif):
        raise ValueError('No se encontraron las columnas desestacionalizadas en el Sheet')
    N = letra(c_niv)
    actuales = {}
    for c in (c_var, c_dif):
        datos = cliente.leer(f'{letra(c)}:{letra(c)}', render='FORMULA')
        actuales[c] = {i + 1: (fila[0] if fila else None) for i, fila in enumerate(datos)}
    escribir = []
    for k, f in enumerate(filas):
        fila, _ = mapa.fila(f['periodo'])
        if fila is None:
            continue
        if k == 0:
            deseadas = {c_var: '', c_dif: ''}
        else:
            deseadas = {c_var: f'={N}{fila}/{N}{fila - 1}*100-100', c_dif: f'={N}{fila}-{N}{fila - 1}'}
        for c, formula in deseadas.items():
            actual = actuales[c].get(fila)
            if (actual or '') == formula:
                continue
            if formula == '' and actual in (None, ''):
                continue
            escribir.append({'range': f"'{HOJA}'!{letra(c)}{fila}", 'values': [[formula]]})
    if escribir:
        cliente.escribir(escribir)
    return len(escribir)


def asegurar_columnas(cliente) -> list[str]:
    """Agrega las columnas que falten, a continuación de la última columna SRT."""
    from sheet_directo import API, HOJA, Mapa, letra
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    mapa = Mapa(encabezados, [])
    faltan = [t for t in COLUMNAS if mapa.columna(t)[1] == 0]
    if not faltan:
        return []
    ocupadas = [i for i, t in enumerate(encabezados, start=1) if t not in (None, '')]
    ultima = max(ocupadas, default=1)
    srt = [i for i in ocupadas if str(encabezados[i - 1]).startswith('SRT ')]
    if not srt:
        raise ValueError('No se encontró el bloque SRT en el Sheet')
    despues_de = max(srt)
    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(sheetId,title,gridProperties))'}, timeout=60)
    r.raise_for_status()
    props = next(s['properties'] for s in r.json()['sheets'] if s['properties']['title'] == HOJA)
    if despues_de < ultima:
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


def main() -> int:
    ap = argparse.ArgumentParser(description='SRT Formosa: trabajadores desestacionalizados (estimación DAEE)')
    ap.add_argument('--solo-calcular', action='store_true', help='Calcula y muestra, sin escribir en el Sheet.')
    args = ap.parse_args()

    from sheet_directo import ClienteSheets, escribir_actualizaciones, texto_periodo
    cliente = ClienteSheets()
    tramo, limpia, aplicados = leer_series(cliente)
    filas = calcular(tramo, limpia)
    ultimo = filas[-1]
    resultado = {
        'fuente': 'SRT - Trabajadores desestacionalizados (estimación DAEE)',
        'periodo': f'{ultimo["periodo"][0]}-{ultimo["periodo"][1]:02d}',
        'desde': f'{filas[0]["periodo"][0]}-{filas[0]["periodo"][1]:02d}',
        'meses': len(filas),
        'atipicos_descontados': aplicados,
        'factores_ultimo_anio': {texto_periodo(f['periodo']): round((f['factor'] - 1) * 100, 1) for f in filas[-12:]},
        'ultimos_meses': [{'periodo': texto_periodo(f['periodo']), 'original': f['original'],
                           'desestacionalizado': f[COL_NIVEL], 'var_im': f[COL_VAR]} for f in filas[-6:]],
    }
    if args.solo_calcular:
        print(json.dumps(resultado, ensure_ascii=False, indent=1, default=str))
        print('ESTADO=OK')
        return 0

    agregadas = asegurar_columnas(cliente)
    poner_notas(cliente)
    # El nivel se escribe como número (lo calcula el método); Var % y Dif van como fórmulas.
    cargas = [{'fuente': FUENTE, 'hoja': 'BASE DE DATOS', 'periodo': f'{f["periodo"][0]}-{f["periodo"][1]:02d}-01',
               'valores': {COL_NIVEL: f[COL_NIVEL]}, 'permitir_reemplazos': True} for f in filas]
    escritura = escribir_actualizaciones(cargas, None, cliente)
    info = escritura.get('por_fuente', {}).get(FUENTE, {})
    formulas = escribir_formulas(cliente, filas)
    info['escritas'] = info.get('escritas', 0) + formulas
    info['esperadas'] = info.get('esperadas', 0) + formulas
    resultado['columnas_agregadas'] = agregadas
    resultado['celdas_actualizadas'] = info.get('escritas', 0)
    resultado['formulas_escritas'] = formulas
    print(json.dumps(resultado, ensure_ascii=False, indent=1, default=str))
    resumen = {k: info.get(k) for k in ('periodos', 'escritas', 'ya_estaban', 'esperadas', 'conservadas',
                                        'no_encontradas', 'errores')}
    resumen['periodos'] = (resumen.get('periodos') or [])[-1:]   # el aviso muestra solo el último mes
    print('SHEET_DIRECTO=' + json.dumps(resumen, ensure_ascii=False, default=str))
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
