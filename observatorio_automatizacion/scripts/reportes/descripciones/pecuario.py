"""Descripciones del Observatorio para el Sector Pecuario (Bovinos, Porcinos, Aves y Consumo).

Lee el bloque "Produccion Pecuaria" de la pestaña "BASE DE DATOS - Indicadores Productivos" del Sheet
y arma, con el mismo formato que el resto de las descripciones, la plantilla de WhatsApp y el texto
de la plataforma de cada página:

* Bovinos  -> https://observatorio.apiseostax.com.ar/bovinos/
* Porcinos -> https://observatorio.apiseostax.com.ar/porcinos/
* Aves     -> https://observatorio.apiseostax.com.ar/aves/
* Consumo  -> https://observatorio.apiseostax.com.ar/consumo-pecuario/  (solo cuando el mes tiene
  el consumo por habitante de las tres carnes: bovina, porcina y aviar)

La hoja mezcla formatos (1.045.293 / 4102.00 / 5092,00) y algunos valores quedaron cargados en miles
sin los ceros finales (p. ej. 628.25 por 628.250, o 187 por 187.000 t): se corrigen comparando con la
magnitud habitual de la columna. Lo que no se puede calcular queda marcado [[completar]].
"""
from __future__ import annotations

import re
import statistics
import unicodedata

WEB = 'https://observatorio.apiseostax.com.ar'
COMPLETAR = '[[completar]]'
MESES = ['', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre',
         'octubre', 'noviembre', 'diciembre']
MESES_CORTOS = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6, 'jul': 7, 'ago': 8,
                'sep': 9, 'set': 9, 'oct': 10, 'nov': 11, 'dic': 12}

# campo -> palabras que tiene que tener el encabezado (igual que actualizar_pecuaria.MAPEO)
CAMPOS = {
    'stock_bov_f': ('stock', 'fsa', 'bovinos'),
    'stock_bov_n': ('stock', 'nacion', 'bovinos'),
    'faena_bov_f': ('faena', 'fsa', 'bovinos'),
    'faena_bov_n': ('faena', 'nacion', 'bovinos'),
    'cons_bov': ('consumo', 'capita', 'bovino'),
    'stock_por_f': ('stock', 'fsa', 'porcinos'),
    'stock_por_n': ('stock', 'n', 'porcinos'),
    'faena_por_n': ('faena', 'total', 'porcinos'),
    'cons_por': ('consumo', 'capita', 'porcin'),
    'huevos': ('produccion', 'industria', 'huevos'),
    'carne_aviar': ('produccion', 'carne', 'tn'),
    'faena_aves': ('faena', 'miles', 'aves'),
    'cons_aves': ('consumo', 'capita', 'aves'),
}
# columnas que son "cabezas", toneladas, etc. (pueden venir sin los ceros de los miles)
ESCALABLES = {'stock_bov_f', 'stock_bov_n', 'faena_bov_n', 'stock_por_f', 'stock_por_n', 'faena_por_n',
              'huevos', 'carne_aviar', 'faena_aves'}


def _norm(t) -> str:
    t = unicodedata.normalize('NFKD', str(t or ''))
    t = ''.join(c for c in t if not unicodedata.combining(c)).casefold()
    return ' '.join(re.sub(r'[^a-z0-9]+', ' ', t).split())


def _num(valor):
    if valor in (None, ''):
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor) or None
    s = str(valor).strip().replace(' ', '').replace('$', '').replace('%', '')
    if not s or s in ('-', 'S/C', 's/c'):
        return None
    if re.fullmatch(r'-?\d{1,3}(\.\d{3})+', s):
        s = s.replace('.', '')
    elif ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        x = float(s)
    except ValueError:
        return None
    return x or None


def _periodo(valor):
    if valor in (None, ''):
        return None
    t = _norm(valor)
    m = re.match(r'^([a-z]+)\s+(\d{2}|\d{4})$', t)
    if m and m.group(1)[:3] in MESES_CORTOS:
        a = int(m.group(2))
        return (a + 2000 if a < 100 else a, MESES_CORTOS[m.group(1)[:3]])
    return None


def _mover(p, n):
    t = p[0] * 12 + p[1] - 1 + n
    return (t // 12, t % 12 + 1)


class DatosPecuarios:
    """Series mensuales por campo: {campo: {(año, mes): valor}}."""

    def __init__(self, filas: list[list]):
        fila_tit = next((i for i, f in enumerate(filas[:15]) if any(_norm(v) == 'periodo' for v in f)), None)
        if fila_tit is None:
            raise ValueError('no se encontró la columna Periodo en la hoja de productivos')
        encab = filas[fila_tit]
        col_p = next(i for i, v in enumerate(encab) if _norm(v) == 'periodo')
        cols = {}
        for campo, tokens in CAMPOS.items():
            hall = []
            for i, v in enumerate(encab):
                limpio = _norm(v)
                palabras = set(limpio.split())
                if 'nacion' in tokens or campo.endswith('_n'):
                    if 'fsa' in palabras:
                        continue
                if all((tok in palabras) if len(tok) == 1 else (tok in limpio) for tok in tokens):
                    hall.append(i)
            if len(hall) == 1:
                cols[campo] = hall[0]
        # períodos (las filas de septiembre a veces quedaron sin el mes escrito: se infieren por posición)
        periodos = [_periodo(f[col_p] if len(f) > col_p else None) for f in filas]
        for i in range(fila_tit + 1, len(filas)):
            if periodos[i] is None and i > 0 and periodos[i - 1] and i + 1 < len(filas) and periodos[i + 1] \
                    and _mover(periodos[i - 1], 2) == periodos[i + 1]:
                periodos[i] = _mover(periodos[i - 1], 1)
        self.series = {}
        for campo, c in cols.items():
            s = {}
            for i in range(fila_tit + 1, len(filas)):
                p = periodos[i]
                if not p or c >= len(filas[i]):
                    continue
                x = _num(filas[i][c])
                if x is not None:
                    s[p] = x
            if campo in ESCALABLES and s:
                med = statistics.median(s.values())
                for p, x in list(s.items()):
                    if med and x < med / 100 and 0.3 <= x * 1000 / med <= 3:
                        s[p] = x * 1000
                    elif campo == 'huevos' and med and x < med / 5 and 0.3 <= x * 10 / med <= 3:
                        s[p] = x * 10  # p. ej. 8115 por 81.150 (se perdió el último cero)
            self.series[campo] = s

    def v(self, campo, p):
        return self.series.get(campo, {}).get(p)

    def ultimo(self, campo):
        s = self.series.get(campo) or {}
        return max(s) if s else None

    def acumulado(self, campo, p):
        total = 0.0
        for m in range(1, p[1] + 1):
            x = self.v(campo, (p[0], m))
            if x is None:
                return None
            total += x
        return total

    def promedio(self, campo, p):
        a = self.acumulado(campo, p)
        return None if a is None else a / p[1]


def cargar() -> DatosPecuarios:
    from sheet_directo import ClienteSheets
    from sheet_productivos import _leer, _titulo_hoja
    cliente = ClienteSheets()
    titulo = _titulo_hoja(cliente)
    filas = _leer(cliente, titulo, 'A1:BZ400', 'FORMATTED_VALUE')
    return DatosPecuarios(filas)


# ----------------------------------------------------------------------------- textos
def _f(x, dec=0):
    if x is None:
        return COMPLETAR
    s = f'{x:,.{dec}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def _var(a, b):
    return None if a is None or not b else (a / b - 1) * 100


def _pct(x, dec=1):
    return COMPLETAR if x is None else _f(abs(x), dec) + '%'


def _signo(x, dec=1):
    return COMPLETAR if x is None else ('+' if x > 0 else '') + _f(x, dec) + '%'


def _mas_menos(x, mas='más', menos='menos'):
    return COMPLETAR if x is None else (mas if x >= 0 else menos)


def _mes(p):
    return MESES[p[1]]


def _titulo_mes(p):
    return f'{MESES[p[1]].capitalize()} {p[0]}'


class Falta(Exception):
    pass


def bovinos(d: DatosPecuarios, p):
    f = d.v('faena_bov_f', p)
    if f is None:
        raise Falta('falta la faena bovina de Formosa')
    ant, ia = _mover(p, -1), _mover(p, -12)
    f_im, f_ia = _var(f, d.v('faena_bov_f', ant)), _var(f, d.v('faena_bov_f', ia))
    acu, acu_a = d.acumulado('faena_bov_f', p), d.acumulado('faena_bov_f', ia)
    n = d.v('faena_bov_n', p)
    n_im, n_ia = _var(n, d.v('faena_bov_n', ant)), _var(n, d.v('faena_bov_n', ia))
    acun, acun_a = d.acumulado('faena_bov_n', p), d.acumulado('faena_bov_n', ia)
    part = (acu / acun * 100) if acu and acun else None
    ps = d.ultimo('stock_bov_f')
    stock = d.v('stock_bov_f', ps) if ps else None
    stock_n = d.v('stock_bov_n', ps) if ps else None
    stock_prev = d.v('stock_bov_f', _mover(ps, -12)) if ps else None
    part_s = (stock / stock_n * 100) if stock and stock_n else None
    w = [('t', f'🐄 **Sector pecuario – Bovinos | Formosa | {_titulo_mes(p)}**'),
         ('p', f'En {_mes(p)} de {p[0]} la faena bovina en Formosa fue de **{_f(f)} cabezas**, '
               f'un **{_signo(f_im)} intermensual** y un **{_signo(f_ia)} interanual**.')]
    w.append(('b', f'Acumulado enero-{_mes(p)}: **{_f(acu)} cabezas** ({_signo(_var(acu, acu_a))} frente a igual período de {p[0] - 1}).'))
    w.append(('b', f'🇦🇷 A nivel país: **{_f(n)} cabezas** ({_signo(n_im)} intermensual y {_signo(n_ia)} interanual).'))
    if stock:
        w.append(('b', f'Stock bovino provincial: **{_f(stock)} cabezas** (diciembre {ps[0]}), {_f(part_s, 2)}% del rodeo nacional.'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/bovinos/'))
    texto = ''
    if stock:
        texto += (f'El stock bovino de Formosa fue de {_f(stock)} cabezas en diciembre de {ps[0]}'
                  + (f', un {_pct(_var(stock, stock_prev))} {_mas_menos(_var(stock, stock_prev), "más", "menos")} que un año antes' if stock_prev else '')
                  + (f', con una participación del {_f(part_s, 2)}% en el rodeo nacional ({_f(stock_n)} cabezas)' if stock_n else '')
                  + '. ')
    texto += (f'En {_mes(p)} de {p[0]} la faena provincial alcanzó {_f(f)} cabezas, un {_pct(f_im)} {_mas_menos(f_im)} que en '
              f'{_mes(ant)} y un {_pct(f_ia)} {_mas_menos(f_ia, "por encima", "por debajo")} de {_mes(p)} de {p[0] - 1}. '
              f'En el acumulado enero-{_mes(p)} se faenaron {_f(acu)} bovinos frente a {_f(acu_a)} en igual período del año '
              f'anterior, una variación interanual del {_signo(_var(acu, acu_a))}. '
              f'A nivel nacional la faena de {_mes(p)} fue de {_f(n)} cabezas, un {_pct(n_im)} {_mas_menos(n_im)} que en '
              f'{_mes(ant)} y un {_pct(n_ia)} {_mas_menos(n_ia)} que un año atrás, con un acumulado de {_f(acun)} cabezas '
              f'({_signo(_var(acun, acun_a))} interanual). La participación de Formosa en la faena nacional acumulada fue '
              f'del {_f(part, 2)}%. {COMPLETAR}')
    return 'Bovinos', w, [('Sector Pecuario - Bovinos', texto)]


def porcinos(d: DatosPecuarios, p):
    n = d.v('faena_por_n', p)
    if n is None:
        raise Falta('falta la faena porcina nacional')
    ant, ia = _mover(p, -1), _mover(p, -12)
    n_im, n_ia = _var(n, d.v('faena_por_n', ant)), _var(n, d.v('faena_por_n', ia))
    acu, acu_a = d.acumulado('faena_por_n', p), d.acumulado('faena_por_n', ia)
    ps = d.ultimo('stock_por_f')
    sf = d.v('stock_por_f', ps) if ps else None
    sn = d.v('stock_por_n', ps) if ps else None
    sf_a = d.v('stock_por_f', _mover(ps, -12)) if ps else None
    sn_a = d.v('stock_por_n', _mover(ps, -12)) if ps else None
    part = sf / sn * 100 if sf and sn else None
    w = [('t', f'🐖 **Sector pecuario – Porcinos | {_titulo_mes(p)}**'),
         ('p', f'En {_mes(p)} de {p[0]} la faena porcina nacional fue de **{_f(n)} cabezas**, un **{_signo(n_im)} intermensual** '
               f'y un **{_signo(n_ia)} interanual**.'),
         ('b', f'Acumulado enero-{_mes(p)}: **{_f(acu)} cabezas** ({_signo(_var(acu, acu_a))} frente a igual período de {p[0] - 1}).')]
    if sf:
        w.append(('b', f'Stock porcino de Formosa: **{_f(sf)} cabezas** ({_mes(ps)} {ps[0]}), {_f(part, 2)}% del total nacional.'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/porcinos/'))
    texto = ''
    if sf:
        texto += (f'El stock porcino de Formosa fue de {_f(sf)} cabezas en {_mes(ps)} de {ps[0]}'
                  + (f', un {_pct(_var(sf, sf_a))} {_mas_menos(_var(sf, sf_a))} que las {_f(sf_a)} cabezas de {_mes(ps)} de {ps[0] - 1}' if sf_a else '')
                  + '. '
                  + (f'A nivel nacional el stock alcanzó {_f(sn)} cabezas'
                     + (f', un {_pct(_var(sn, sn_a))} {_mas_menos(_var(sn, sn_a))} que un año antes' if sn_a else '')
                     + f', por lo que la participación de Formosa fue del {_f(part, 2)}%. ' if sn else ''))
    texto += (f'Entre enero y {_mes(p)} de {p[0]} la faena porcina nacional alcanzó {_f(acu)} cabezas, un {_pct(_var(acu, acu_a))} '
              f'{_mas_menos(_var(acu, acu_a), "por encima", "por debajo")} de igual período de {p[0] - 1}, y en {_mes(p)} se faenaron '
              f'{_f(n)} cabezas, un {_pct(n_im)} {_mas_menos(n_im)} que en {_mes(ant)} y un {_pct(n_ia)} {_mas_menos(n_ia)} que en '
              f'{_mes(p)} de {p[0] - 1}. No hay información de faena porcina provincial. {COMPLETAR}')
    return 'Porcinos', w, [('Sector Pecuario - Porcinos', texto)]


def aves(d: DatosPecuarios, p):
    fa = d.v('faena_aves', p)
    if fa is None:
        raise Falta('falta la faena aviar')
    ant, ia = _mover(p, -1), _mover(p, -12)
    fa_im, fa_ia = _var(fa, d.v('faena_aves', ant)), _var(fa, d.v('faena_aves', ia))
    acu, acu_a = d.acumulado('faena_aves', p), d.acumulado('faena_aves', ia)
    ca = d.v('carne_aviar', p)
    ca_im, ca_ia = _var(ca, d.v('carne_aviar', ant)), _var(ca, d.v('carne_aviar', ia))
    cacu, cacu_a = d.acumulado('carne_aviar', p), d.acumulado('carne_aviar', ia)
    ph = d.ultimo('huevos')
    hu = d.v('huevos', ph) if ph else None
    hu_im = _var(hu, d.v('huevos', _mover(ph, -1))) if ph else None
    hu_ia = _var(hu, d.v('huevos', _mover(ph, -12))) if ph else None
    hacu = d.acumulado('huevos', ph) if ph else None
    hacu_a = d.acumulado('huevos', _mover(ph, -12)) if ph else None
    mill = lambda x: None if x is None else x / 1000  # noqa: E731  (miles de aves -> millones)
    milt = lambda x: None if x is None else x / 1000  # noqa: E731  (toneladas -> miles de t)
    w = [('t', f'🐔 **Sector pecuario – Aves | {_titulo_mes(p)}**'),
         ('p', f'En {_mes(p)} de {p[0]} se faenaron **{_f(mill(fa), 3)} millones de aves** a nivel nacional, un '
               f'**{_signo(fa_im)} intermensual** y un **{_signo(fa_ia)} interanual**.')]
    if ca:
        w.append(('b', f'Producción de carne aviar: **{_f(milt(ca))} mil toneladas** ({_signo(ca_im)} intermensual).'))
    if hu:
        w.append(('b', f'Huevos destinados a industria ({_mes(ph)}): **{_f(hu)} miles** ({_signo(hu_im)} intermensual y {_signo(hu_ia)} interanual).'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/aves/'))
    texto = ''
    if cacu:
        texto += (f'Entre enero y {_mes(p)} de {p[0]} la producción nacional de carne aviar alcanzó {_f(milt(cacu))} mil toneladas, '
                  f'frente a {_f(milt(cacu_a))} mil toneladas en igual período de {p[0] - 1} ({_signo(_var(cacu, cacu_a))}). ')
    texto += (f'La faena acumuló {_f(mill(acu), 3)} millones de aves, un {_pct(_var(acu, acu_a))} '
              f'{_mas_menos(_var(acu, acu_a))} que los {_f(mill(acu_a), 3)} millones del mismo período del año anterior. '
              f'En {_mes(p)}' + (f' la producción fue de {_f(milt(ca))} mil toneladas, un {_pct(ca_im)} {_mas_menos(ca_im, "superior", "inferior")} a {_mes(ant)}, mientras que' if ca else '')
              + f' la faena {"aumentó" if (fa_im or 0) >= 0 else "disminuyó"} {_pct(fa_im)} intermensual, al pasar de '
              f'{_f(mill(d.v("faena_aves", ant)), 3)} a {_f(mill(fa), 3)} millones de aves. ')
    if hu:
        texto += (f'En {_mes(ph)} se registraron {_f(hu)} miles de huevos destinados a industria, un {_pct(hu_im)} '
                  f'{_mas_menos(hu_im)} que en {_mes(_mover(ph, -1))} y un {_pct(hu_ia)} {_mas_menos(hu_ia)} que en {_mes(ph)} de '
                  f'{ph[0] - 1}; en el acumulado enero-{_mes(ph)} sumaron {_f(hacu)} miles ({_signo(_var(hacu, hacu_a))} interanual). ')
    texto += COMPLETAR
    return 'Aves', w, [('Sector Pecuario - Aves', texto)]


def consumo(d: DatosPecuarios, p):
    b, po, a = d.v('cons_bov', p), d.v('cons_por', p), d.v('cons_aves', p)
    if None in (b, po, a):
        raise Falta('falta el consumo por habitante de alguna de las tres carnes')
    ia, ant = _mover(p, -12), _mover(p, -1)
    pr = {k: (d.promedio(k, p), d.promedio(k, ia)) for k in ('cons_bov', 'cons_por', 'cons_aves')}
    tot = sum(x for x in (b, po, a))
    tot_pr = [sum(pr[k][i] for k in pr) if all(pr[k][i] for k in pr) else None for i in (0, 1)]
    def linea(nombre, k, x):
        return (f'El consumo {nombre} promedió {_f(pr[k][0], 2)} kg/hab/año frente a {_f(pr[k][1], 2)} kg/hab/año un año antes '
                f'({_signo(_var(*pr[k]))}); en {_mes(p)} alcanzó {_f(x, 2)} kg/hab/año, un {_pct(_var(x, d.v(k, ant)))} '
                f'{_mas_menos(_var(x, d.v(k, ant)))} que en {_mes(ant)} y un {_pct(_var(x, d.v(k, ia)))} '
                f'{_mas_menos(_var(x, d.v(k, ia)))} que en {_mes(p)} de {p[0] - 1}.')
    w = [('t', f'🥩 **Sector pecuario – Consumo de carnes | {_titulo_mes(p)}**'),
         ('p', f'En {_mes(p)} de {p[0]} el consumo por habitante fue de **{_f(b, 2)} kg/hab/año de carne bovina**, '
               f'**{_f(a, 2)} de carne aviar** y **{_f(po, 2)} de carne porcina** ({_f(tot, 2)} kg/hab/año en total).'),
         ('b', f'Promedio enero-{_mes(p)}: {_signo(_var(*tot_pr))} frente a igual período de {p[0] - 1} en el total de carnes.'),
         ('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/consumo-pecuario/')]
    texto = (f'El consumo total promedio per cápita de carnes registró una variación del {_signo(_var(*tot_pr))} entre '
             f'enero-{_mes(p)} de {p[0]} y el mismo período de {p[0] - 1}. ' + linea('bovino', 'cons_bov', b) + ' '
             + linea('porcino', 'cons_por', po) + ' ' + linea('aviar', 'cons_aves', a) + ' '
             + f'En niveles, en {_mes(p)} la carne bovina se ubicó en {_f(b, 2)}, la aviar en {_f(a, 2)} y la porcina en '
               f'{_f(po, 2)} kg/hab/año. {COMPLETAR}')
    return 'Consumo_pecuario', w, [('Sector Pecuario - Consumo', texto)]


INDICADORES = {'bovinos': bovinos, 'porcinos': porcinos, 'aves': aves, 'consumo_pecuario': consumo}
NOMBRES = {'bovinos': 'Bovinos', 'porcinos': 'Porcinos', 'aves': 'Aves', 'consumo_pecuario': 'Consumo de carnes'}
# palabras del encabezado de la columna cambiada -> página
ESPECIE = (('bovino', 'bovinos'), ('porcin', 'porcinos'), ('aves', 'aves'), ('huevo', 'aves'))
PRINCIPAL = {'bovinos': 'faena_bov_f', 'porcinos': 'faena_por_n', 'aves': 'faena_aves'}


def pedidos_desde_cambios(cambios: list[dict], d: DatosPecuarios) -> dict:
    """Qué páginas tienen datos NUEVOS (celda que estaba vacía) y para qué mes.
    Consumo: solo si el mes ya tiene el consumo de las tres carnes y alguno entró ahora."""
    from sheet_directo import vacio_o_cero
    nuevos = [c for c in cambios if vacio_o_cero(c.get('antes')) and not c.get('mismo_dato')]
    especies, consumo_nuevo = set(), False
    for c in nuevos:
        t = _norm(c.get('titulo'))
        for clave, esp in ESPECIE:
            if clave in t:
                especies.add(esp)
        if 'consumo' in t and 'capita' in t:
            consumo_nuevo = True
    pedidos = {}
    for esp in especies:
        p = d.ultimo(PRINCIPAL[esp])
        if p:
            pedidos[esp] = p
    if consumo_nuevo:
        comunes = [p for p in d.series.get('cons_bov', {}) if d.v('cons_por', p) and d.v('cons_aves', p)]
        if comunes:
            pedidos['consumo_pecuario'] = max(comunes)
    return pedidos
