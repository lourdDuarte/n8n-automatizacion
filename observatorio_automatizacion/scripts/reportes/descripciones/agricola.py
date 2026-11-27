"""Descripción + plantilla de WhatsApp de Producción agrícola (Estimaciones MAGyP), UNA POR CULTIVO.

La página del Observatorio (/productivo-produccion/) tiene un texto con todos los cultivos juntos;
cuando entran datos nuevos de un cultivo (p. ej. soja 2025/26) se arma el Word solo con ese cultivo:
superficie sembrada y cosechada, producción, rendimiento, comparación con la campaña anterior y
participación en el total nacional. Datos: bloque "Producción" (Cultivo / Campaña) de la hoja
"BASE DE DATOS - Indicadores Productivos".
"""
from __future__ import annotations

import re
import unicodedata

WEB = 'https://observatorio.apiseostax.com.ar'
COMPLETAR = '[[completar]]'
# orden de las columnas después de Cultivo y Campaña (igual que estimaciones.TITULOS)
CAMPOS = ['sem_f', 'sem_n', 'cos_f', 'cos_n', 'prod_f', 'prod_n']
NOMBRE = {'soja': 'Soja', 'trigo': 'Trigo', 'maiz': 'Maíz', 'arroz': 'Arroz', 'algodon': 'Algodón',
          'girasol': 'Girasol', 'sorgo': 'Sorgo', 'cartamo': 'Cártamo', 'garbanzo': 'Garbanzo'}
EMOJI = {'soja': '🌱', 'trigo': '🌾', 'maiz': '🌽', 'arroz': '🍚', 'algodon': '☁️', 'girasol': '🌻'}
SUBTIPOS = {'soja': [('soja 1ra', 'de primera'), ('soja 2da', 'de segunda')],
            'trigo': [('trigo candeal', 'candeal')]}


class Falta(Exception):
    pass


def _norm(t) -> str:
    t = unicodedata.normalize('NFKD', str(t or ''))
    t = ''.join(c for c in t if not unicodedata.combining(c)).casefold()
    return ' '.join(re.sub(r'[^a-z0-9/]+', ' ', t).split())


def _num(valor):
    if valor in (None, ''):
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    s = str(valor).strip().replace(' ', '').replace('%', '')
    if re.fullmatch(r'-?\d{1,3}(\.\d{3})+', s):
        s = s.replace('.', '')
    elif ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def grupo(cultivo) -> str:
    """'Soja 1ra' -> 'soja', 'Trigo total' -> 'trigo', 'Maíz' -> 'maiz'."""
    t = _norm(cultivo)
    return t.split()[0] if t else ''


class DatosAgricolas:
    """{(cultivo normalizado, campaña): {campo: valor}}"""

    def __init__(self, filas: list[list]):
        enc = next(((i, f) for i, f in enumerate(filas[:15])
                    if 'cultivo' in [_norm(v) for v in f] and 'campana' in [_norm(v) for v in f]), None)
        if enc is None:
            raise ValueError('no se encontró el encabezado Cultivo / Campaña en la hoja de productivos')
        i0, f0 = enc
        n = [_norm(v) for v in f0]
        c_cult, c_camp = n.index('cultivo'), n.index('campana')
        self.series, self.nombres = {}, {}
        for f in filas[i0 + 1:]:
            cu = f[c_cult] if c_cult < len(f) else ''
            ca = str(f[c_camp] if c_camp < len(f) else '').strip()
            if _norm(cu) == 'cultivo':
                break  # otro cuadro debajo
            if not cu or not re.fullmatch(r'\d{4}/\d{2}', ca):
                continue
            vals = {}
            for k, campo in enumerate(CAMPOS):
                c = c_camp + 1 + k
                vals[campo] = _num(f[c]) if c < len(f) else None
            self.series[(_norm(cu), ca)] = vals
            self.nombres[_norm(cu)] = str(cu).strip()

    def principal(self, g: str) -> str | None:
        """Fila que representa al cultivo: 'soja total' si existe, si no la única del grupo."""
        cands = sorted({c for c, _ in self.series if grupo(c) == g})
        if f'{g} total' in cands:
            return f'{g} total'
        sin_sub = [c for c in cands if c == g]
        return sin_sub[0] if sin_sub else (cands[0] if len(cands) == 1 else None)

    def v(self, cultivo, campana):
        return self.series.get((cultivo, campana))

    def campanas(self, cultivo):
        return sorted(ca for c, ca in self.series if c == cultivo)


def cargar() -> DatosAgricolas:
    from sheet_directo import ClienteSheets
    from sheet_productivos import _leer, _titulo_hoja
    cliente = ClienteSheets()
    titulo = _titulo_hoja(cliente)
    return DatosAgricolas(_leer(cliente, titulo, 'A1:BZ400', 'UNFORMATTED_VALUE'))


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


def _sube(x, si='creció', no='cayó'):
    return COMPLETAR if x is None else (si if x >= 0 else no)


def _part(v):
    return v['prod_f'] / v['prod_n'] * 100 if v and v.get('prod_f') and v.get('prod_n') else None


def _rinde(v):
    return v['prod_f'] / v['cos_f'] if v and v.get('prod_f') and v.get('cos_f') else None


def _millones(x):
    return COMPLETAR if x is None else _f(x / 1e6, 2) + ' millones de t'


def describir(d: DatosAgricolas, g: str, campana: str):
    cult = d.principal(g)
    v = d.v(cult, campana) if cult else None
    if not v or v.get('prod_f') is None:
        raise Falta(f'faltan los datos de {NOMBRE.get(g, g)} {campana}')
    nombre = NOMBRE.get(g, d.nombres.get(cult, g).capitalize())
    previas = [c for c in d.campanas(cult) if c < campana]
    ant = previas[-1] if previas else None
    va = d.v(cult, ant) if ant else None
    vp, vs = _var(v['prod_f'], va and va['prod_f']), _var(v['sem_f'], va and va['sem_f'])
    part, part_a = _part(v), _part(va)
    rinde, rinde_a = _rinde(v), _rinde(va)
    subs = [(etq, d.v(c, campana)) for c, etq in SUBTIPOS.get(g, []) if d.v(c, campana)]
    emoji = EMOJI.get(g, '🌾')

    w = [('t', f'{emoji} **Sector agrícola – Producción de {nombre.lower()} | Formosa | Campaña {campana}**'),
         ('p', f'En la campaña {campana} la producción de {nombre.lower()} en Formosa fue de **{_f(v["prod_f"])} toneladas**'
               + (f', un **{_signo(vp)}** frente a la campaña {ant}' if ant else '') + '.')]
    w.append(('b', f'Superficie sembrada: **{_f(v["sem_f"])} ha**' + (f' ({_signo(vs)})' if ant else '')
                   + f'; cosechada: **{_f(v["cos_f"])} ha**.'))
    if rinde:
        w.append(('b', f'Rendimiento: **{_f(rinde, 2)} t/ha**' + (f' (en {ant}: {_f(rinde_a, 2)} t/ha)' if rinde_a else '') + '.'))
    for etq, x in subs:
        w.append(('b', f'{nombre} {etq}: {_f(x["prod_f"])} t en {_f(x["sem_f"])} ha sembradas.'))
    w.append(('b', f'🇦🇷 Participación en la producción nacional: **{_f(part, 2)}%** ({_millones(v["prod_n"])} en el país)'
                   + (f'; en {ant} fue del {_f(part_a, 2)}%' if part_a else '') + '.'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/productivo-produccion/'))

    t = (f'En la campaña {campana}, {"la" if nombre.lower() not in ("maíz", "arroz", "algodón", "girasol", "sorgo", "trigo", "cártamo", "garbanzo") else "el"} '
         f'{nombre.lower()} en Formosa registró una superficie sembrada de {_f(v["sem_f"])} hectáreas y una cosechada de '
         f'{_f(v["cos_f"])} hectáreas, con una producción de {_f(v["prod_f"])} toneladas')
    if ant:
        t += (f'. Frente a la campaña {ant} ({_f(va["prod_f"])} toneladas sobre {_f(va["sem_f"])} hectáreas sembradas), '
              f'la producción {_sube(vp)} un {_pct(vp)} y la superficie sembrada '
              f'{_sube(vs, "aumentó", "disminuyó")} un {_pct(vs)}')
    t += '.'
    if rinde:
        t += f' El rendimiento fue de {_f(rinde, 2)} toneladas por hectárea cosechada' + \
             (f' ({_f(rinde_a, 2)} t/ha en {ant})' if rinde_a else '') + '.'
    if subs:
        t += ' Por tipo, ' + ' y '.join(f'la {nombre.lower()} {etq} aportó {_f(x["prod_f"])} toneladas '
                                         f'({_f(x["sem_f"])} ha sembradas)' for etq, x in subs) + '.'
    t += (f' A nivel nacional se produjeron {_millones(v["prod_n"])}, por lo que la participación de Formosa fue del '
          f'{_f(part, 2)}%' + (f', frente al {_f(part_a, 2)}% de la campaña anterior' if part_a else '') + f'. {COMPLETAR}')
    return f'Produccion_{nombre}_{campana.replace("/", "-")}', w, [(f'Sector Agrícola - Producción ({nombre})', t)]


def pedidos_desde_cambios(cambios: list[dict], d: DatosAgricolas) -> dict:
    """{grupo: campaña} de los cultivos con datos NUEVOS (celda que estaba vacía) en esta revisión."""
    pedidos = {}
    for c in cambios:
        if c.get('antes') not in (None, '') or c.get('formula'):
            continue
        m = re.match(r'^(.*\S)\s+(\d{4}/\d{2})$', str(c.get('titulo') or '').strip())
        if not m:
            continue
        g, campana = grupo(m.group(1)), m.group(2)
        if g and campana > pedidos.get(g, ''):
            pedidos[g] = campana
    return {g: ca for g, ca in pedidos.items() if d.principal(g)}
