"""Descripción + plantilla de WhatsApp de Comercialización de granos (página /productivo-comercializacion/).

Datos: bloque "COMERCIALIZACION" de la hoja "BASE DE DATOS - Indicadores Productivos" (columnas A a R):
por cultivo, precio interno ($/tn, Rosario), variación mensual (%), precio FOB (US$/tn) y variación FOB (%).
Orden de los bloques: Maíz (C-F), Sorgo (G-J), Soja (K-N), Trigo (O-R).
"""
from __future__ import annotations

import re
import unicodedata

WEB = 'https://observatorio.apiseostax.com.ar'
COMPLETAR = '[[completar]]'
MESES = ['', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre',
         'octubre', 'noviembre', 'diciembre']
CULTIVOS = [('maiz', 'El maíz', 2), ('sorgo', 'El sorgo', 6), ('soja', 'La soja', 10), ('trigo', 'El trigo', 14)]


class Falta(Exception):
    pass


def _norm(t) -> str:
    t = unicodedata.normalize('NFKD', str(t or ''))
    return ''.join(c for c in t if not unicodedata.combining(c)).strip().casefold()


def _num(v):
    if v in (None, ''):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace('$', '').replace(' ', '').replace('%', '')
    if re.fullmatch(r'-?\d{1,3}(\.\d{3})+(,\d+)?', s):
        s = s.replace('.', '').replace(',', '.')
    elif re.fullmatch(r'-?\d{1,3}(,\d{3})+(\.\d+)?', s):
        s = s.replace(',', '')
    else:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


class DatosComercializacion:
    """{(año, mes): {cultivo: (precio, var_mensual, fob, var_fob)}}"""

    def __init__(self, filas: list[list]):
        self.series = {}
        anio = None
        for f in filas:
            etiqueta = _norm(f[1] if len(f) > 1 else '') or _norm(f[0] if f else '')
            if re.fullmatch(r'20\d{2}', etiqueta):
                anio = int(etiqueta)
                continue
            if anio is None or etiqueta not in MESES:
                continue
            mes = MESES.index(etiqueta)
            valores = {}
            for clave, _, c in CULTIVOS:
                vals = [_num(f[c + k]) if c + k < len(f) else None for k in range(4)]
                if vals[0]:
                    valores[clave] = tuple(vals)
            if len(valores) == len(CULTIVOS):
                self.series[(anio, mes)] = valores

    def ultimo(self):
        return max(self.series) if self.series else None


def cargar() -> DatosComercializacion:
    from sheet_directo import ClienteSheets
    from sheet_productivos import _leer, _titulo_hoja
    cliente = ClienteSheets()
    return DatosComercializacion(_leer(cliente, _titulo_hoja(cliente), 'A1:R400', 'UNFORMATTED_VALUE'))


def _f(x, dec=0):
    if x is None:
        return COMPLETAR
    s = f'{x:,.{dec}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def _var(a, b):
    return None if a is None or not b else (a / b - 1) * 100


def describir(d: DatosComercializacion, p=None):
    p = p or d.ultimo()
    if not p or p not in d.series:
        raise Falta('faltan los precios del mes')
    ant = (p[0] - (p[1] == 1), (p[1] - 2) % 12 + 1)
    previo = d.series.get(ant, {})
    datos = []
    for clave, nombre, _ in CULTIVOS:
        precio, vm, fob, vf = d.series[p][clave]
        if previo.get(clave):  # se recalculan desde los precios (las columnas de variación pueden venir redondeadas)
            vm = _var(precio, previo[clave][0])
            vf = _var(fob, previo[clave][2]) if fob else vf
        datos.append((clave, nombre, precio, vm, fob, vf))
    if all(abs(x[3] or 0) < 0.005 for x in datos):
        raise Falta('los precios internos del mes son iguales a los del mes anterior en los cuatro cultivos; '
                    'revisar la carga antes de describir')
    orden = sorted(datos, key=lambda x: -(x[3] or 0))
    suben = [x for x in datos if (x[3] or 0) > 0]
    bajan = [x for x in datos if (x[3] or 0) < 0]
    mes = f'{MESES[p[1]]} de {p[0]}'

    def frase(x, primera=False):
        _, nombre, precio, vm, fob, vf = x
        verbo = ('una suba' if vm > 0 else 'una baja') if vm is not None else COMPLETAR
        fob_txt = (f'{"una mejora" if vf > 0 else "una caída"} del {_f(abs(vf), 1)}% en su precio de exportación, '
                   f'que {"alcanzó" if vf > 0 else "se ubicó en"} los {_f(fob)} US$/tn') if vf is not None else COMPLETAR
        return (f'{nombre} se ubicó en ${_f(precio)} por tonelada, con {verbo} mensual del {_f(abs(vm), 1)}%, '
                f'y {fob_txt}.')

    if suben and bajan:
        intro = f'En {mes} los precios del mercado interno mostraron comportamientos dispares entre cultivos. '
    elif suben:
        intro = f'En {mes} los precios del mercado interno subieron en los cuatro cultivos relevados. '
    else:
        intro = f'En {mes} los precios del mercado interno bajaron en los cuatro cultivos relevados. '
    texto = intro + ' '.join(frase(x) for x in orden)
    fob_max = max(datos, key=lambda x: x[4] or 0)
    opuestos = [x[1].split(' ', 1)[1] for x in datos if x[3] and x[5] and (x[3] > 0) != (x[5] > 0)]
    texto += (f' El precio FOB más alto fue el de {fob_max[1].split(" ", 1)[1]} ({_f(fob_max[4])} US$/tn).'
              + (f' En {", ".join(opuestos)} el precio interno y el de exportación se movieron en sentido opuesto.'
                 if opuestos else ' Los precios de exportación acompañaron la evolución del mercado interno.'))
    w = [('t', f'🌾 **Sector agrícola – Comercialización de granos | {MESES[p[1]].capitalize()} {p[0]}**'),
         ('p', f'Precios del mercado interno (Rosario) y de exportación (FOB) en {mes}:')]
    for _, nombre, precio, vm, fob, vf in orden:
        w.append(('b', f'{nombre.split(" ", 1)[1].capitalize()}: **${_f(precio)}/tn** ({"+" if (vm or 0) > 0 else ""}{_f(vm, 1)}% mensual) · '
                       f'FOB {_f(fob)} US$/tn ({"+" if (vf or 0) > 0 else ""}{_f(vf, 1)}%)'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/productivo-comercializacion/'))
    return f'Comercializacion_granos', w, [('Sector Agrícola - Comercialización', texto + f' {COMPLETAR}')], p
