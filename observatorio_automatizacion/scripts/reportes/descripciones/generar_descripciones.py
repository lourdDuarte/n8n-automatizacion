"""Descripciones para el Observatorio Económico (observatorio.apiseostax.com.ar).

Cada vez que una revisión carga datos nuevos de un indicador que está publicado en el
Observatorio, arma un Word con:
  1) la plantilla para WhatsApp (con emojis y negritas, lista para copiar), y
  2) la "Descripción plataforma": el texto del recuadro de cada página del Observatorio.

Los números salen de la hoja BASE DE DATOS del Sheet (la misma que alimenta la plataforma).
Lo que no está en la base (p. ej. comparaciones con el NEA o un análisis propio) queda marcado
en amarillo como [[completar]] para escribirlo a mano antes de publicar.

El Word se guarda en automatizacion_indicadores/descripciones/AAAA-MM/ y se manda como adjunto
al grupo de Telegram.

Uso
    python generar_descripciones.py --resumen salidas_n8n/resumen_general.json   (lo usan las revisiones)
    python generar_descripciones.py --indicador supermercados                     (último período cargado)
    python generar_descripciones.py --indicador emae --periodo 2026-08 --sin-telegram
    python generar_descripciones.py --listar
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import unicodedata
import urllib.request
import uuid
from datetime import datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
sys.path.insert(0, str(AQUI))

from word_simple import Documento  # noqa: E402

CARPETA = RAIZ / 'descripciones'
BASE_LOCAL = RAIZ / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx'
CONFIG_TELEGRAM = RAIZ / 'config' / 'telegram.env'
CHAT_GRUPO = '-5396908101'
WEB = 'https://observatorio.apiseostax.com.ar'
MESES = ['', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre',
         'octubre', 'noviembre', 'diciembre']
COMPLETAR = '[[completar]]'


# ------------------------------------------------------------------------------- formato
def num(x, dec=0) -> str:
    texto = f'{abs(x):,.{dec}f}'.replace(',', '§').replace('.', ',').replace('§', '.')
    return ('-' if x < 0 else '') + texto


def pct(x, dec=1) -> str:
    return num(abs(x), dec) + '%'


def con_signo(x, dec=1) -> str:
    return ('+' if x > 0 else '') + num(x, dec) + '%'


def mes(p) -> str:
    return MESES[p[1]]


def mes_anio(p) -> str:
    return f'{MESES[p[1]]} de {p[0]}'


def titulo_mes(p) -> str:
    return f'{MESES[p[1]].capitalize()} {p[0]}'


def mover(p, meses: int):
    total = p[0] * 12 + (p[1] - 1) + meses
    return (total // 12, total % 12 + 1)


def var(a, b):
    if a is None or b in (None, 0):
        return None
    return (a / b - 1) * 100


def subio(x, si='aumentaron', no='disminuyeron', igual='se mantuvieron sin cambios'):
    if x is None:
        return COMPLETAR
    return si if x > 0.05 else no if x < -0.05 else igual


def norm(texto) -> str:
    s = unicodedata.normalize('NFD', str(texto or ''))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower().replace('*', ' ')
    return ' '.join(s.split())


# ------------------------------------------------------------------------------- datos
class Datos:
    """Hoja BASE DE DATOS: busca columnas por título y valores por período (vacío o 0 = sin dato)."""

    def __init__(self, encabezados: list, filas: list):
        from sheet_directo import periodo_de
        self.titulos = [norm(t) for t in encabezados]
        self.originales = [str(t or '').strip() for t in encabezados]
        self.filas = {}
        for fila in filas:
            p = periodo_de(fila[0] if fila else None)
            if p and p not in self.filas:
                self.filas[p] = fila

    def col(self, titulo: str) -> int:
        buscado = norm(titulo)
        for i, t in enumerate(self.titulos):
            if t == buscado:
                return i
        for i, t in enumerate(self.titulos):
            if t.startswith(buscado):
                return i
        raise KeyError(f'No está la columna "{titulo}" en BASE DE DATOS')

    def columnas(self, prefijo: str) -> list[tuple[int, str]]:
        buscado = norm(prefijo)
        return [(i, self.originales[i]) for i, t in enumerate(self.titulos) if t.startswith(buscado)]

    def v(self, titulo_o_indice, p):
        i = titulo_o_indice if isinstance(titulo_o_indice, int) else self.col(titulo_o_indice)
        fila = self.filas.get(p)
        if not fila or i >= len(fila):
            return None
        valor = fila[i]
        if isinstance(valor, bool):
            return None
        if isinstance(valor, str):
            texto = valor.strip().replace('%', '').replace('$', '').replace(' ', '')
            if re.fullmatch(r'-?[\d.]+,\d+', texto):
                texto = texto.replace('.', '').replace(',', '.')
            try:
                valor = float(texto)
            except ValueError:
                return None
        if not isinstance(valor, (int, float)) or valor == 0:
            return None
        return float(valor)

    def ultimo(self, titulo):
        con_dato = [p for p in self.filas if self.v(titulo, p) is not None]
        return max(con_dato) if con_dato else None

    def acumulado(self, titulo, p):
        """Suma enero..p del año de p (None si falta algún mes)."""
        total = 0.0
        for m in range(1, p[1] + 1):
            x = self.v(titulo, (p[0], m))
            if x is None:
                return None
            total += x
        return total


def cargar_datos() -> Datos:
    """Primero el Sheet (valores calculados); si no hay conexión, la copia local del Excel."""
    try:
        from sheet_directo import ClienteSheets
        valores = ClienteSheets().leer('A1:BZZ2000', render='UNFORMATTED_VALUE')
        if valores:
            return Datos(valores[0], valores[1:])
    except Exception as error:  # noqa: BLE001
        print(f'AVISO=No se pudo leer el Sheet ({error}); se usa la copia local.', flush=True)
    import openpyxl
    libro = openpyxl.load_workbook(BASE_LOCAL, read_only=True, data_only=True)
    filas = [list(f) for f in libro['BASE DE DATOS'].iter_rows(values_only=True)]
    libro.close()
    return Datos(filas[0], filas[1:])


class Falta(Exception):
    """El dato principal del período todavía no está en la base."""


def requerido(x, que):
    if x is None:
        raise Falta(f'falta {que}')
    return x


# ------------------------------------------------------------------------------- indicadores
# Cada función devuelve: (titulo, bloques_whatsapp, [(subtitulo, texto_plataforma), ...])
# bloques: ('t', texto) título · ('h', texto) subtítulo · ('p', texto) párrafo · ('b', texto) viñeta

def supermercados(d: Datos, p):
    C = 'Ventas totales de supermercados (en millones de pesos) F'
    K = 'Ventas totales de supermercados (a precios constantes y en millones de pesos) F'
    KN = 'Ventas totales de supermercados (a precios constantes y en millones de pesos) N'
    CN = 'Ventas totales de supermercados (en millones de pesos) N'
    OP, OPN = 'Cantidad de operaciones realizadas en supermercados F', 'Cantidad de operaciones realizadas en supermercados N'
    ant, ia, a23 = mover(p, -1), mover(p, -12), (2023, p[1])
    c, c_ant = requerido(d.v(C, p), 'ventas a precios corrientes'), d.v(C, ant)
    k, k_ant = requerido(d.v(K, p), 'ventas a precios constantes'), d.v(K, ant)
    c_im, c_ia = var(c, c_ant), var(c, d.v(C, ia))
    k_im, k_ia, k_23 = var(k, k_ant), var(k, d.v(K, ia)), var(k, d.v(K, a23))
    ops, ops_im, ops_ia = d.v(OP, p), var(d.v(OP, p), d.v(OP, ant)), var(d.v(OP, p), d.v(OP, ia))
    ticket = c * 1e6 / ops if ops else None
    ticket_n = d.v(CN, p) * 1e6 / d.v(OPN, p) if d.v(CN, p) and d.v(OPN, p) else None
    kn_ia, kn_im = var(d.v(KN, p), d.v(KN, ia)), var(d.v(KN, p), d.v(KN, ant))
    t = f'📊 **Ventas en Supermercados – Formosa | {titulo_mes(p)}**'
    w = [('t', t), ('h', '🛒 **Ventas a precios constantes**'),
         ('p', f'En {mes_anio(p)}, las ventas en supermercados de Formosa a precios constantes (sin el efecto de la '
               f'inflación) fueron de **${num(k, 1)} millones**, con una variación del **{con_signo(k_im)} intermensual** '
               f'y del **{con_signo(k_ia)} interanual**.'),
         ('h', '📅 **Comparación con años anteriores**')]
    if k_23 is not None and p[0] > 2023:
        w.append(('b', f'Frente a {mes_anio(a23)}: variación del **{con_signo(k_23)}** a precios constantes.'))
    w.append(('b', f'💰 **Ventas a precios corrientes:** superaron los **${num(int(c))} millones**, con una variación del '
                   f'**{con_signo(c_im)} intermensual** y del **{con_signo(c_ia)} interanual**.'))
    if ops:
        w += [('h', '🧾 **Operaciones y ticket promedio**'),
              ('p', f'Se realizaron **{num(ops)} operaciones** ({con_signo(ops_im)} intermensual y {con_signo(ops_ia)} '
                    f'interanual), con un ticket promedio de **${num(ticket)}**'
                    + (f' (promedio nacional: ${num(ticket_n)}).' if ticket_n else '.'))]
    if kn_ia is not None:
        w.append(('p', f'🇦🇷 **A nivel nacional:** las ventas a precios constantes {subio(kn_ia, "crecieron", "cayeron", "no variaron")} '
                       f'{pct(kn_ia)} interanual y ' + (f'{pct(kn_im)} ' if (kn_im or 0) * (kn_ia or 0) > 0 else f'{subio(kn_im, "crecieron", "cayeron", "no variaron")} {pct(kn_im)} ')
                       + f'respecto a {mes(ant)}.'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/precio-constante/'))
    constante = (f'Considerando que la Encuesta de Supermercados relevada por el INDEC, es con base diciembre 2016=100, al '
                 f'analizar los guarismos a precios constantes (sin el efecto de la inflación), se observa que las ventas '
                 f'totales en supermercados de Formosa en el mes de {mes(p)} fueron de ${num(k, 1)} millones (en {mes(ant)} '
                 f'de $ {num(k_ant, 1) if k_ant else COMPLETAR} millones). Estas {subio(k_im)} un {pct(k_im) if k_im is not None else COMPLETAR}, con respecto '
                 f'al mes de {mes(ant)} y {subio(k_ia, "se incrementaron", "se redujeron", "se mantuvieron")} un '
                 f'{pct(k_ia) if k_ia is not None else COMPLETAR} con respecto a {mes_anio(ia)}.')
    if k_23 is not None and p[0] > 2023:
        constante += (f' Si se compara con {mes_anio(a23)} las ventas de supermercado a precios constantes '
                      f'{subio(k_23, "aumentaron", "cayeron", "no variaron")} un {pct(k_23)}.')
    corriente = (f'Las ventas totales en supermercados de Formosa, superaron los ${num(int(c))} millones en el mes de {mes(p)} '
                 f'del {p[0]} (en {mes(ant)} fue de ${num(c_ant) if c_ant else COMPLETAR} millones). Estas {subio(c_im)} en un '
                 f'{pct(c_im) if c_im is not None else COMPLETAR}, con respecto al mes de {mes(ant)}, y '
                 f'{subio(c_ia, "se incrementaron", "se redujeron", "se mantuvieron")} un {pct(c_ia) if c_ia is not None else COMPLETAR} '
                 f'con respecto a {mes(p)} del año {ia[0]}.')
    return 'Supermercados', w, [('Supermercado precio constante', constante), ('Supermercado precio corriente', corriente)]


def emae(d: Datos, p):
    O, D = 'EMAE Indice Serie Original 2004=100', 'EMAE Indice Serie Desestacionalizada 2004=100'
    o, de = requerido(d.v(O, p), 'el índice EMAE'), d.v(D, p)
    # Variaciones publicadas por el INDEC (calculadas sin redondear); si faltan, se calculan con el índice.
    o_ia = d.v('EMAE Respecto a igual mes del año anterior (serie original) N', p) or var(o, d.v(O, mover(p, -12)))
    d_im = d.v('EMAE Respecto al mes anterior (desestacionalizado) N', p) or var(de, d.v(D, mover(p, -1)))
    d_ia = var(de, d.v(D, mover(p, -12)))
    o24, d24 = var(o, d.v(O, mover(p, -24))), var(de, d.v(D, mover(p, -24)))
    o23, d23 = var(o, d.v(O, mover(p, -36))), var(de, d.v(D, mover(p, -36)))
    sectores = []
    for i, titulo in d.columnas('EMAE por sector de actividad económica. Var% respecto a igual período del año anterior'):
        x = d.v(i, p)
        if x is not None and 'impuestos netos' not in norm(titulo):  # no es un sector de actividad
            nombre = re.sub(r'(?i)^.*?a[ñn]o anterior\.?\s*', '', titulo).strip().rstrip('.')
            sectores.append((nombre[:1].upper() + nombre[1:], x))
    suben = sorted([s for s in sectores if s[1] > 0], key=lambda s: -s[1])
    bajan = sorted([s for s in sectores if s[1] < 0], key=lambda s: s[1])
    lista = lambda xs: ', '.join(f'{n} ({con_signo(x)})' for n, x in xs[:-1]) + (' y ' if len(xs) > 1 else '') + \
        (f'{xs[-1][0]} ({con_signo(xs[-1][1])})' if xs else '')
    t = f'📊 **Estimador Mensual de Actividad Económica (EMAE) | {titulo_mes(p)}**'
    w = [('t', t), ('h', '📉 **Actividad económica nacional**' if (o_ia or 0) < 0 else '📈 **Actividad económica nacional**'),
         ('p', f'Durante {mes_anio(p)}, el EMAE registró una variación interanual del **{con_signo(o_ia)}** en la serie '
               f'original. Por su parte, la serie desestacionalizada presentó una variación del **{con_signo(d_im)} '
               f'intermensual** y del **{con_signo(d_ia)} interanual**.'),
         ('h', '📅 **Comparación con años anteriores**')]
    if o24 is not None:
        w.append(('b', f'Frente a {mes_anio(mover(p, -24))}: variación del **{con_signo(o24)}** en la serie original'
                       + (f' y del **{con_signo(d24)}** en la desestacionalizada.' if d24 is not None else '.')))
    if o23 is not None:
        w.append(('b', f'Frente a {mes_anio(mover(p, -36))}: variación del **{con_signo(o23)}** en la serie original'
                       + (f' y del **{con_signo(d23)}** en la desestacionalizada.' if d23 is not None else '.')))
    if sectores:
        w += [('h', '🏭 **Desempeño por sectores económicos**'),
              ('p', f'De los {len(sectores)} sectores que conforman el EMAE, **{len(suben)}** registraron aumentos '
                    f'interanuales y **{len(bajan)}** presentaron caídas.')]
        if suben:
            w.append(('p', f'📈 **Mayores aumentos:** {lista(suben[:3])}.'))
        if bajan:
            w.append(('p', f'📉 **Mayores caídas:** {lista(bajan[:3])}.'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/emae-indice/'))
    indice = (f'En {mes_anio(p)}, el EMAE registró en la serie original {"una suba" if (o_ia or 0) > 0 else "una baja"} del '
              f'{pct(o_ia)} respecto al mismo mes del año anterior. La serie desestacionalizada '
              f'{"aumentó" if (d_im or 0) > 0 else "disminuyó"} {pct(d_im)} intermensual y registró '
              f'{"un aumento" if (d_ia or 0) > 0 else "una disminución"} del {pct(d_ia)} interanual.')
    if o24 is not None:
        indice += (f' Comparado con {mes_anio(mover(p, -24))}, la serie original {"aumentó" if o24 > 0 else "cayó"} '
                   f'{pct(o24)}' + (f' y la desestacionalizada {pct(d24)}' if d24 is not None else '') + '.')
    if o23 is not None:
        indice += (f' Respecto a {mes_anio(mover(p, -36))}, la serie original registró '
                   f'{"un aumento" if o23 > 0 else "una caída"} de {pct(o23)}'
                   + (f' y la desestacionalizada de {pct(d23)}' if d23 is not None else '') + '.')
    indice += f' {COMPLETAR} (análisis de la trayectoria del año y comparación con el REM y el presupuesto).'
    sector = (f'{titulo_mes(p)} con relación a igual mes de {p[0] - 1}: {len(suben)} de los {len(sectores)} sectores de '
              f'la actividad que conforman el EMAE registraron subas interanuales en {mes_anio(p)}, mientras que los '
              f'{len(bajan)} restantes presentaron caídas; se destacaron {lista(suben[:2]) if suben else COMPLETAR}.'
              f' Sectores ganadores y perdedores frente a {mes_anio(mover(p, -36))}: {COMPLETAR} (la base solo tiene '
              f'la variación interanual por sector).')
    return 'EMAE', w, [('EMAE - Índice', indice), ('EMAE - Por sector de actividad', sector)]


def ipc(d: Datos, p):
    nea = [d.v('IPC nivel general con respecto al mes anterior (Region NEA)', p),
           d.v('IPC nivel general con respecto a diciembre del año anterior (IPC Acumulado) Region NEA', p),
           d.v('IPC nivel general con respecto al mismo mes del año anterior (Region NEA)', p)]
    nac = [d.v('IPC nivel general con respecto al mes anterior N', p),
           d.v('IPC nivel general con respecto a diciembre del año anterior (IPC acumulado) N', p),
           d.v('IPC nivel general con respecto al mismo mes del año anterior N', p)]
    requerido(nea[0] if nea[0] is not None else nac[0], 'el IPC del mes')
    divisiones = []
    for i, titulo in d.columnas('Var % mensual IPC nivel general segun divisiones (Region NEA)') + \
            d.columnas('Var. % mensual IPC nivel general segun divisiones (Region NEA)'):
        x = d.v(i, p)
        if x is not None:
            divisiones.append((titulo.split(':', 1)[-1].strip(), x))
    divisiones.sort(key=lambda s: -s[1])
    f = lambda x: pct(x) if x is not None else COMPLETAR
    w = [('t', f'📊 **Inflación (IPC) – NEA y país | {titulo_mes(p)}**'),
         ('h', '🛒 **Región Noreste (NEA)**'),
         ('p', f'En {mes_anio(p)} la inflación del NEA fue de **{f(nea[0])} mensual**, acumuló **{f(nea[1])}** en el año '
               f'y **{f(nea[2])} interanual**.'),
         ('h', '🇦🇷 **Nivel nacional**'),
         ('p', f'El IPC nacional registró **{f(nac[0])} mensual**, **{f(nac[1])}** acumulado en el año y '
               f'**{f(nac[2])} interanual**.')]
    if len(divisiones) >= 2:
        w += [('h', '🔝 **Divisiones en el NEA**'),
              ('b', f'Mayores subas: {divisiones[0][0]} ({pct(divisiones[0][1])}) y {divisiones[1][0]} ({pct(divisiones[1][1])}).'),
              ('b', f'Menores subas: {divisiones[-1][0]} ({pct(divisiones[-1][1])}) y {divisiones[-2][0]} ({pct(divisiones[-2][1])}).')]
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/ipc/'))
    texto = (f'En {mes(p)}, el Índice de Precios al Consumidor (IPC) nacional registró una variación mensual del {f(nac[0])} '
             f'y acumuló un {f(nac[1])} en lo que va del año, mientras que la medición interanual se ubicó en el {f(nac[2])}. '
             f'Por su parte, la región del Noreste (NEA) registró una inflación mensual del {f(nea[0])}, un incremento '
             f'acumulado en el año del {f(nea[1])} y una variación interanual del {f(nea[2])}.')
    return 'IPC', w, [('IPC - Índice de precios al consumidor', texto)]


def _registro(d: Datos, base: str, base_acum: str, p, extra_n: bool):
    """Patentamientos o transferencias (DNRPA): textos de la plataforma para Formosa y, si corresponde, país."""
    F, FA = f'{base} F', f'{base_acum} F'
    x = requerido(d.v(F, p), base.lower())
    ant, ia, a23 = mover(p, -1), mover(p, -12), (2023, p[1])
    def bloque(col, col_acum):
        v, v_ant, v_ia, v_23 = d.v(col, p), d.v(col, ant), d.v(col, ia), d.v(col, a23)
        acu = d.v(col_acum, p) or d.acumulado(col, p)
        acu_ia = d.v(col_acum, ia) or d.acumulado(col, ia)
        acu_23 = d.v(col_acum, a23) or d.acumulado(col, a23)
        return v, var(v, v_ant), var(v, v_ia), (v - v_ia) if v_ia else None, var(v, v_23), (v - v_23) if v_23 else None, \
            acu, var(acu, acu_ia), (acu - acu_ia) if acu_ia else None, var(acu, acu_23), (acu - acu_23) if acu_23 else None
    return bloque, x


def _texto_registro(nombre_cosa, que, p, b, nacional=False):
    v, im, ia, dia, v23, d23, acu, acu_ia, dacu, acu23, dacu23 = b
    u = lambda n: f'{num(abs(n))} unidades {"más" if n > 0 else "menos"}' if n is not None else COMPLETAR
    lugar = 'a nivel país' if nacional else 'En Formosa'
    s = (f'{lugar} se registró {"un incremento" if (im or 0) > 0 else "una disminución"} intermensual del '
         f'{pct(im) if im is not None else COMPLETAR}, con {num(v)} unidades {que} en el mes. ')
    s += (f'A nivel interanual, se registró {"un aumento" if (ia or 0) > 0 else "una caída"} del '
          f'{pct(ia) if ia is not None else COMPLETAR} ({u(dia)}). ')
    if v23 is not None and p[0] > 2023:
        s += (f'Si se lo compara con {mes_anio((2023, p[1]))}, el registro de unidades '
              f'{"aumentó" if v23 > 0 else "disminuyó"} un {pct(v23)} ({u(d23)}). ')
    if acu is not None:
        s += f'El acumulado a {mes(p)} de {p[0]} registró un total de {num(acu)} unidades'
        if acu_ia is not None:
            s += f', {pct(acu_ia)} {"más" if acu_ia > 0 else "menos"} comparado con el acumulado a {mes(p)} de {p[0] - 1} ({u(dacu)})'
        s += '.'
        if acu23 is not None and p[0] > 2023:
            s += (f' Comparado con el acumulado a {mes(p)} de 2023 se registró '
                  f'{"un incremento" if acu23 > 0 else "una caída"} del {pct(acu23)} ({u(dacu23)}).')
    return s


def _texto_nacional(cosa, p, b):
    v, im, ia, dia, v23, d23, acu, acu_ia, dacu, acu23, dacu23 = b
    u = lambda n: f'{num(abs(n))} unidades {"más" if n > 0 else "menos"}' if n is not None else COMPLETAR
    s = (f'En {mes_anio(p)} el patentamiento de {cosa} a nivel país registró {"una suba" if (im or 0) > 0 else "una baja"} '
         f'intermensual del {pct(im) if im is not None else COMPLETAR} y {"un incremento" if (ia or 0) > 0 else "una baja"} del '
         f'{pct(ia) if ia is not None else COMPLETAR} de manera interanual, totalizando {num(v)} unidades para el mes de referencia.')
    if v23 is not None and p[0] > 2023:
        s += f' Si se compara con {mes_anio((2023, p[1]))}, se registró {"un incremento" if v23 > 0 else "una caída"} del {pct(v23)} ({u(d23)}).'
    if acu is not None:
        s += f' El acumulado a {mes(p)} de {p[0]} registró {num(acu)} unidades'
        if acu_ia is not None:
            s += f', {pct(acu_ia)} {"más" if acu_ia > 0 else "menos"} que en el acumulado a {mes(p)} del año pasado'
        if acu23 is not None and p[0] > 2023:
            s += f' y {pct(acu23)} {"más" if acu23 > 0 else "menos"} si se lo compara con el mismo período de 2023'
        s += '.'
    return s


def autos(d: Datos, p):
    return _vehiculos(d, p, 'automóviles', 'Patentamiento de automóviles', 'Patentamiento de automóviles - Acumulado',
                      'Transferencias de automoviles', 'Transferencias de automoviles acumulada',
                      'Autos', 'patentamiento-auto', 'transferencia-auto', '🚗')


def motos(d: Datos, p):
    return _vehiculos(d, p, 'motos', 'Patentamiento de Motos', 'Patentamiento de Motos - Acumulado',
                      'Transferencias de Motos', 'Transferencias de Motos - Acumulado',
                      'Motos', 'patentamiento-moto', 'transferencia-moto', '🏍️')


def _vehiculos(d, p, cosa, pat, pat_acu, tra, tra_acu, nombre, url_pat, url_tra, emoji):
    bloque, _ = _registro(d, pat, pat_acu, p, True)
    bf = bloque(f'{pat} F', f'{pat_acu} F')
    bn = bloque(f'{pat} N', f'{pat_acu} N')
    tb, _ = _registro(d, tra, tra_acu, p, False) if d.v(f'{tra} F', p) else (None, None)
    tf = tb(f'{tra} F', f'{tra_acu} F') if tb else None
    w = [('t', f'📊 **Patentamientos y transferencias de {cosa} – Formosa | {titulo_mes(p)}**'),
         ('h', f'{emoji} **Patentamientos (0 km)**'),
         ('p', f'En {mes_anio(p)} se patentaron **{num(bf[0])} {cosa}** en Formosa, con una variación del '
               f'**{con_signo(bf[1]) if bf[1] is not None else COMPLETAR} intermensual** y del '
               f'**{con_signo(bf[2]) if bf[2] is not None else COMPLETAR} interanual**.')]
    if bf[6] is not None:
        w.append(('b', f'Acumulado del año: **{num(bf[6])} unidades** ({con_signo(bf[7]) if bf[7] is not None else COMPLETAR} '
                       f'frente al mismo período de {p[0] - 1}).'))
    if bn[0] is not None:
        w.append(('b', f'🇦🇷 A nivel país: **{num(bn[0])} unidades** ({con_signo(bn[1]) if bn[1] is not None else COMPLETAR} '
                       f'intermensual y {con_signo(bn[2]) if bn[2] is not None else COMPLETAR} interanual).'))
    if tf:
        w += [('h', '🔁 **Transferencias (usados)**'),
              ('p', f'Se transfirieron **{num(tf[0])} {cosa}**, con una variación del '
                    f'**{con_signo(tf[1]) if tf[1] is not None else COMPLETAR} intermensual** y del '
                    f'**{con_signo(tf[2]) if tf[2] is not None else COMPLETAR} interanual**.')]
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/{url_pat}/'))
    plataforma = []
    texto_pat = (_texto_nacional(cosa, p, bn) + ' ') if bn[0] is not None else ''
    texto_pat += _texto_registro(cosa, 'patentadas', p, bf)
    plataforma.append((f'{nombre} - Inscripción inicial', texto_pat))
    if tf:
        plataforma.append((f'{nombre} - Transferencias', _texto_registro(cosa, 'transferidas', p, tf)))
    return nombre, w, plataforma


def construccion(d: Datos, p):
    S = 'Salario promedio del sector Construcción (en pesos- con estacionalidad) F'
    P = 'Puestos de trabajo registrados en el sector construccion F'
    I = 'Indice de precios al consumidor- Nivel general (Región NEA)'
    s, pu = d.v(S, p), d.v(P, p)
    requerido(s or pu, 'salario y puestos de construcción')
    ant, ia, a23 = mover(p, -1), mover(p, -12), (2023, p[1])
    real = lambda nom, a, b: var((1 + nom / 100) * 100, (d.v(I, a) / d.v(I, b)) * 100) if nom is not None and d.v(I, a) and d.v(I, b) else None
    plataforma, w = [], [('t', f'📊 **Sector construcción – Formosa | {titulo_mes(p)}**')]
    if s:
        # Variaciones publicadas (IERIC) si están; si no, se calculan con el salario.
        s_im = d.v('Var. IM salario nominal del sector construccion (C/E) F', p) or var(s, d.v(S, ant))
        s_ia = d.v('Var. IA salario nominal del sector construccion (C/E) F', p) or var(s, d.v(S, ia))
        s_23 = var(s, d.v(S, a23))
        r_im, r_23 = real(s_im, p, ant), real(s_23, p, a23)
        r_ia = d.v('Var. IA salario real sector construcción (C/E)F', p) or real(s_ia, p, ia)
        inf_im = var(d.v(I, p), d.v(I, ant))
        inf_ia = var(d.v(I, p), d.v(I, ia))
        sac = ', en parte por efecto del SAC,' if p[1] in (1, 7) and (s_im or 0) < 0 else ''
        w += [('h', '💰 **Salario promedio**'),
              ('p', f'El salario promedio nominal fue de **${num(s)}** ({con_signo(s_im) if s_im is not None else COMPLETAR} '
                    f'intermensual y {con_signo(s_ia) if s_ia is not None else COMPLETAR} interanual). En términos reales: '
                    f'**{con_signo(r_im) if r_im is not None else COMPLETAR} i.m.** y **{con_signo(r_ia) if r_ia is not None else COMPLETAR} i.a.**')]
        t = (f'El salario promedio nominal del sector construcción en {mes_anio(p)} se ubicó en ${num(s)}. Registró '
             f'{"un incremento" if (s_im or 0) > 0 else "una caída"} del {pct(s_im) if s_im is not None else COMPLETAR} i.m{sac} y '
             f'{"un incremento" if (s_ia or 0) > 0 else "una caída"} del {pct(s_ia) if s_ia is not None else COMPLETAR} i.a.')
        if s_23 is not None and p[0] > 2023:
            t += f' Comparando este último registro con respecto al mismo mes de 2023, se registró {"un aumento" if s_23 > 0 else "una caída"} de {pct(s_23, 0)}.'
        t += (f' Salario real mes de {mes(p)}: Registró {"un aumento" if (r_im or 0) > 0 else "una caída"} real intermensual del '
              f'{pct(r_im) if r_im is not None else COMPLETAR}'
              + (f' (inflación mensual {pct(inf_im)})' if inf_im is not None else '')
              + f', y {"un aumento" if (r_ia or 0) > 0 else "una caída"} interanual de {pct(r_ia) if r_ia is not None else COMPLETAR} en términos reales'
              + (f', ya que el aumento nominal {pct(s_ia)} fue {"superior" if (r_ia or 0) > 0 else "inferior"} a la inflación interanual {pct(inf_ia)}' if inf_ia and s_ia is not None else '')
              + '.')
        if r_23 is not None and p[0] > 2023:
            t += (f' Comparando este último registro con respecto al mismo mes de 2023, se registró '
                  f'{"un aumento" if r_23 > 0 else "una caída"} del poder adquisitivo del {pct(r_23, 0)}.')
        plataforma.append(('Construcción - Salarios', t))
    if pu:
        p_ant, p_ia, p_23 = d.v(P, ant), d.v(P, ia), d.v(P, a23)
        im = d.v('Var. IM Puestos de trabajo registrados en el sector construccion F', p) or var(pu, p_ant)
        iav = d.v('Var. IA  Puestos de trabajo registrados en el sector construccion F', p) or var(pu, p_ia)
        v23 = var(pu, p_23)
        tr = lambda a, b: f'{num(abs(a - b))} trabajadores {"más" if a > b else "menos"}' if b else COMPLETAR
        w += [('h', '👷 **Puestos de trabajo registrados**'),
              ('p', f'**{num(pu)} puestos** ({con_signo(im) if im is not None else COMPLETAR} intermensual y '
                    f'{con_signo(iav) if iav is not None else COMPLETAR} interanual).')]
        t = (f'Puestos de trabajo registrados en el sector construcción mes de {mes(p)}, registró '
             f'{"un incremento" if (im or 0) > 0 else "una disminución"} de {pct(im) if im is not None else COMPLETAR} i.m., '
             f'lo que representa {tr(pu, p_ant)}. En términos i.a., los puestos de trabajo '
             f'{"aumentaron" if (iav or 0) > 0 else "disminuyeron"} un {pct(iav) if iav is not None else COMPLETAR}, '
             f'lo que representa, en valores absolutos, {tr(pu, p_ia)}.')
        if v23 is not None and p[0] > 2023:
            t += (f' Comparando este último registro con respecto a {mes_anio(a23)}, {"aumentó" if v23 > 0 else "disminuyó"} '
                  f'un {pct(v23)}, lo que representa, en valores absolutos, {tr(pu, p_23)}.')
        plataforma.append(('Construcción - Puestos de trabajo', t))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/salario-construccion/'))
    return 'Construccion', w, plataforma


def sector_privado(d: Datos, p):
    F = 'Cant. asalariados registrados del sector privado. Sin estacionalidad F'
    N = 'Cant. asalariados registrados del sector privado. Sin estacionalidad. En miles. N'
    f = requerido(d.v(F, p), 'asalariados privados de Formosa')
    ant, ia, a23 = mover(p, -1), mover(p, -12), (2023, p[1])
    n, n_ant, n_ia = d.v(N, p), d.v(N, ant), d.v(N, ia)
    f_ant, f_ia, f_23 = d.v(F, ant), d.v(F, ia), d.v(F, a23)
    tr = lambda a, b: f'{num(abs(a - b))} trabajadores {"más" if a > b else "menos"}' if b else COMPLETAR
    # Diferencias publicadas en la base (SIPA) si están.
    dif_im, dif_ia = d.v('Var. absoluta i.m. asalariados registrados del sector privado S/E F', p), d.v('Var. absoluta i.a. asalariados registrados del sector privado S/E F', p)
    if dif_im is not None:
        f_ant = f - dif_im
    if dif_ia is not None:
        f_ia = f - dif_ia
    texto = ''
    if n:
        texto += (f'En el mes de {mes_anio(p)} se registraron, según el informe del SIPA (serie desestacionalizada), un total de '
                  f'{num(n, 1)} miles de asalariados registrados del sector privado en el país. Esto representa '
                  f'{"un aumento" if (var(n, n_ant) or 0) > 0 else "una disminución"} del {pct(var(n, n_ant), 2) if n_ant else COMPLETAR} con respecto al mes anterior '
                  f'(aproximadamente {num(abs(n - n_ant) * 1000) if n_ant else COMPLETAR} puestos {"más" if n_ant and n > n_ant else "menos"}) y '
                  f'{"un aumento" if (var(n, n_ia) or 0) > 0 else "una caída"} del {pct(var(n, n_ia), 2) if n_ia else COMPLETAR} interanual '
                  f'({num(abs(n - n_ia) * 1000) if n_ia else COMPLETAR} puestos {"más" if n_ia and n > n_ia else "menos"}). ')
    im, iav, v23 = var(f, f_ant), var(f, f_ia), var(f, f_23)
    texto += (f'En Formosa, en la serie sin estacionalidad, el número de asalariados registrados del sector privado fue de '
              f'{num(f)} en {mes(p)}, con una variación {"positiva" if (im or 0) > 0 else "negativa"} del {pct(im) if im is not None else COMPLETAR} i.m. '
              f'(equivalente a {tr(f, f_ant)}) y {"un aumento" if (iav or 0) > 0 else "una disminución"} del '
              f'{pct(iav) if iav is not None else COMPLETAR} i.a. ({tr(f, f_ia)}).')
    if v23 is not None and p[0] > 2023:
        texto += f' Si se compara con {mes_anio(a23)}, esto representa {"un aumento" if v23 > 0 else "una disminución"} del {pct(v23)} ({tr(f, f_23)}).'
    # Ramas de actividad (SIPA, país, serie desestacionalizada): cuántas suben, se mantienen o bajan.
    ramas = [i for i, _ in d.columnas('Cant. asalariados registrados del sector privado (en miles), según rama de actividad')]
    def conteo(b):
        arriba = igual = abajo = 0
        for i in ramas:
            a, c = d.v(i, p), d.v(i, b)
            if a is None or c is None:
                continue
            x = round(var(a, c), 1)
            arriba, igual, abajo = arriba + (x > 0), igual + (x == 0), abajo + (x < 0)
        return arriba, igual, abajo
    w = [('t', f'📊 **Empleo privado registrado – Formosa | {titulo_mes(p)}**'),
         ('p', f'En {mes_anio(p)} Formosa tuvo **{num(f)} asalariados registrados** en el sector privado (serie sin '
               f'estacionalidad): **{con_signo(im) if im is not None else COMPLETAR} intermensual** ({tr(f, f_ant)}) y '
               f'**{con_signo(iav) if iav is not None else COMPLETAR} interanual** ({tr(f, f_ia)}).')]
    if v23 is not None and p[0] > 2023:
        w.append(('b', f'Frente a {mes_anio(a23)}: **{con_signo(v23)}** ({tr(f, f_23)}).'))
    if n:
        w.append(('b', f'🇦🇷 País: {num(n, 1)} miles de asalariados privados ({con_signo(var(n, n_ant), 2) if n_ant else COMPLETAR} i.m. y '
                       f'{con_signo(var(n, n_ia), 2) if n_ia else COMPLETAR} i.a.).'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/sector-privado/'))
    plataforma = [('Sector privado - Evolución del empleo', texto)]
    if ramas:
        a1, i1, b1 = conteo(ant)
        a2, i2, b2 = conteo(ia)
        a3, i3, b3 = conteo(a23)
        if a1 + i1 + b1:
            r = (f'En {mes_anio(p)}, se observó una dinámica {"heterogénea" if a1 and b1 else "homogénea"} del empleo a nivel sectorial. '
                 f'El análisis intermensual desestacionalizado revela que {a1} ramas de actividad aumentaron sus puestos de trabajo, '
                 f'{i1} se mantuvieron estables y las {b1} restantes experimentaron contracciones.')
            if a2 + i2 + b2:
                r += f' En la comparación interanual, {b2} sectores registraron pérdida de ocupación, {i2} se mantuvieron sin variaciones y {a2} evidenciaron aumentos.'
            if a3 + i3 + b3 and p[0] > 2023:
                r += f' Por último, con respecto al mismo mes del año 2023, {b3} sectores registraron caídas en los puestos de trabajo y {a3} mostraron crecimientos.'
            plataforma.append(('Sector privado - Trabajadores por rama', r))
    return 'Sector_privado', w, plataforma


def recaudacion(d: Datos, p):
    R, A = 'Recaudación tributaria F', 'Recaudación tributaria acumulada F'
    I = 'Indice de precios al consumidor- Nivel general (Región NEA)'
    r = requerido(d.v(R, p), 'la recaudación del mes')
    ant, ia, a23 = mover(p, -1), mover(p, -12), (2023, p[1])
    r_ant = d.v(R, ant)
    im_nom = var(r, r_ant)
    im_real = d.v('Variación intermensual recaudación tributaria F Real', p)
    ia_real = d.v('Variación interanual recaudación tributaria F Real', p)
    infl = lambda a, b: d.v(I, a) / d.v(I, b) if d.v(I, a) and d.v(I, b) else None
    r23 = var(r / infl(p, a23), d.v(R, a23)) if infl(p, a23) and d.v(R, a23) else None
    iibb, part = d.v('Recaudación de IIBB F', p), d.v('IIBB/Recaudacion tributaria(%) F', p)
    iibb_im = d.v('Variación intermensual recaudación de IIBB F Real', p)
    iibb_ia = d.v('Variación interanual recaudación de IIBB F Real', p)
    iibb23 = var(iibb / infl(p, a23), d.v('Recaudación de IIBB F', a23)) if iibb and infl(p, a23) and d.v('Recaudación de IIBB F', a23) else None
    acu, acu_ia = d.v(A, p) or d.acumulado(R, p), d.v(A, ia) or d.acumulado(R, ia)
    f = lambda x: pct(x) if x is not None else COMPLETAR
    texto = (f'La recaudación provincial del mes de {mes(p)} en Formosa experimentó {"un incremento" if (im_nom or 0) > 0 else "una caída"} '
             f'intermensual del {pct(im_nom, 2) if im_nom is not None else COMPLETAR} en términos nominales, siendo de ${num(r)} en el mes de '
             f'{mes(p)} en comparación con los ${num(r_ant) if r_ant else COMPLETAR} registrados en {mes(ant)}. Al ajustar estos valores por '
             f'inflación{f" (IPC NEA de {mes(p)}: {pct(var(d.v(I, p), d.v(I, ant)))})" if infl(p, ant) else f" ({COMPLETAR} estimación de inflación utilizada)"}, '
             f'la variación real intermensual arroja {"un incremento" if (im_real or 0) > 0 else "una caída"} del {f(im_real)}; y la variación '
             f'real interanual arroja {"un incremento" if (ia_real or 0) > 0 else "una caída"} del {f(ia_real)}.')
    texto += (f' La comparación real con {mes_anio(a23)} registra así {"un incremento" if (r23 or 0) > 0 else "una caída"} del {pct(r23, 0)}.'
              if r23 is not None else f' La comparación real con {mes_anio(a23)}: {COMPLETAR}.')
    if iibb:
        texto += (f' El impuesto sobre los Ingresos Brutos (IIBB) sigue siendo un pilar fundamental de la recaudación, contribuyendo con '
                  f'el {f(part)} de la recaudación total en {mes_anio(p)}. El mismo tuvo una variación real '
                  f'{"positiva" if (iibb_im or 0) > 0 else "negativa"} del {f(iibb_im)} respecto al mes de {mes(ant)} y la variación '
                  f'interanual real registró {"un aumento" if (iibb_ia or 0) > 0 else "una caída"} del {f(iibb_ia)}.')
        if iibb23 is not None:
            texto += f' La comparación real con {mes_anio(a23)} registra {"un aumento" if iibb23 > 0 else "una caída"} del {pct(iibb23)}.'
    if acu:
        texto += (f' La recaudación acumulada durante el año {p[0]} ascendió a ${num(acu)}'
                  + (f', registrando un crecimiento nominal del {pct(var(acu, acu_ia), 2)} en comparación con el acumulado de {p[0] - 1}.' if acu_ia else '.'))
    w = [('t', f'📊 **Recaudación tributaria provincial (ATP) – Formosa | {titulo_mes(p)}**'),
         ('p', f'En {mes_anio(p)} la recaudación fue de **${num(r / 1e6)} millones**, con una variación real de '
               f'**{con_signo(im_real) if im_real is not None else COMPLETAR} intermensual** y **{con_signo(ia_real) if ia_real is not None else COMPLETAR} interanual**.')]
    if iibb:
        w.append(('b', f'🧾 Ingresos Brutos: **{f(part)}** del total ({con_signo(iibb_ia) if iibb_ia is not None else COMPLETAR} real interanual).'))
    if r23 is not None:
        w.append(('b', f'📅 Frente a {mes_anio(a23)}: **{con_signo(r23)}** en términos reales.'))
    if acu:
        w.append(('b', f'Acumulado {p[0]}: **${num(acu / 1e6)} millones**' + (f' ({con_signo(var(acu, acu_ia))} nominal).' if acu_ia else '.')))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/recaudacion/'))
    return 'Recaudacion_ATP', w, [('ATP - Recaudación', texto)]


def transferencias(d: Datos, p):
    T = 'Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ F'
    t = requerido(d.v(T, p), 'las transferencias del mes')
    nom = d.v('Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Nominal F', p)
    real = d.v('Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Var% Anual Real F', p)
    real_n = d.v('Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Real N', p)
    acu = d.v('Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Acumulado anual', p) or d.acumulado(T, p)
    infl = ((1 + nom / 100) / (1 + real / 100) - 1) * 100 if nom is not None and real is not None else None
    texto = (f'En {mes_anio(p)}, las transferencias automáticas (coparticipación, leyes complementarias y compensaciones) enviadas a '
             f'Formosa alcanzaron ${num(t)} millones, lo que representó una variación nominal {"positiva" if (nom or 0) > 0 else "negativa"} '
             f'del {pct(nom) if nom is not None else COMPLETAR} respecto a {mes_anio(mover(p, -12))}.')
    if real is not None:
        texto += (f' Al descontar la inflación estimada del {pct(infl)} interanual, se registró {"un incremento" if real > 0 else "una caída"} '
                  f'real interanual del {pct(real)}'
                  + (f', {"por encima" if real > real_n else "por debajo"} del promedio del consolidado de provincias y CABA ({pct(real_n)})' if real_n is not None else '')
                  + '.')
    if acu:
        texto += f' En el acumulado de los {p[1]} meses del año en curso, Formosa recibió ${num(acu)} millones, con una variación real acumulada de {COMPLETAR} frente al período equivalente de {p[0] - 1}.'
    texto += f' {COMPLETAR} (comparación con las provincias del NEA y con {mes_anio((2023, p[1]))}).'
    w = [('t', f'📊 **Transferencias automáticas a Formosa | {titulo_mes(p)}**'),
         ('p', f'En {mes_anio(p)} Formosa recibió **${num(t)} millones** por coparticipación, leyes complementarias y compensaciones: '
               f'**{con_signo(nom) if nom is not None else COMPLETAR} nominal** y **{con_signo(real) if real is not None else COMPLETAR} real** interanual.')]
    if real_n is not None:
        w.append(('b', f'🇦🇷 Consolidado de provincias y CABA: {con_signo(real_n)} real interanual.'))
    if acu:
        w.append(('b', f'Acumulado {p[0]}: **${num(acu)} millones**.'))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/transferencias/'))
    return 'Transferencias', w, [('Transferencias automáticas', texto)]


def energia(d: Datos, p):
    cats = [('Residencial', 'Demanda (MWh) Categoría Tarifa Residencial F', 'consumo-energia-residencial'),
            ('Industrial', 'Demanda (MWh) Categoría Tarifa Industrial / Comercial Grande F', 'consumo-energia-industrial'),
            ('Comercial', 'Demanda (MWh) Categoría Tarifa Comercial F', 'consumo-energia-comercial')]
    vals = {n: d.v(c, p) for n, c, _ in cats}
    requerido(all(vals.values()) or None, 'la demanda de las tres tarifas')
    total = sum(vals.values())
    ant, ia, a23 = mover(p, -1), mover(p, -12), (2023, p[1])
    plataforma, resumen = [], []
    w = [('t', f'📊 **Demanda de energía eléctrica – Formosa | {titulo_mes(p)}**')]
    for nombre, col, slug in cats:
        v = vals[nombre]
        im, iav, v23 = var(v, d.v(col, ant)), var(v, d.v(col, ia)), var(v, d.v(col, a23))
        part = v / total * 100
        signo = lambda x: f'variación {"positiva" if x > 0 else "negativa"} del {pct(x)}' if x is not None else COMPLETAR
        t = (f'En el mes de {mes_anio(p)} representó el {pct(part)} de la demanda total de energía eléctrica'
             f'{" de la provincia" if nombre == "Residencial" else ""}. El registro en este caso arrojó {num(v)} MWh, '
             f'con una {signo(im)} i.m y una {signo(iav)} i.a.')
        if v23 is not None and p[0] > 2023:
            t += (f' Al realizar la comparación con el mes de {mes_anio(a23)}, se registra '
                  f'{"un aumento" if v23 > 0 else "una disminución"} del {pct(v23)} en la demanda del sector.')
        plataforma.append((f'Energía - Tarifa {nombre.lower()}', t))
        resumen.append(f'Tarifa {nombre}: ' + t.replace(f'En el mes de {mes_anio(p)} representó', 'representó'))
        w.append(('b', f'**{nombre}:** {num(v)} MWh ({pct(part)} del total) · {con_signo(im) if im is not None else COMPLETAR} i.m. · '
                       f'{con_signo(iav) if iav is not None else COMPLETAR} i.a.'))
    plataforma.insert(0, ('Energía - Resumen', f'{titulo_mes(p)} / ' + ' / '.join(resumen)))
    w.append(('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/demanda-usuarios/'))
    return 'Energia', w, plataforma


def internet(d: Datos, p):
    T, TV, TVN = 'Cantidad de accesos a internet F', 'Cantidad de accesos a internet - Var. Interanual (%) F', 'Cantidad de accesos a internet - Var. Interanual (%) N'
    FV, FVN = 'Cantidad de accesos fijos a internet - Var. Interanual (%) F', 'Cantidad de accesos fijos a internet - Var. Interanual (%) N'
    MV, MVN = 'Cantidad de accesos moviles a internet - Var. Interanual (%) F', 'Cantidad de accesos moviles a internet - Var. Interanual (%) N'
    t = requerido(d.v(T, p), 'los accesos a internet')
    meses3 = [mover(p, -2), mover(p, -1), p]
    serie = lambda c: [d.v(c, m) for m in meses3]
    tv, tvn = d.v(TV, p) or var(t, d.v(T, mover(p, -12))), d.v(TVN, p)
    f = lambda xs: ', '.join(f'{pct(x)} en {mes(m)}' for x, m in zip(xs, meses3) if x is not None) or COMPLETAR
    fij, fijn, mov, movn = serie(FV), serie(FVN), serie(MV), serie(MVN)
    brecha = [f'{num(a - b, 1)} pp. en {mes(m)}' for a, b, m in zip(fij, fijn, meses3) if a is not None and b is not None]
    texto = (f'Formosa registró {num(t)} accesos totales en {mes_anio(p)}, con un crecimiento del {pct(tv) if tv is not None else COMPLETAR} interanual'
             + (f', frente a la media nacional ({pct(tvn)})' if tvn is not None else '') + '. '
             f'Los accesos fijos crecieron en términos interanuales un {f(fij)}'
             + (f', con una diferencia frente al desempeño nacional de {", ".join(brecha)}' if brecha else '') + '. '
             f'En lo que respecta a accesos móviles, crecieron en términos interanuales un {f(mov)}'
             + (f' (nacional: {f(movn)})' if any(movn) else '') + f'. {COMPLETAR} (lectura del trimestre).')
    w = [('t', f'📊 **Accesos a internet – Formosa | {titulo_mes(p)}**'),
         ('p', f'Formosa registró **{num(t)} accesos** en {mes_anio(p)}, **{con_signo(tv) if tv is not None else COMPLETAR} interanual**'
               + (f' (media nacional: {con_signo(tvn)}).' if tvn is not None else '.')),
         ('b', f'🏠 Accesos fijos: {f(fij)} interanual.'),
         ('b', f'📱 Accesos móviles: {f(mov)} interanual.'),
         ('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/acceso-internet/')]
    return 'Internet', w, [('Internet - Acceso a internet', texto)]


def pobreza(d: Datos, p):
    HP, PP, CP = 'Pobreza (% sobre el total de hogares) F', 'Pobreza (% sobre el total de personas) F', 'Pobreza (cant. personas) F'
    HI, PI, CI = 'Indigencia (% sobre el total de hogares) F', 'Indigencia (% sobre el total de personas) F', 'Indigencia (cant. personas) F'
    if p[1] not in (6, 12):  # la EPH es semestral: junio = 1er semestre, diciembre = 2do
        raise Falta('la pobreza es semestral (junio o diciembre)')
    hp, pp = requerido(d.v(HP, p), 'la pobreza en hogares'), d.v(PP, p)
    sem = 'primer' if p[1] <= 6 else 'segundo'
    ia, a23 = mover(p, -12), (2023, p[1])
    pp_txt = lambda a, b: (f'{"se incrementó" if a > b else "se redujo"} {num(abs(a - b), 1)} puntos porcentuales '
                           f'(de {num(b, 1)}% a {num(a, 1)}%)') if a is not None and b is not None else COMPLETAR
    texto = (f'En el {sem} semestre de {p[0]} el aglomerado Formosa registró una tasa de pobreza en hogares del {num(hp, 1)}% y en '
             f'personas del {num(pp, 1) if pp else COMPLETAR}%; en términos absolutos, {num(d.v(CP, p)) if d.v(CP, p) else COMPLETAR} personas residen '
             f'en hogares por debajo de la línea de pobreza. Entre el {sem} semestre de {ia[0]} y el {sem} semestre de {p[0]}, la pobreza '
             f'en hogares de Formosa {pp_txt(hp, d.v(HP, ia))} y la pobreza en personas {pp_txt(pp, d.v(PP, ia))}.')
    if p[0] > 2023:
        texto += (f' Si se extiende el análisis entre el {sem} semestre de 2023 y el {sem} semestre de {p[0]}, la pobreza en hogares '
                  f'{pp_txt(hp, d.v(HP, a23))} y la pobreza en personas {pp_txt(pp, d.v(PP, a23))}.')
    texto += f' {COMPLETAR} (comparación con el promedio del NEA y párrafo de contexto sobre políticas provinciales).'
    hi, pi = d.v(HI, p), d.v(PI, p)
    indig = (f'La indigencia afectó al {num(hi, 1) if hi else COMPLETAR}% de los hogares y al {num(pi, 1) if pi else COMPLETAR}% de las personas. '
             f'En términos absolutos, {num(d.v(CP, p)) if d.v(CP, p) else COMPLETAR} personas residen en hogares por debajo de la línea de '
             f'pobreza, y {num(d.v(CI, p)) if d.v(CI, p) else COMPLETAR} personas se encuentran en situación de indigencia. En el período '
             f'analizado, la indigencia en hogares {pp_txt(hi, d.v(HI, ia))} y la indigencia en personas {pp_txt(pi, d.v(PI, ia))}. '
             f'{COMPLETAR} (comparación con el promedio del NEA).')
    w = [('t', f'📊 **Pobreza e indigencia – Formosa | {sem.capitalize()} semestre {p[0]}**'),
         ('p', f'En el {sem} semestre de {p[0]}, la pobreza alcanzó al **{num(hp, 1)}% de los hogares** y al '
               f'**{num(pp, 1) if pp else COMPLETAR}% de las personas** del aglomerado Formosa.'),
         ('b', f'Indigencia: **{num(hi, 1) if hi else COMPLETAR}%** de los hogares y **{num(pi, 1) if pi else COMPLETAR}%** de las personas.'),
         ('b', f'Frente al {sem} semestre de {ia[0]}: pobreza en personas {pp_txt(pp, d.v(PP, ia))}.'),
         ('p', f'🔎 **Más información en nuestro Observatorio Económico:** {WEB}/pobreza/')]
    return 'Pobreza', w, [('Indicador social - Pobreza', texto), ('Indicador social - Indigencia', indig)]


# fuente de la revisión -> indicadores del Observatorio que dependen de ella
INDICADORES = {
    'supermercados': supermercados, 'emae': emae, 'ipc': ipc, 'autos': autos, 'motos': motos,
    'construccion': construccion, 'sector_privado': sector_privado, 'recaudacion': recaudacion,
    'transferencias': transferencias, 'energia': energia, 'internet': internet, 'pobreza': pobreza,
}
NOMBRES = {'supermercados': 'Supermercados', 'emae': 'EMAE', 'ipc': 'IPC', 'autos': 'Autos', 'motos': 'Motos',
           'construccion': 'Construcción', 'sector_privado': 'Sector privado', 'recaudacion': 'Recaudación ATP',
           'transferencias': 'Transferencias', 'energia': 'Energía', 'internet': 'Internet', 'pobreza': 'Pobreza e indigencia'}
POR_FUENTE = {
    'supermercados': ['supermercados'], 'emae': ['emae'], 'ipc': ['ipc'], 'dnrpa': ['autos', 'motos'],
    'ieric': ['construccion'], 'ieric_correccion': ['construccion'], 'sipa': ['sector_privado'], 'atp': ['recaudacion'],
    'transferencias_fiscales': ['transferencias'], 'correccion_transferencias': ['transferencias'],
    'cammesa': ['energia'], 'internet': ['internet'], 'pobreza': ['pobreza'],
}


# ------------------------------------------------------------------------------- Word y Telegram
def armar_word(nombre, bloques, plataforma, p) -> Path:
    doc = Documento()
    doc.parrafo(f'Borrador generado automáticamente el {datetime.now():%d/%m/%Y %H:%M} con los datos de la BASE DE DATOS de DAEE '
                f'(Dirección de Asuntos Estratégicos y Económicos). Revisar antes de publicar; lo marcado en amarillo se completa a mano.',
                'Nota')
    doc.titulo('Plantilla para WhatsApp', 2)
    for tipo, texto in bloques:
        if tipo == 't':
            doc.parrafo(texto)
        elif tipo == 'h':
            doc.parrafo(texto)
        elif tipo == 'b':
            doc.vineta(texto)
        else:
            doc.parrafo(texto)
    doc.separador()
    doc.titulo('Descripción plataforma', 2)
    for subtitulo, texto in plataforma:
        doc.parrafo(f'**{subtitulo}**')
        doc.parrafo(texto)
    ruta = CARPETA / f'{p[0]}-{p[1]:02d}' / f'{nombre}_{MESES[p[1]]}{p[0]}.docx'
    return doc.guardar(ruta)


def _token():
    token = os.environ.get('TELEGRAM_BOT_TOKEN')
    if not token and CONFIG_TELEGRAM.is_file():
        for linea in CONFIG_TELEGRAM.read_text(encoding='utf-8').splitlines():
            if linea.strip().startswith('TELEGRAM_BOT_TOKEN='):
                token = linea.split('=', 1)[1].strip()
    m = re.search(r'\d{6,}:[A-Za-z0-9_-]{30,}', token or '')
    return m.group(0) if m else None


def enviar_telegram(ruta: Path, texto: str) -> str:
    token = _token()
    if not token:
        return 'sin token de Telegram (config/telegram.env)'
    limite = '----' + uuid.uuid4().hex
    partes = []
    for campo, valor in (('chat_id', CHAT_GRUPO), ('caption', texto)):
        partes.append(f'--{limite}\r\nContent-Disposition: form-data; name="{campo}"\r\n\r\n{valor}\r\n'.encode('utf-8'))
    partes.append((f'--{limite}\r\nContent-Disposition: form-data; name="document"; filename="{ruta.name}"\r\n'
                   'Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n').encode('utf-8'))
    partes.append(ruta.read_bytes())
    partes.append(f'\r\n--{limite}--\r\n'.encode('utf-8'))
    pedido = urllib.request.Request(f'https://api.telegram.org/bot{token}/sendDocument', data=b''.join(partes),
                                    headers={'Content-Type': f'multipart/form-data; boundary={limite}'})
    with urllib.request.urlopen(pedido, timeout=120) as r:
        respuesta = json.loads(r.read().decode('utf-8'))
    return 'enviado' if respuesta.get('ok') else respuesta.get('description', 'error de Telegram')


def generar(indicador: str, periodo=None, datos: Datos | None = None, telegram=True) -> dict:
    from sheet_directo import periodo_de
    datos = datos or cargar_datos()
    funcion = INDICADORES[indicador]
    p = periodo_de(periodo) if periodo else None
    if p is None:  # último período con el dato principal
        for candidato in sorted(datos.filas, reverse=True):
            try:
                funcion(datos, candidato)
                p = candidato
                break
            except Falta:
                continue
    if p is None:
        return {'indicador': indicador, 'estado': 'SIN_DATOS'}
    try:
        nombre, bloques, plataforma = funcion(datos, p)
    except Falta as falta:
        return {'indicador': indicador, 'periodo': f'{p[0]}-{p[1]:02d}', 'estado': 'SIN_DATOS', 'detalle': str(falta)}
    ruta = armar_word(nombre, bloques, plataforma, p)
    resultado = {'indicador': indicador, 'periodo': f'{p[0]}-{p[1]:02d}', 'estado': 'OK', 'archivo': str(ruta)}
    if telegram:
        try:
            resultado['telegram'] = enviar_telegram(
                ruta, f'📝 Descripción para el Observatorio: {nombre.replace("_", " ")} – {titulo_mes(p)}.\n'
                      'Plantilla de WhatsApp + texto de la plataforma (revisar lo marcado en amarillo).')
        except Exception as error:  # noqa: BLE001
            resultado['telegram'] = f'ERROR: {error}'
    return resultado


def generar_desde_resumen(resumen: dict, telegram=True) -> list[dict]:
    """Para cada fuente con datos nuevos en esta revisión, arma el Word de sus indicadores publicados."""
    from sheet_directo import periodo_de
    pedidos = {}
    for r in resumen.get('resultados', []):
        cambios = r.get('hubo_cambios') in (True, 'True', 'true') or int(r.get('celdas_actualizadas') or 0) > 0
        if r.get('estado') != 'OK' or not cambios:
            continue
        for indicador in POR_FUENTE.get(str(r.get('fuente', '')).lower(), []):
            pedidos[indicador] = periodo_de(str(r.get('periodo') or '')[:10]) if r.get('periodo') else None
    if not pedidos:
        return []
    datos = cargar_datos()
    salida = []
    for indicador, p in pedidos.items():
        try:
            salida.append(generar(indicador, p, datos, telegram))
        except Exception as error:  # noqa: BLE001
            salida.append({'indicador': indicador, 'estado': 'ERROR', 'detalle': str(error)[:300]})
    return salida


def para_revision(resumen: dict) -> tuple[list[dict], str]:
    """Lo llaman las revisiones (general y complejos). Nunca frena la revisión.
    Devuelve (detalle, línea para agregar al aviso de Telegram)."""
    try:
        lista = generar_desde_resumen(resumen)
    except Exception as error:  # noqa: BLE001
        lista = [{'estado': 'ERROR', 'detalle': str(error)[:300]}]
    return lista, _lineas(lista)


def _lineas(lista: list[dict]) -> str:
    if not lista:
        return ''
    lineas = []
    for x in lista:
        nombre = NOMBRES.get(x.get('indicador'), x.get('indicador', 'Descripción'))
        if x.get('estado') == 'OK':
            t = x.get('telegram')
            envio = ' y enviado al grupo' if t == 'enviado' else ' (guardado en la carpeta descripciones)' if t is None else \
                f' (guardado en la carpeta descripciones; no se pudo mandar por Telegram: {t})'
            a, m = (x.get('periodo') or '0000-00').split('-')
            lineas.append(f"📝 {nombre} ({MESES[int(m)][:3]} {a[2:]}): Word con la descripción para el Observatorio listo{envio}.")
        elif x.get('estado') == 'SIN_DATOS':
            lineas.append(f"📝 {nombre}: no se armó la descripción ({x.get('detalle', 'faltan datos en la base')}).")
        else:
            lineas.append(f"⚠️ {nombre}: error al armar la descripción ({x.get('detalle', '')}).")
    return '\n'.join(lineas)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--resumen', help='resumen_*.json de una revisión')
    ap.add_argument('--indicador', help=', '.join(INDICADORES))
    ap.add_argument('--periodo', help='AAAA-MM (por defecto, el último cargado)')
    ap.add_argument('--sin-telegram', action='store_true')
    ap.add_argument('--listar', action='store_true')
    args = ap.parse_args()
    if args.listar:
        print(json.dumps({'indicadores': list(INDICADORES), 'por_fuente': POR_FUENTE}, ensure_ascii=False, indent=1))
        return 0
    if args.resumen:
        resumen = json.loads(Path(args.resumen).read_text(encoding='utf-8-sig'))
        salida = generar_desde_resumen(resumen, not args.sin_telegram)
    elif args.indicador:
        salida = [generar(x.strip(), args.periodo, telegram=not args.sin_telegram) for x in args.indicador.split(',')]
    else:
        ap.error('indicar --resumen o --indicador')
    print('DESCRIPCIONES=' + json.dumps(salida, ensure_ascii=False))
    if args.indicador:  # corrida a mano (o "descripcion_prueba" desde n8n)
        print('NOTA_TELEGRAM=' + _lineas(salida).replace('\n', ' · '))
        print('ESTADO=' + ('OK' if all(x.get('estado') == 'OK' for x in salida) else 'ERROR'))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
