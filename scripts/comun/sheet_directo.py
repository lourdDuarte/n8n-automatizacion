"""Escritura directa y verificada en el Google Sheet del Observatorio.

Reemplaza la escritura que hacía el subworkflow de n8n. Para cada carga
(hoja, período, {encabezado: valor}):

* ubica la fila por la columna PERIODO y la columna por el encabezado;
* escribe solo lo que cambió (lo que ya está igual no se toca);
* las fórmulas vienen del Excel local y se traducen a las columnas y filas
  reales del Sheet (el Sheet tiene columnas en otras posiciones);
* nunca pisa una fórmula del Sheet con un número ni un dato manual con una
  fórmula;
* relee lo escrito y verifica.

Devuelve, por fuente, cuántas celdas se escribieron, cuáles ya estaban y
qué encabezados no se encontraron, para que el aviso de Telegram diga
exactamente qué pasó.
"""
from __future__ import annotations

import math
import os
import re
import unicodedata
import urllib.parse
from datetime import date, datetime, timedelta
from pathlib import Path

SHEET_ID = '1cVd_KBbtjMN8vUHlUks1IKwPf0c4Oq0RcYn4B5N5PkA'
HOJA = 'BASE DE DATOS'
API = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}'
MESES = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
         'jul': 7, 'ago': 8, 'sep': 9, 'set': 9, 'oct': 10, 'nov': 11, 'dic': 12}
MESES_TXT = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')

LIMITE_REEMPLAZOS = 12  # datos ya cargados que una carga puede corregir

ICC_MENSUAL = 'Indice de Confianza del Consumidor Interior Variación respecto al mes anterior'
ICC_ACUMULADO = 'Indice de Confianza del Consumidor Interior Acumulado del año'


# ---------------------------------------------------------------- utilidades
def clave(texto) -> str:
    simple = unicodedata.normalize('NFD', str(texto or ''))
    simple = ''.join(c for c in simple if unicodedata.category(c) != 'Mn')
    simple = ' '.join(simple.split()).casefold()
    simple = re.sub(r'\s*([.,:;()\-/%*])\s*', r'\1', simple)
    return simple


def letra(numero: int) -> str:
    texto = ''
    while numero:
        numero, resto = divmod(numero - 1, 26)
        texto = chr(65 + resto) + texto
    return texto


def numero_columna(letras: str) -> int:
    n = 0
    for c in letras:
        n = n * 26 + ord(c) - 64
    return n


def periodo_de(valor):
    """(año, mes) desde fecha, número de serie, '2026-09-01', 'sep 26', etc."""
    if valor in (None, ''):
        return None
    if isinstance(valor, (datetime, date)):
        return (valor.year, valor.month)
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if 20000 < valor < 80000:
            f = date(1899, 12, 30) + timedelta(days=int(valor))
            return (f.year, f.month)
        return None
    texto = unicodedata.normalize('NFD', str(valor).strip().lower())
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    m = re.match(r'^(\d{4})-(\d{1,2})', texto)
    if m:
        return (int(m[1]), int(m[2]))
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', texto)
    if m:
        return (int(m[3]), int(m[2]))
    m = re.match(r'^([a-z]+)\.?[\s\-/]+(\d{2}|\d{4})$', texto)
    if m and m[1][:3] in MESES:
        anio = int(m[2])
        return (anio + 2000 if anio < 100 else anio, MESES[m[1][:3]])
    return None


def texto_periodo(p) -> str:
    return f'{MESES_TXT[p[1] - 1]} {p[0] % 100:02d}' if p else '?'


def es_formula(valor) -> bool:
    return isinstance(valor, str) and valor.startswith('=')


def vacio_o_cero(valor) -> bool:
    if valor is None:
        return True
    if isinstance(valor, bool):
        return False
    if isinstance(valor, (int, float)):
        return valor == 0
    texto = str(valor).strip()
    return texto == '' or re.fullmatch(r'0+([.,]0+)?', texto) is not None


def a_numero(valor):
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor) if math.isfinite(valor) else None
    texto = str(valor or '').strip().replace('%', '').replace(' ', '')
    if not texto:
        return None
    if ',' in texto and '.' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    else:
        texto = texto.replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return None


def iguales(a, b) -> bool:
    if es_formula(a) or es_formula(b):
        return isinstance(a, str) and isinstance(b, str) and a.replace(' ', '') == b.replace(' ', '')
    na, nb = a_numero(a), a_numero(b)
    if na is not None and nb is not None:
        return abs(na - nb) <= max(1e-9, abs(nb) * 1e-12)
    return str(a if a is not None else '').strip() == str(b if b is not None else '').strip()


def rango(a1: str) -> str:
    return urllib.parse.quote(f"'{HOJA}'!{a1}", safe='')


# -------------------------------------------------------------- cliente API
def _con_reintentos(sesion) -> None:
    """Google limita las lecturas por minuto (error 429). En vez de fallar,
    espera y reintenta: 2, 4, 8, 16, 32 y 60 segundos."""
    import time
    if getattr(sesion, '_con_reintentos', False):
        return
    for metodo in ('get', 'post'):
        original = getattr(sesion, metodo)

        def envoltura(*args, _original=original, **kwargs):
            espera = 2
            for intento in range(7):
                respuesta = _original(*args, **kwargs)
                codigo = getattr(respuesta, 'status_code', 200)
                if codigo not in (429, 500, 502, 503) or intento == 6:
                    return respuesta
                time.sleep(espera)
                espera = min(espera * 2, 60)
            return respuesta

        setattr(sesion, metodo, envoltura)
    try:
        sesion._con_reintentos = True
    except AttributeError:
        pass

class ClienteSheets:
    def __init__(self, sesion=None):
        if sesion is None:
            if not os.environ.get('GOOGLE_APPLICATION_CREDENTIALS'):
                raise RuntimeError('Falta configurar GOOGLE_APPLICATION_CREDENTIALS')
            import google.auth
            import requests
            from google.auth.transport.requests import Request
            credenciales, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets'])
            credenciales.refresh(Request())
            sesion = requests.Session()
            sesion.headers['Authorization'] = f'Bearer {credenciales.token}'
        self.sesion = sesion
        _con_reintentos(sesion)

    def leer(self, a1: str, render: str = 'UNFORMATTED_VALUE') -> list[list]:
        r = self.sesion.get(f'{API}/values/{rango(a1)}',
                            params={'valueRenderOption': render, 'dateTimeRenderOption': 'SERIAL_NUMBER'},
                            timeout=90)
        r.raise_for_status()
        return r.json().get('values', [])

    def escribir(self, datos: list[dict]) -> None:
        for i in range(0, len(datos), 500):
            r = self.sesion.post(f'{API}/values:batchUpdate',
                                 json={'valueInputOption': 'USER_ENTERED', 'data': datos[i:i + 500]},
                                 timeout=90)
            r.raise_for_status()


# ------------------------------------------------------- mapas Sheet / local
class Mapa:
    """Encabezados (fila 1) y períodos (columna A) de una hoja."""

    def __init__(self, encabezados: list, periodos: list):
        self.encabezados = encabezados
        self.por_clave: dict[str, list[int]] = {}
        for col, titulo in enumerate(encabezados, start=1):
            if titulo not in (None, ''):
                self.por_clave.setdefault(clave(titulo), []).append(col)
        self.fila_por_periodo: dict[tuple, list[int]] = {}
        self.periodo_por_fila: dict[int, tuple] = {}
        for fila, valor in enumerate(periodos, start=1):
            p = periodo_de(valor)
            if p and fila > 1:
                self.fila_por_periodo.setdefault(p, []).append(fila)
                self.periodo_por_fila[fila] = p

    def columna(self, titulo):
        cols = self.por_clave.get(clave(titulo), [])
        return cols[0] if len(cols) == 1 else None, len(cols)

    def fila(self, periodo):
        filas = self.fila_por_periodo.get(periodo, [])
        return filas[0] if len(filas) == 1 else None, len(filas)


def mapa_local(ruta: Path) -> Mapa | None:
    if not ruta or not Path(ruta).is_file():
        return None
    from openpyxl import load_workbook
    libro = load_workbook(ruta, read_only=True, data_only=False)
    try:
        ws = libro[HOJA]
        encabezados = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
        periodos = [fila[0] for fila in ws.iter_rows(min_col=1, max_col=1, values_only=True)]
    finally:
        libro.close()
    return Mapa(encabezados, periodos)


REF = re.compile(r"(?<![A-Za-z0-9_!'\.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")


def traducir_formula(formula: str, local: Mapa, sheet: Mapa) -> str:
    """Traduce referencias A1 del Excel local a la ubicación real en el Sheet."""
    if '!' in formula:
        raise ValueError('la fórmula referencia otra hoja')

    def reemplazo(m):
        col_local = numero_columna(m[2])
        fila_local = int(m[4])
        titulo = local.encabezados[col_local - 1] if col_local <= len(local.encabezados) else None
        col_sheet, n = sheet.columna(titulo) if titulo not in (None, '') else (None, 0)
        if col_sheet is None:
            raise ValueError(f'la columna {m[2]} ({titulo!r}) no existe en el Sheet')
        if fila_local == 1:
            fila_sheet = 1
        else:
            p = local.periodo_por_fila.get(fila_local)
            fila_sheet, _ = sheet.fila(p) if p else (None, 0)
            if fila_sheet is None:
                raise ValueError(f'la fila {fila_local} del Excel no tiene período equivalente en el Sheet')
        return f'{m[1]}{letra(col_sheet)}{m[3]}{fila_sheet}'

    return REF.sub(reemplazo, formula)


# ------------------------------------------------------------ escritura
def escribir_actualizaciones(cargas: list[dict], base_local: Path | None = None,
                             cliente: ClienteSheets | None = None) -> dict:
    """cargas: [{'fuente', 'hoja', 'periodo', 'valores'}]. Solo hoja BASE DE DATOS."""
    resumen = {'estado': 'OK', 'por_fuente': {}, 'errores': []}
    cargas = [c for c in cargas if c.get('hoja', HOJA) == HOJA and c.get('valores')]
    if not cargas:
        return resumen
    cliente = cliente or ClienteSheets()
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    periodos = [f[0] if f else None for f in cliente.leer('A:A')]
    sheet = Mapa(encabezados, periodos)
    local = mapa_local(base_local) if base_local else None

    def info(fuente):
        return resumen['por_fuente'].setdefault(fuente, {
            'periodos': [], 'escritas': 0, 'ya_estaban': 0, 'esperadas': 0, 'conservadas': [],
            'no_encontradas': [], 'errores': [], 'cambios': []})

    for carga in cargas:
        fuente = carga.get('fuente') or 'sin_fuente'
        datos_fuente = info(fuente)
        try:
            p = periodo_de(carga.get('periodo'))
            fila, n = sheet.fila(p) if p else (None, 0)
            if fila is None:
                raise ValueError(f"período {carga.get('periodo')!r}: {n} filas en PERIODO")
            if texto_periodo(p) not in datos_fuente['periodos']:
                datos_fuente['periodos'].append(texto_periodo(p))
            valores = dict(carga['valores'])
            # ICC: acumulado anual compuesto (antes lo calculaba n8n).
            if any(clave(t) == clave(ICC_MENSUAL) for t in valores) and not any(clave(t) == clave(ICC_ACUMULADO) for t in valores):
                valores[ICC_ACUMULADO] = ('__ICC__', next(v for t, v in valores.items() if clave(t) == clave(ICC_MENSUAL)))
            datos_fuente['esperadas'] += sum(1 for v in valores.values() if v is not None)

            destinos = {}
            for titulo, valor in valores.items():
                col, n = sheet.columna(titulo)
                if col is None:
                    datos_fuente['no_encontradas'].append(f'{" ".join(str(titulo).split())} ({n} coincidencias)')
                    continue
                destinos[col] = (titulo, valor)
            if not destinos:
                continue
            c1, c2 = min(destinos), max(destinos)
            actual = cliente.leer(f'{letra(c1)}{fila}:{letra(c2)}{fila}', render='FORMULA')
            actual = actual[0] if actual else []
            arriba = []
            if any(es_formula(v) for v in valores.values()) and fila > 2:
                arriba = cliente.leer(f'{letra(c1)}{fila - 1}:{letra(c2)}{fila - 1}', render='FORMULA')
                arriba = arriba[0] if arriba else []

            def existente(col):
                i = col - c1
                return actual[i] if i < len(actual) else None

            escribir, cambios = [], []
            for col, (titulo, valor) in sorted(destinos.items()):
                previo = existente(col)
                celda = f'{letra(col)}{fila}'
                if isinstance(valor, tuple) and valor[0] == '__ICC__':
                    mensual = a_numero(valor[1])
                    if mensual is None:
                        continue
                    if p[1] == 1:
                        nuevo = round(mensual, 2)
                    else:
                        fila_ant, _ = sheet.fila((p[0], p[1] - 1))
                        ant = cliente.leer(f'{letra(col)}{fila_ant}', render='UNFORMATTED_VALUE') if fila_ant else []
                        acum_ant = a_numero(ant[0][0]) if ant and ant[0] else None
                        if acum_ant is None:
                            datos_fuente['errores'].append('ICC: falta el acumulado del mes anterior')
                            continue
                        nuevo = round(((1 + acum_ant / 100) * (1 + mensual / 100) - 1) * 100, 2)
                    valor = nuevo
                if valor is None:
                    continue  # sin dato en origen: no se toca
                if es_formula(valor):
                    if local is None:
                        datos_fuente['errores'].append(f'{celda}: no hay Excel local para traducir la fórmula')
                        continue
                    try:
                        valor = traducir_formula(valor, local, sheet)
                    except ValueError as error:
                        datos_fuente['errores'].append(f'{celda}: {error}')
                        continue
                    if iguales(previo, valor):
                        datos_fuente['ya_estaban'] += 1
                        continue
                    if not vacio_o_cero(previo) and not es_formula(previo) and not carga.get('permitir_reemplazos'):
                        # Columna calculada (el mes anterior del Sheet también tiene fórmula): el número
                        # pegado a mano se reemplaza por la fórmula, así el cálculo sigue la serie.
                        i_arriba = col - c1
                        if not (i_arriba < len(arriba) and es_formula(arriba[i_arriba])):
                            datos_fuente['conservadas'].append(f'{celda} (dato manual)')
                            continue
                else:
                    if iguales(previo, valor):
                        datos_fuente['ya_estaban'] += 1
                        continue
                    if es_formula(previo):
                        datos_fuente['conservadas'].append(f'{celda} (fórmula del Sheet)')
                        continue
                    if valor == '' and vacio_o_cero(previo):
                        continue
                escribir.append({'range': f"'{HOJA}'!{celda}", 'values': [[valor]]})
                cambios.append({'celda': celda, 'titulo': ' '.join(str(titulo).split()), 'antes': previo, 'despues': valor})

            # Protección: una actualización normal agrega datos nuevos o corrige pocos.
            # Si fuera a reemplazar muchos datos ya cargados, algo no cuadra
            # (fila o columna equivocada): no se escribe nada y se avisa.
            pisadas = [c for c in cambios if not es_formula(c['despues'])
                       and not vacio_o_cero(c['antes']) and not es_formula(c['antes'])]
            datos_carga = sum(1 for v in valores.values() if v not in (None, '') and not es_formula(v))
            if not carga.get('permitir_reemplazos') and (len(pisadas) > LIMITE_REEMPLAZOS or (len(pisadas) >= 3 and len(pisadas) * 2 > datos_carga)):
                raise RuntimeError(f"se iban a reemplazar {len(pisadas)} datos ya cargados "
                                   f"(p. ej. {pisadas[0]['celda']}: {pisadas[0]['antes']} -> {pisadas[0]['despues']}). "
                                   "No se escribió nada por seguridad; revisar la fuente.")
            if escribir:
                cliente.escribir(escribir)
                leido = cliente.leer(f'{letra(c1)}{fila}:{letra(c2)}{fila}', render='FORMULA')
                leido = leido[0] if leido else []
                for cambio in cambios:
                    col = numero_columna(re.match(r'[A-Z]+', cambio['celda'])[0])
                    obtenido = leido[col - c1] if col - c1 < len(leido) else None
                    if cambio['despues'] == '' and obtenido in (None, ''):
                        continue
                    if not iguales(obtenido, cambio['despues']):
                        raise RuntimeError(f"verificación fallida en {cambio['celda']}: se leyó {obtenido!r}")
            datos_fuente['escritas'] += len(cambios)
            datos_fuente['cambios'].extend(cambios)
        except Exception as error:  # una carga fallida no frena a las demás
            datos_fuente['errores'].append(str(error))
    for fuente, datos_fuente in resumen['por_fuente'].items():
        if datos_fuente['errores']:
            resumen['errores'].append(fuente)
    if resumen['errores']:
        resumen['estado'] = 'ERROR'
    return resumen


def describir(datos: dict) -> str:
    """Texto corto para Telegram."""
    periodos = ', '.join(datos.get('periodos') or []) or 's/período'
    partes = []
    if datos.get('escritas'):
        partes.append(f"✅ Sheet actualizado ({periodos}): {datos['escritas']} celdas nuevas")
    else:
        partes.append(f"Sin novedades: {periodos} ya estaba completo en el Sheet")
    if datos.get('no_encontradas'):
        partes.append('⚠️ Columnas no encontradas: ' + '; '.join(datos['no_encontradas'][:4]))
    if datos.get('conservadas'):
        partes.append(f"{len(datos['conservadas'])} celdas conservadas (fórmula o dato manual)")
    if datos.get('errores'):
        partes.append('❌ ' + '; '.join(datos['errores'][:3]))
    return '. '.join(partes)


# ----------------------------------------------- el Sheet manda: espejo local
def espejar_desde_sheet(base_local: Path, cliente: ClienteSheets | None = None) -> dict:
    """Copia al Excel local los datos que hay en el Google Sheet (hoja BASE DE DATOS).

    El Sheet de la plataforma es la referencia. Antes de cada revisión el Excel
    local se pone al día con lo que tiene el Sheet, así:
    * una corrección hecha a mano en el Sheet no se pierde ni se "des-corrige";
    * el Excel local deja de ser una copia que hay que mantener a mano;
    * los cálculos que usan meses anteriores parten de los valores publicados.

    Reglas (conservadoras): solo celdas numéricas; no toca fórmulas (ni del Sheet
    ni del Excel); no borra datos locales si en el Sheet la celda está vacía;
    columnas y filas se ubican por encabezado y por período, nunca por posición.
    """
    from openpyxl import load_workbook
    import shutil
    if not base_local or not Path(base_local).is_file():
        return {'estado': 'SIN_BASE', 'celdas': 0}
    cliente = cliente or ClienteSheets()
    valores = cliente.leer('A1:BZZ2000', render='UNFORMATTED_VALUE')
    formulas = cliente.leer('A1:BZZ2000', render='FORMULA')
    if not valores:
        return {'estado': 'SHEET_VACIO', 'celdas': 0}
    sheet = Mapa(valores[0], [f[0] if f else None for f in valores])
    libro = load_workbook(base_local, data_only=False, keep_links=True)
    ws = libro[HOJA]
    cols = {}
    for c in range(1, ws.max_column + 1):
        titulo = ws.cell(1, c).value
        if titulo in (None, '') or clave(titulo) == clave('PERIODO'):
            continue
        col_sheet, n = sheet.columna(titulo)
        if col_sheet:
            cols[c] = col_sheet
    cambios = []
    for f in range(2, ws.max_row + 1):
        p = periodo_de(ws.cell(f, 1).value)
        fila_sheet, n = sheet.fila(p) if p else (None, 0)
        if not fila_sheet:
            continue
        fila_v = valores[fila_sheet - 1] if fila_sheet - 1 < len(valores) else []
        fila_f = formulas[fila_sheet - 1] if fila_sheet - 1 < len(formulas) else []
        for c, cs in cols.items():
            v = fila_v[cs - 1] if cs - 1 < len(fila_v) else None
            fo = fila_f[cs - 1] if cs - 1 < len(fila_f) else None
            if es_formula(fo) or isinstance(v, bool) or not isinstance(v, (int, float)):
                continue
            celda = ws.cell(f, c)
            actual = celda.value
            if es_formula(actual):
                continue
            if v == 0:
                continue  # en el Sheet 0 equivale a "sin dato": no borra lo local
            if isinstance(actual, (int, float)) and not isinstance(actual, bool) and math.isclose(float(actual), float(v), rel_tol=0, abs_tol=1e-9):
                continue
            cambios.append({'celda': celda.coordinate, 'periodo': texto_periodo(p), 'titulo': ' '.join(str(ws.cell(1, c).value).split())[:80],
                            'antes': actual, 'despues': v})
            celda.value = v
    respaldo = None
    if cambios:
        carpeta = Path(base_local).parent / 'respaldos'
        carpeta.mkdir(parents=True, exist_ok=True)
        respaldo = carpeta / f"antes_espejo_sheet_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(base_local, respaldo)
        libro.save(base_local)
    libro.close()
    return {'estado': 'OK', 'celdas': len(cambios), 'columnas_comparadas': len(cols),
            'muestra': cambios[:15], 'respaldo': str(respaldo) if respaldo else None}
