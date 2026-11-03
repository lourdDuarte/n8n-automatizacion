"""Escritura directa y verificada de los indicadores productivos en el Sheet.

Pestaña "BASE DE DATOS - Indicadores Pro..." (gid 1954947101). Tiene dos
bloques con formas distintas de ubicar el período:

* Comercialización de granos: columna "Año/Mes" con una fila de año (2026)
  seguida de las filas de meses (ENERO, FEBRERO, ...).
* Actividad pecuaria: columna "Periodo" con textos como "ago-26".

Las filas del Excel local y del Sheet NO coinciden en número, por eso cada
fila se ubica por su período y cada columna por su encabezado. Las fórmulas
del Excel local se traducen a las filas y columnas reales del Sheet.

Además repara el bloque de comercialización: compara los últimos 18 meses
contra el Excel local y corrige las celdas que el camino anterior por n8n
había escrito en filas equivocadas.
"""
from __future__ import annotations

import re
import unicodedata
import urllib.parse
from datetime import date, datetime
from pathlib import Path

from sheet_directo import (API, ClienteSheets, clave, es_formula, iguales, letra,
                           numero_columna, vacio_o_cero)

GID = 1954947101
HOJA_LOCAL = 'BASE DE DATOS - Indicadores Pro'
FILA_ENCABEZADO_SHEET = 3
FILA_ENCABEZADO_LOCAL = 2
MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
         'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
MESES_CORTOS = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6, 'jul': 7,
                'ago': 8, 'sep': 9, 'set': 9, 'oct': 10, 'nov': 11, 'dic': 12}
COLUMNAS_COMERCIALIZACION = 16  # C..R del Excel local: 8 precios + 8 variaciones
REPARAR_MESES = 18
REPARAR_MESES_PECUARIA = 12
# Columnas que solo existen en el Excel local (cálculos internos): no van al Sheet
# y no cuentan como faltantes.
SOLO_LOCAL = {clave('Consumo Total de Proteina de origen animal / Kg N')}

# Formato de cada columna pecuaria tal como está cargada históricamente en el Sheet,
# para que los datos nuevos sigan el mismo formato que toda la columna:
#   'coma2'  -> texto con coma decimal y 2 decimales ("47,08")
#   'miles3' -> número en miles con 3 decimales (784.238 = 784.238 cabezas)
#   'puntos' -> texto con punto de miles ("5.387.190")
#   'entero' -> número entero
FORMATOS_PECUARIA = {clave(t): f for t, f in {
    'Stock Fsa (cabezas) Bovinos': 'entero',
    'Stock Nación (cabezas) Bovinos': 'entero',
    'Faena Fsa (cabezas) Bovinos': 'entero',
    'Faena Nación (cabezas) Bovinos': 'entero',
    'Consumo per cápita N (kg/hab/año) Bovinos': 'coma2',
    'Stock Fsa (cabezas) Porcinos': 'miles3',
    'Stock N (cabezas) Porcinos': 'puntos',
    'Faena Total N (cabezas) Porcinos': 'miles3',
    'Consumo per cápita N (kg/hab/año) Porcinos': 'coma2',
    'Produccion Destino Industria (miles de huevos) N Aves': 'miles3',
    'Produccion/ Carne (tn)': 'miles3',
    'Faena (miles cab) N Aves': 'miles3',
    'Consumo per cápita N (kg/hab/año) Aves': 'coma2',
}.items()}


def a_formato_sheet(valor, formato):
    """Convierte el dato original (unidades) al formato histórico de la columna."""
    if formato is None or valor is None or es_formula(valor) or isinstance(valor, str):
        return valor
    v = float(valor)
    if formato == 'coma2':
        return f'{v:.2f}'.replace('.', ',')
    if formato == 'miles3':
        return round(v / 1000, 3)
    if formato == 'puntos':
        return f'{int(round(v)):,}'.replace(',', '.')
    if formato == 'entero':
        return int(round(v))
    return valor


def desde_formato_sheet(valor, formato):
    """Dato en unidades a partir de lo que muestra el Sheet (para comparar)."""
    if valor in (None, '') or es_formula(valor):
        return None
    try:
        if isinstance(valor, str):
            t = valor.strip()
            if formato == 'puntos' or re.fullmatch(r'\d{1,3}(\.\d{3})+', t):
                return float(t.replace('.', ''))
            t = t.replace('.', '').replace(',', '.') if ',' in t else t
            v = float(t)
        else:
            v = float(valor)
    except ValueError:
        return None
    return v * 1000 if formato == 'miles3' else v
LIMITE_REEMPLAZOS = 12


def _texto(valor) -> str:
    t = unicodedata.normalize('NFD', str(valor if valor is not None else ''))
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    return ' '.join(t.split()).upper()


def _anio(valor):
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)) and 2000 <= valor <= 2100 and float(valor).is_integer():
        return int(valor)
    t = _texto(valor)
    return int(t) if re.fullmatch(r'20\d{2}', t) else None


def _mes(valor):
    t = _texto(valor).replace('SETIEMBRE', 'SEPTIEMBRE')
    return MESES.index(t) + 1 if t in MESES else None


def periodo_pecuario(valor, local: bool):
    """Sheet: 'ago-26' / 'AGO 26'. Local: fecha con día = año (2026-08-26 -> ago-26)."""
    if valor in (None, ''):
        return None
    # Las fechas se muestran con formato mmm-aa: vale el año y el mes de la fecha.
    if isinstance(valor, (datetime, date)):
        return (valor.year, valor.month)
    if isinstance(valor, (int, float)) and not isinstance(valor, bool) and 20000 < valor < 80000:
        from datetime import timedelta
        f = date(1899, 12, 30) + timedelta(days=int(valor))
        return (f.year, f.month)
    t = _texto(valor).lower()
    m = re.match(r'^([a-z]+)\.?[\s\-/]+(\d{2}|\d{4})$', t)
    if m and m.group(1)[:3] in MESES_CORTOS:
        anio = int(m.group(2))
        return (anio + 2000 if anio < 100 else anio, MESES_CORTOS[m.group(1)[:3]])
    m = re.match(r'^(\d{4})-(\d{1,2})', t)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return None


class Grilla:
    """Encabezados + claves de fila (comercialización y pecuaria) de una hoja."""

    def __init__(self, filas: list[list], fila_encabezado: int, local: bool):
        self.filas = filas
        # La fila de encabezados se detecta buscando títulos conocidos: no se
        # adivina por número de fila.
        buscados = {clave('Precio Nacional ($/Tn) Maiz'), clave('Periodo')}
        detectada = next((n for n, fila in enumerate(filas[:15], start=1)
                          if sum(1 for v in fila if clave(v) in buscados) >= 2), None)
        if detectada is None:
            muestra = [[str(v)[:25] for v in f[:4]] for f in filas[:6]]
            raise RuntimeError(f'No se encontró la fila de encabezados de la pestaña productiva. Primeras filas: {muestra}')
        fila_encabezado = detectada
        encabezados = filas[fila_encabezado - 1] if len(filas) >= fila_encabezado else []
        self.encabezados = list(encabezados)
        self.por_clave = {}
        for col, titulo in enumerate(self.encabezados, start=1):
            if titulo not in (None, ''):
                self.por_clave.setdefault(clave(titulo), []).append(col)
        anio_mes = self.por_clave.get(clave('Año/Mes'), [1])[0]
        periodo = self.por_clave.get(clave('Periodo'), [None])[0]
        self.col_periodo = periodo
        self.comercial, self.pecuaria = {}, {}
        self.clave_comercial, self.clave_pecuaria = {}, {}
        anio_actual = None
        for n, fila in enumerate(filas, start=1):
            if n <= fila_encabezado:
                continue
            valor = fila[anio_mes - 1] if len(fila) >= anio_mes else None
            if _anio(valor):
                anio_actual = _anio(valor)
            elif anio_actual and _mes(valor):
                k = (anio_actual, _mes(valor))
                self.comercial.setdefault(k, []).append(n)
                self.clave_comercial[n] = k
            if periodo and len(fila) >= periodo:
                p = periodo_pecuario(fila[periodo - 1], local)
                if p:
                    self.pecuaria.setdefault(p, []).append(n)
                    self.clave_pecuaria[n] = p

    def columna(self, titulo):
        cols = self.por_clave.get(clave(titulo), [])
        return (cols[0] if len(cols) == 1 else None), len(cols)

    def celda(self, fila, col):
        f = self.filas[fila - 1] if fila - 1 < len(self.filas) else []
        return f[col - 1] if col - 1 < len(f) else None


def _bloque(titulos, local: Grilla) -> str:
    cols = [local.columna(t)[0] for t in titulos]
    cols = [c for c in cols if c]
    return 'pecuaria' if cols and local.col_periodo and min(cols) >= local.col_periodo else 'comercializacion'


REF = re.compile(r"(?<![A-Za-z0-9_!'\.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")


def _traducir(formula, local: Grilla, sheet: Grilla, bloque: str):
    if '!' in formula:
        raise ValueError('la fórmula referencia otra hoja')
    claves_local = local.clave_comercial if bloque == 'comercializacion' else local.clave_pecuaria
    filas_sheet = sheet.comercial if bloque == 'comercializacion' else sheet.pecuaria

    def reemplazo(m):
        col_local = numero_columna(m.group(2))
        titulo = local.encabezados[col_local - 1] if col_local <= len(local.encabezados) else None
        col_sheet, _ = sheet.columna(titulo) if titulo not in (None, '') else (None, 0)
        if col_sheet is None:
            raise ValueError(f'la columna {m.group(2)} ({titulo!r}) no existe en el Sheet')
        k = claves_local.get(int(m.group(4)))
        destino = filas_sheet.get(k, []) if k else []
        if len(destino) != 1:
            raise ValueError(f'la fila {m.group(4)} del Excel no tiene período equivalente en el Sheet')
        return f'{m.group(1)}{letra(col_sheet)}{m.group(3)}{destino[0]}'

    return REF.sub(reemplazo, formula)


def _titulo_hoja(cliente: ClienteSheets) -> str:
    r = cliente.sesion.get(API, params={'fields': 'sheets.properties(sheetId,title)'}, timeout=60)
    r.raise_for_status()
    for hoja in r.json()['sheets']:
        if hoja['properties']['sheetId'] == GID:
            return hoja['properties']['title']
    raise RuntimeError(f'No se encontró la pestaña productiva (gid {GID}) en el Sheet')


def _leer(cliente, titulo, a1, render):
    rango = urllib.parse.quote(f"'{titulo}'!{a1}", safe='')
    r = cliente.sesion.get(f'{API}/values/{rango}', params={'valueRenderOption': render,
                                                          'dateTimeRenderOption': 'SERIAL_NUMBER'}, timeout=90)
    r.raise_for_status()
    return r.json().get('values', [])


def _mapa_local(base: Path) -> Grilla:
    from openpyxl import load_workbook
    libro = load_workbook(base, read_only=True, data_only=False)
    try:
        filas = [list(f) for f in libro[HOJA_LOCAL].iter_rows(values_only=True)]
    finally:
        libro.close()
    return Grilla(filas, FILA_ENCABEZADO_LOCAL, local=True)


def _cargas_reparacion(local: Grilla) -> list[dict]:
    """Últimos meses de comercialización con dato en el Excel local (fuente de verdad)."""
    claves = sorted(k for k in local.comercial if len(local.comercial[k]) == 1
                    and any(not es_formula(v) and not vacio_o_cero(v)
                            for v in local.filas[local.comercial[k][0] - 1][2:2 + COLUMNAS_COMERCIALIZACION]))
    cargas = []
    titulos = [t for t in local.encabezados[2:2 + COLUMNAS_COMERCIALIZACION] if t not in (None, '')]
    for k in claves[-REPARAR_MESES:]:
        ultimo = k == claves[-1]
        fila = local.comercial[k][0]
        valores = {}
        for t in titulos:
            col, _ = local.columna(t)
            v = local.celda(fila, col) if col else None
            if v not in (None, ''):
                valores[t] = v
        if any(not es_formula(v) and not vacio_o_cero(v) for v in valores.values()):
            # El último mes cuenta como dato nuevo; los anteriores, como reparación.
            # El último mes lo acaba de confirmar la extracción oficial de esta misma revisión:
            # si en el Sheet quedó otra cosa (p. ej. se copió el mes anterior), se corrige.
            cargas.append({'fuente': 'comercializacion', 'row_number': fila, 'valores': valores, 'reparacion': not ultimo,
                           'confirmada_por_fuente': ultimo})
    return cargas


def _cargas_reparacion_pecuaria(local: Grilla) -> list[dict]:
    """Últimos meses de pecuaria del Excel local: completa en el Sheet solo las celdas VACÍAS.

    Así un dato que quedó cargado en el Excel pero no llegó al Sheet (corte, error de red)
    se completa en la próxima revisión, sin pisar nada que ya esté en el Sheet."""
    if not local.col_periodo:
        return []
    titulos = [t for t in local.encabezados[local.col_periodo:] if t not in (None, '') and clave(t) not in SOLO_LOCAL]
    cargas = []
    for k in sorted(local.pecuaria)[-REPARAR_MESES_PECUARIA:]:
        if len(local.pecuaria[k]) != 1:
            continue
        fila = local.pecuaria[k][0]
        valores = {}
        for t in titulos:
            col, _ = local.columna(t)
            v = local.celda(fila, col) if col else None
            if v not in (None, '') and not es_formula(v) and not vacio_o_cero(v):
                valores[t] = v
        if valores:
            cargas.append({'fuente': 'pecuaria', 'fuente_bloque': 'pecuaria', 'row_number': fila,
                           'valores': valores, 'reparacion': True, 'solo_vacias': True})
    return cargas


def escribir_productivos(cargas: list[dict], base_local: Path, cliente: ClienteSheets | None = None,
                         reparar: bool = True) -> dict:
    resumen = {'estado': 'OK', 'por_fuente': {}, 'errores': []}
    cliente = cliente or ClienteSheets()
    titulo = _titulo_hoja(cliente)
    local = _mapa_local(base_local)
    todas = [dict(c) for c in cargas if c.get('valores')]
    if reparar:
        todas += _cargas_reparacion(local)
        todas += _cargas_reparacion_pecuaria(local)

    for carga in todas:
        sheet = Grilla(_leer(cliente, titulo, 'A1:BZ400', 'FORMULA'), FILA_ENCABEZADO_SHEET, local=False)
        valores = {t: v for t, v in carga['valores'].items() if clave(t) != clave('Periodo') and clave(t) not in SOLO_LOCAL}
        if not valores:
            continue
        bloque = carga.get('fuente_bloque') or _bloque(valores, local)
        fuente = carga.get('fuente') or bloque
        info = resumen['por_fuente'].setdefault(fuente, {'periodos': [], 'escritas': 0, 'ya_estaban': 0,
                                                         'esperadas': 0, 'reparadas': 0, 'conservadas': [],
                                                         'no_encontradas': [], 'errores': [], 'cambios': []})
        try:
            fila_local = int(carga['row_number'])
            k = (local.clave_comercial if bloque == 'comercializacion' else local.clave_pecuaria).get(fila_local)
            if not k:
                raise ValueError(f'la fila {fila_local} del Excel no tiene período ({bloque})')
            destino = (sheet.comercial if bloque == 'comercializacion' else sheet.pecuaria).get(k, [])
            if len(destino) != 1:
                raise ValueError(f'{MESES[k[1] - 1].lower()} {k[0]}: {len(destino)} filas en el Sheet ({bloque})')
            fila = destino[0]
            etiqueta = f'{MESES[k[1] - 1][:3].lower()} {k[0] % 100:02d}'
            if not carga.get('reparacion') and etiqueta not in info['periodos']:
                info['periodos'].append(etiqueta)
            escribir, cambios = [], []
            if not carga.get('reparacion'):
                info['esperadas'] += sum(1 for v in valores.values() if v is not None)
            for t, valor in valores.items():
                col, n = sheet.columna(t)
                if col is None:
                    if carga.get('reparacion'):
                        continue
                    info['no_encontradas'].append(f'{" ".join(str(t).split())} ({n} coincidencias)')
                    continue
                previo = sheet.celda(fila, col)
                celda = f'{letra(col)}{fila}'
                if valor is None:
                    continue
                if carga.get('solo_vacias') and not vacio_o_cero(previo):
                    continue
                formato = FORMATOS_PECUARIA.get(clave(t)) if bloque == 'pecuaria' else None
                original = valor
                if formato:
                    valor = a_formato_sheet(valor, formato)
                if es_formula(valor):
                    try:
                        valor = _traducir(valor, local, sheet, bloque)
                    except ValueError as error:
                        info['errores'].append(f'{celda}: {error}')
                        continue
                    if iguales(previo, valor):
                        info['ya_estaban'] += 0 if carga.get('reparacion') else 1
                        continue
                    if not vacio_o_cero(previo) and not es_formula(previo):
                        info['conservadas'].append(f'{celda} (dato manual)')
                        continue
                else:
                    igual = (str(previo).strip() == valor) if isinstance(valor, str) else iguales(previo, valor)
                    if igual:
                        info['ya_estaban'] += 0 if carga.get('reparacion') else 1
                        continue
                    if es_formula(previo):
                        info['conservadas'].append(f'{celda} (fórmula del Sheet)')
                        continue
                # Mismo dato con otro formato (p. ej. 784237 escrito como 784.237): se corrige
                # el formato pero no cuenta como reemplazo de un dato.
                mismo = False
                if formato and not es_formula(valor):
                    # El valor anterior pudo quedar en el formato correcto o crudo (en unidades).
                    candidatos = [desde_formato_sheet(previo, formato), desde_formato_sheet(previo, None)]
                    try:
                        mismo = any(x is not None and abs(x - float(original)) <= max(1.0, abs(float(original)) * 0.001)
                                    for x in candidatos)
                    except (TypeError, ValueError):
                        mismo = False
                escribir.append({'range': f"'{titulo}'!{celda}", 'values': [[valor]]})
                cambios.append({'celda': celda, 'titulo': ' '.join(str(t).split()), 'antes': previo,
                                'despues': valor, 'reparacion': bool(carga.get('reparacion')), 'mismo_dato': mismo})
            # Protección: una actualización normal agrega datos nuevos o corrige pocos.
            # Si fuera a reemplazar muchos datos ya cargados, algo no cuadra
            # (fila o columna equivocada): no se escribe nada y se avisa.
            pisadas = [c for c in cambios if not es_formula(c['despues']) and not c.get('mismo_dato')
                       and not vacio_o_cero(c['antes']) and not es_formula(c['antes'])]
            datos_carga = sum(1 for v in valores.values() if v not in (None, '') and not es_formula(v))
            if not carga.get('reparacion') and not carga.get('confirmada_por_fuente') and (len(pisadas) > LIMITE_REEMPLAZOS or (len(pisadas) >= 3 and len(pisadas) * 2 > datos_carga)):
                raise RuntimeError(f"se iban a reemplazar {len(pisadas)} datos ya cargados "
                                   f"(p. ej. {pisadas[0]['celda']}: {pisadas[0]['antes']} -> {pisadas[0]['despues']}). "
                                   "No se escribió nada por seguridad; revisar la fuente.")
            if escribir:
                r = cliente.sesion.post(f'{API}/values:batchUpdate',
                                        json={'valueInputOption': 'USER_ENTERED', 'data': escribir}, timeout=90)
                r.raise_for_status()
                fila_leida = _leer(cliente, titulo, f'A{fila}:BZ{fila}', 'FORMULA')
                fila_leida = fila_leida[0] if fila_leida else []
                for c in cambios:
                    col = numero_columna(re.match(r'[A-Z]+', c['celda']).group(0))
                    obtenido = fila_leida[col - 1] if col - 1 < len(fila_leida) else None
                    if not iguales(obtenido, c['despues']):
                        raise RuntimeError(f"verificación fallida en {c['celda']}: se leyó {obtenido!r}")
            if not carga.get('reparacion'):  # columnas calculadas (p. ej. variación del maíz)
                try:
                    from extender_formulas import FILAS_ATRAS, planear
                    ini = max(2, fila - FILAS_ATRAS)
                    bloque = _leer(cliente, titulo, f'A{ini}:BZ{fila}', 'FORMULA')
                    cols = {sheet.columna(t)[0] for t, v in valores.items() if v is not None} - {None}
                    plan = planear(bloque, ini, fila, cols)
                    if plan:
                        r = cliente.sesion.post(f'{API}/values:batchUpdate', json={
                            'valueInputOption': 'USER_ENTERED',
                            'data': [{'range': f"'{titulo}'!{a1}", 'values': [[f]]} for a1, f in plan]}, timeout=90)
                        r.raise_for_status()
                        info.setdefault('formulas_extendidas', []).extend(a1 for a1, _ in plan)
                except Exception as error:  # noqa: BLE001 - nunca frena la carga
                    info.setdefault('avisos', []).append(f'no se pudieron extender las fórmulas: {error}')
            if carga.get('reparacion'):
                info['reparadas'] += len(cambios)
            else:
                info['escritas'] += len(cambios)
            info['cambios'].extend(cambios)
        except Exception as error:
            info['errores'].append(str(error))
    for fuente, info in resumen['por_fuente'].items():
        if info['errores']:
            resumen['errores'].append(fuente)
    if resumen['errores']:
        resumen['estado'] = 'ERROR'
    return resumen
