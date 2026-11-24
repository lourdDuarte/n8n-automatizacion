"""Escritura directa de DNRPA en el Google Sheet del Observatorio.

Escribe los 8 datos originales (patentamientos y transferencias de autos y
motos, Formosa y Nación) y completa las 3 columnas calculadas de cada uno
(acumulado, var. % intermensual, var. % interanual) con fórmulas A1, igual
que en el Excel local. No depende del subworkflow de n8n.

Usa la cuenta de servicio de GOOGLE_APPLICATION_CREDENTIALS, que ya tiene
permiso de Editor sobre el Sheet.
"""
from __future__ import annotations

import os
import re
import unicodedata
import urllib.parse
from datetime import date, timedelta

SHEET_ID = '1cVd_KBbtjMN8vUHlUks1IKwPf0c4Oq0RcYn4B5N5PkA'
HOJA = 'BASE DE DATOS'
API = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}'

# Dato original + sus 3 columnas calculadas, en el orden del Sheet.
GRUPOS = (
    ('Patentamiento de automóviles F', 'Patentamiento de automóviles - Acumulado F', 'Patentamiento de automóviles - Var % intermensual F', 'Patentamiento de automóviles - Var % interanual F'),
    ('Transferencias de automoviles F', 'Transferencias de automoviles acumulada F', 'Transferencias de automóviles - Var % intermensual F', 'Transferencias de automóviles - Var % interanual F'),
    ('Patentamiento de automóviles N', 'Patentamiento de automóviles - Acumulado N', 'Patentamiento de automóviles - Var % intermensual N', 'Patentamiento de automóviles - Var % interanual N'),
    ('Transferencias de automoviles N', 'Transferencias de automoviles acumulada N', 'Transferencias de automóviles - Var % intermensual N', 'Transferencias de automóviles - Var % interanual N'),
    ('Patentamiento de Motos F', 'Patentamiento de Motos - Acumulado F', 'Patentamiento de Motos- Var % intermensual F', 'Patentamiento de Motos- Var % interanual F'),
    ('Transferencias de Motos F', 'Transferencias de Motos - Acumulado F', 'Transferencias de Motos- Var % intermensual F', 'Trasnferencias de Motos- Var % interanual F'),
    ('Patentamiento de Motos N', 'Patentamiento de Motos - Acumulado N', 'Patentamiento de Motos- Var % intermensual N', 'Patentamiento de Motos- Var % interanual N'),
    ('Transferencias de Motos N', 'Transferencias de Motos - Acumulado N', 'Transferencias de Motos- Var % intermensual N', 'Trasnferencias de Motos- Var % interanual N'),
)

MESES = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
         'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12}


def clave(texto) -> str:
    simple = unicodedata.normalize('NFD', str(texto or ''))
    simple = ''.join(c for c in simple if unicodedata.category(c) != 'Mn')
    simple = re.sub(r'\s*-\s*', '-', ' '.join(simple.split()).casefold())
    return simple


def letra(numero: int) -> str:
    texto = ''
    while numero:
        numero, resto = divmod(numero - 1, 26)
        texto = chr(65 + resto) + texto
    return texto


def fecha_celda(valor):
    """Convierte el valor de la columna PERIODO a (año, mes)."""
    if valor in (None, ''):
        return None
    if isinstance(valor, (int, float)):
        f = date(1899, 12, 30) + timedelta(days=int(valor))
        return (f.year, f.month)
    texto = str(valor).strip().lower()
    m = re.match(r'^(\d{4})-(\d{1,2})', texto)
    if m:
        return (int(m[1]), int(m[2]))
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', texto)
    if m:
        return (int(m[3]), int(m[2]))
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    m = re.match(r'^([a-z]+)\.?[\s\-/]+(\d{2}|\d{4})$', texto)
    if m and m[1][:3] in MESES:
        anio = int(m[2])
        return (anio + 2000 if anio < 100 else anio, MESES[m[1][:3]])
    return None


def es_vacio_o_cero(valor) -> bool:
    if valor is None:
        return True
    if isinstance(valor, (int, float)):
        return valor == 0
    texto = str(valor).strip().replace('.', '').replace(',', '')
    return texto == '' or re.fullmatch(r'0+', texto) is not None


def rango(a1: str) -> str:
    return urllib.parse.quote(f"'{HOJA}'!{a1}", safe='')


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

    def leer(self, a1: str, render: str = 'UNFORMATTED_VALUE') -> list[list]:
        r = self.sesion.get(f'{API}/values/{rango(a1)}', params={'valueRenderOption': render, 'dateTimeRenderOption': 'SERIAL_NUMBER'}, timeout=60)
        r.raise_for_status()
        return r.json().get('values', [])

    def escribir(self, datos: list[dict]) -> dict:
        r = self.sesion.post(f'{API}/values:batchUpdate', json={'valueInputOption': 'USER_ENTERED', 'data': datos}, timeout=60)
        r.raise_for_status()
        return r.json()


def ubicar_columnas(encabezados: list) -> list[list[int]]:
    """Devuelve, por grupo, los números de columna (1 = A) de las 4 columnas."""
    indice = {}
    for posicion, titulo in enumerate(encabezados, start=1):
        indice.setdefault(clave(titulo), []).append(posicion)
    grupos = []
    for grupo in GRUPOS:
        columnas = []
        for titulo in grupo:
            encontrados = indice.get(clave(titulo), [])
            if len(encontrados) != 1:
                raise RuntimeError(f'Sheet: el encabezado {titulo!r} aparece {len(encontrados)} veces')
            columnas.append(encontrados[0])
        if columnas != list(range(columnas[0], columnas[0] + 4)):
            raise RuntimeError(f'Sheet: las cuatro columnas de {grupo[0]!r} no están contiguas ({[letra(c) for c in columnas]})')
        grupos.append(columnas)
    return grupos


def ubicar_fila(periodos: list[list], anio: int, mes: int) -> int:
    filas = [n for n, fila in enumerate(periodos, start=1) if fila and fecha_celda(fila[0]) == (anio, mes)]
    if len(filas) != 1:
        raise RuntimeError(f'Sheet: se encontraron {len(filas)} filas para {anio}-{mes:02d} en la columna PERIODO')
    return filas[0]


def formulas(col: int, fila: int, mes: int) -> list[str]:
    o, ac = letra(col), letra(col + 1)
    acumulado = f'={o}{fila}' if mes == 1 else f'={ac}{fila - 1}+{o}{fila}'
    return [
        acumulado,
        f'={o}{fila}/{o}{fila - 1}*100-100',
        f'={o}{fila}/{o}{fila - 12}*100-100',
    ]


def escribir_periodo(cliente: ClienteSheets, grupos: list[list[int]], periodos: list[list], anio: int, mes: int, valores: list[int], completar_formulas: bool = True) -> tuple[list[dict], list[dict]]:
    """valores: los 8 datos originales en el orden de GRUPOS.

    Solo completa celdas vacías: un mes ya cargado no se corrige semana a semana
    (las diferencias con lo publicado se informan en el log, sin escribir)."""
    fila = ubicar_fila(periodos, anio, mes)
    primera, ultima = grupos[0][0], grupos[-1][-1]
    actual = cliente.leer(f'{letra(primera)}{fila}:{letra(ultima)}{fila}', render='FORMULA')
    actual = actual[0] if actual else []

    def existente(col):
        pos = col - primera
        return actual[pos] if pos < len(actual) else None

    datos, cambios, diferencias, escritos = [], [], [], []
    for columnas, valor in zip(grupos, valores):
        celda = f'{letra(columnas[0])}{fila}'
        antes = existente(columnas[0])
        if es_vacio_o_cero(antes):
            datos.append({'range': f"'{HOJA}'!{celda}", 'values': [[valor]]})
            cambios.append({'celda': celda, 'antes': antes, 'despues': valor})
            escritos.append((columnas[0], valor))
        elif antes != valor:
            diferencias.append({'celda': celda, 'cargado': antes, 'publicado_hoy': valor})
        if not completar_formulas:
            continue
        for col, formula in zip(columnas[1:], formulas(columnas[0], fila, mes)):
            previo = existente(col)
            if isinstance(previo, str) and previo.startswith('='):
                continue  # ya tiene fórmula: se recalcula sola
            if es_vacio_o_cero(previo):
                celda_f = f'{letra(col)}{fila}'
                datos.append({'range': f"'{HOJA}'!{celda_f}", 'values': [[formula]]})
                cambios.append({'celda': celda_f, 'antes': previo, 'despues': formula})
    if datos:
        cliente.escribir(datos)
    # Verificación: releer los originales escritos.
    if escritos:
        leido = cliente.leer(f'{letra(primera)}{fila}:{letra(ultima)}{fila}', render='UNFORMATTED_VALUE')
        leido = leido[0] if leido else []
        for col, valor in escritos:
            pos = col - primera
            obtenido = leido[pos] if pos < len(leido) else None
            if obtenido != valor:
                raise RuntimeError(f'Sheet: verificación fallida en {letra(col)}{fila}: se esperaba {valor}, se leyó {obtenido!r}')
    return cambios, diferencias


def actualizar_sheet_dnrpa(resultado: dict, nombres_series: list[str], cliente: ClienteSheets | None = None) -> dict:
    """Escribe el último mes cerrado (con fórmulas) solo en celdas vacías."""
    cliente = cliente or ClienteSheets()
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    grupos = ubicar_columnas(encabezados)
    periodos = cliente.leer('A:A')
    anio, mes = resultado['anio'], resultado['mes_numero']
    orden = [('formosa', 'patentamientos_autos'), ('formosa', 'transferencias_autos'), ('nacion', 'patentamientos_autos'), ('nacion', 'transferencias_autos'),
             ('formosa', 'patentamientos_motos'), ('formosa', 'transferencias_motos'), ('nacion', 'patentamientos_motos'), ('nacion', 'transferencias_motos')]
    series = resultado['series']
    salida = {'estado': 'OK', 'periodo': f'{anio}-{mes:02d}', 'cambios': []}
    # Se escribe el mes solo si sus celdas están vacías; ya no se revisa el mes anterior
    # (DNRPA lo ajusta de a poco y eso generaba avisos semanales).
    salida['cambios'], salida['diferencias_no_escritas'] = escribir_periodo(cliente, grupos, periodos, anio, mes, [series[s][k] for k, s in orden])
    # Solo cuentan los datos originales (las fórmulas acompañan): 8 = mes nuevo completo.
    salida['celdas_escritas'] = sum(1 for c in salida['cambios'] if not str(c['despues']).startswith('='))
    if salida['diferencias_no_escritas']:
        print(f"AVISO: DNRPA publica hoy cifras distintas a las cargadas en {salida['periodo']} "
              f"({len(salida['diferencias_no_escritas'])} de 8); no se corrigen automáticamente.")
    # Los 8 datos originales del mes se releen del Sheet y se comparan (si no, escribir_periodo falla).
    salida['datos_esperados'] = 8
    salida['datos_verificados'] = 8
    return salida
