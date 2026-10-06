"""Control de la hoja BASE DE DATOS del Google Sheet (corre al inicio de cada revisión).

El equipo puede seguir cargando valores a mano y completando los cálculos de la
tabla (acumulado, var. intermensual, var. interanual, etc.). Lo que se controla
son los cálculos FUERA de lo estipulado:

1. En filas de meses futuros (todavía no hay datos oficiales) y en columnas sin
   encabezado o filas sin PERIODO (fuera de la tabla): se guarda una copia
   (celda, fórmula y valor) en reportes/control_base/, se BORRA (0 en meses
   futuros, vacío fuera de la tabla) y se avisa por Telegram. En meses futuros
   se aceptan 0/vacío y las fórmulas de columna calculada (miran su fila y hasta
   12 meses atrás y dan 0, vacío o error por falta de datos).
2. Dentro de la tabla (meses con datos): una fórmula que mira filas que no
   corresponden (más de 12 meses atrás, más de 1 mes adelante) u otra pestaña es
   un cálculo fuera de lo estipulado: se AVISA con la celda y la fórmula, pero
   no se toca (puede ser una carga legítima; la decide la responsable). Las
   fórmulas que ya existían la primera vez quedan como línea de base y no se
   vuelven a avisar.
3. La hoja no queda bloqueada (el equipo tiene que poder editar). Si existe la
   protección que se había creado el 5/10/2026, se quita.

Devuelve un resumen que el aviso de Telegram muestra.
"""
from __future__ import annotations

import json
import os
import re
from datetime import date, datetime
from pathlib import Path

from sheet_directo import API, HOJA, ClienteSheets, es_formula, letra, periodo_de, texto_periodo, vacio_o_cero

RAIZ = Path(__file__).resolve().parents[2]
CARPETA_REPORTES = RAIZ / 'reportes' / 'control_base'
RANGO = 'A1:BZZ2000'

# Cuentas que pueden editar BASE DE DATOS además de la automatización.
EDITORES_BASE = ['nunezanna789@gmail.com']
DESCRIPCION_PROTECCION = 'Observatorio: solo la automatización y la responsable editan BASE DE DATOS'

# Si aparecen más hallazgos que esto, no se borra nada (puede ser un cambio de
# estructura de la hoja): solo se avisa.
LIMITE_LIMPIEZA = 150
LINEA_BASE = CARPETA_REPORTES / 'linea_base_formulas.json'
PROTEGER = False  # True = solo la automatización y EDITORES_BASE editan la hoja

REF = re.compile(r"(?<![A-Za-z0-9_!'\.])\$?([A-Z]{1,3})\$?(\d+)(?![A-Za-z0-9_(])")


def _celda(valores, fila, col):
    f = valores[fila - 1] if fila - 1 < len(valores) else []
    return f[col - 1] if col - 1 < len(f) else None


def _es_error(valor) -> bool:
    return isinstance(valor, str) and valor.startswith('#')


def _formula_propia(formula: str, fila: int) -> bool:
    """Fórmula de columna calculada: solo mira su fila y hasta 12 filas arriba."""
    if '!' in formula:
        return False
    filas = [int(m[2]) for m in REF.finditer(formula)]
    return bool(filas) and all(fila - 12 <= f <= fila for f in filas)


def _formula_estipulada(formula: str, fila: int) -> bool:
    """Cálculo de la tabla: mira su fila, hasta 12 meses atrás o 1 adelante (IERIC S/E)."""
    if '!' in formula:
        return False
    return all(fila - 12 <= int(m[2]) <= fila + 1 for m in REF.finditer(formula))


def revisar(cliente: ClienteSheets, hoy: date | None = None) -> dict:
    hoy = hoy or date.today()
    actual = (hoy.year, hoy.month)
    valores = cliente.leer(RANGO, render='UNFORMATTED_VALUE')
    formulas = cliente.leer(RANGO, render='FORMULA')
    if not valores:
        return {'estado': 'SHEET_VACIO', 'hallazgos': []}
    encabezados = valores[0]
    ult_col = max((i for i, t in enumerate(encabezados, start=1) if t not in (None, '')), default=1)
    periodos = {}
    for fila in range(2, max(len(valores), len(formulas)) + 1):
        p = periodo_de(_celda(valores, fila, 1))
        if p:
            periodos[fila] = p
    ult_fila = max(periodos, default=1)

    hallazgos = []
    n_filas = max(len(valores), len(formulas))
    for fila in range(1, n_filas + 1):
        n_cols = max(len(valores[fila - 1]) if fila - 1 < len(valores) else 0,
                     len(formulas[fila - 1]) if fila - 1 < len(formulas) else 0)
        p = periodos.get(fila)
        for col in range(1, n_cols + 1):
            v = _celda(valores, fila, col)
            fo = _celda(formulas, fila, col)
            if v in (None, '') and fo in (None, ''):
                continue
            motivo = None
            if col > ult_col:
                motivo = 'columna sin encabezado'
            elif fila > ult_fila:
                motivo = 'fila sin PERIODO (debajo de la tabla)'
            elif fila > 1 and col > 1 and p and p <= actual and es_formula(fo) and not _formula_estipulada(fo, fila):
                motivo = f'cálculo fuera de lo estipulado ({texto_periodo(p)})'
            elif fila > 1 and col > 1 and p and p > actual:
                if es_formula(fo):
                    if not (_formula_propia(fo, fila) and (vacio_o_cero(v) or _es_error(v))):
                        motivo = f'cálculo en un mes futuro ({texto_periodo(p)})'
                elif not vacio_o_cero(v):
                    motivo = f'valor en un mes futuro ({texto_periodo(p)})'
            if motivo:
                hallazgos.append({
                    'celda': f'{letra(col)}{fila}',
                    'encabezado': encabezados[col - 1] if col - 1 < len(encabezados) else None,
                    'periodo': texto_periodo(p) if p else None,
                    'formula': fo if es_formula(fo) else None,
                    'valor': v,
                    'motivo': motivo,
                    'limpiar_con': None if motivo.startswith('cálculo fuera de lo estipulado') else (
                        0 if (p and col <= ult_col and fila <= ult_fila) else ''),
                })
    return {'estado': 'OK', 'hallazgos': hallazgos, 'ultima_columna': letra(ult_col), 'ultima_fila': ult_fila}


def limpiar(cliente: ClienteSheets, hallazgos: list[dict]) -> int:
    datos = [{'range': f"'{HOJA}'!{h['celda']}", 'values': [[h['limpiar_con']]]} for h in hallazgos
             if h.get('limpiar_con') is not None]
    for i in range(0, len(datos), 200):
        cliente.escribir(datos[i:i + 200])
    return len(datos)


# ------------------------------------------------------------- protección
def cuenta_automatizacion(cliente: ClienteSheets) -> str | None:
    ruta = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
    try:
        datos = json.loads(Path(ruta).read_text(encoding='utf-8')) if ruta else {}
        if datos.get('client_email'):
            return datos['client_email'].lower()
    except (OSError, ValueError):
        pass
    try:
        token = cliente.sesion.headers.get('Authorization', '').replace('Bearer ', '')
        r = cliente.sesion.get('https://oauth2.googleapis.com/tokeninfo', params={'access_token': token}, timeout=30)
        if r.ok and r.json().get('email'):
            return r.json()['email'].lower()
    except Exception:  # noqa: BLE001
        pass
    return None


def asegurar_proteccion(cliente: ClienteSheets) -> dict:
    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(sheetId,title),protectedRanges)'}, timeout=60)
    r.raise_for_status()
    hoja = next((s for s in r.json().get('sheets', []) if s['properties']['title'] == HOJA), None)
    if not hoja:
        return {'estado': 'ERROR', 'detalle': f'no existe la hoja {HOJA}'}
    sheet_id = hoja['properties']['sheetId']
    propia = cuenta_automatizacion(cliente)
    editores = sorted({e.lower() for e in EDITORES_BASE} | ({propia} if propia else set()))
    existentes = [p for p in hoja.get('protectedRanges', [])
                  if p.get('description') == DESCRIPCION_PROTECCION]
    if existentes:
        prot = existentes[0]
        actuales = sorted(u.lower() for u in prot.get('editors', {}).get('users', []))
        faltan = [e for e in editores if e not in actuales]
        sobran = [e for e in actuales if e not in editores]
        if not faltan and not sobran and not prot.get('warningOnly'):
            return {'estado': 'OK', 'accion': 'sin_cambios', 'editores': actuales}
        pedido = {'updateProtectedRange': {
            'protectedRange': {'protectedRangeId': prot['protectedRangeId'], 'warningOnly': False,
                               'editors': {'users': editores, 'domainUsersCanEdit': False}},
            'fields': 'warningOnly,editors'}}
        accion = 'editores_corregidos'
        rid = prot['protectedRangeId']
    else:
        pedido = {'addProtectedRange': {'protectedRange': {
            'range': {'sheetId': sheet_id}, 'description': DESCRIPCION_PROTECCION, 'warningOnly': False,
            'editors': {'users': editores, 'domainUsersCanEdit': False}}}}
        accion = 'creada'
        rid = None
    r = cliente.sesion.post(f'{API}:batchUpdate', json={'requests': [pedido]}, timeout=60)
    r.raise_for_status()
    if rid is None:
        rid = r.json()['replies'][0]['addProtectedRange']['protectedRange']['protectedRangeId']
    # Comprobación: la automatización tiene que poder seguir escribiendo.
    try:
        titulo = (cliente.leer('A1', render='FORMULA') or [['PERIODO']])[0][0]
        cliente.escribir([{'range': f"'{HOJA}'!A1", 'values': [[titulo]]}])
    except Exception as error:  # noqa: BLE001
        cliente.sesion.post(f'{API}:batchUpdate', json={'requests': [{'deleteProtectedRange': {'protectedRangeId': rid}}]},
                            timeout=60)
        return {'estado': 'ERROR', 'detalle': f'la automatización no podía escribir con la protección; se quitó ({error})'}
    return {'estado': 'OK', 'accion': accion, 'editores': editores, 'cuenta_automatizacion': propia}


def quitar_proteccion(cliente: ClienteSheets) -> dict:
    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(title),protectedRanges(protectedRangeId,description))'},
                           timeout=60)
    r.raise_for_status()
    ids = [p['protectedRangeId'] for s in r.json().get('sheets', []) if s['properties']['title'] == HOJA
           for p in s.get('protectedRanges', []) if p.get('description') == DESCRIPCION_PROTECCION]
    if not ids:
        return {'estado': 'OK', 'accion': 'sin_proteccion'}
    r = cliente.sesion.post(f'{API}:batchUpdate', json={'requests': [{'deleteProtectedRange': {'protectedRangeId': i}} for i in ids]},
                            timeout=60)
    r.raise_for_status()
    return {'estado': 'OK', 'accion': 'quitada'}


def _linea_base(hallazgos: list[dict]) -> list[dict]:
    """Primera vez: guarda las fórmulas fuera de patrón que ya existían y no las avisa."""
    dentro = [h for h in hallazgos if h.get('limpiar_con') is None]
    clave = lambda h: f"{h['celda']}|{h.get('formula')}"  # noqa: E731
    if not LINEA_BASE.exists():
        CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
        LINEA_BASE.write_text(json.dumps(sorted(clave(h) for h in dentro), ensure_ascii=False, indent=1), encoding='utf-8')
        conocidas = {clave(h) for h in dentro}
    else:
        conocidas = set(json.loads(LINEA_BASE.read_text(encoding='utf-8')))
    return [h for h in hallazgos if h.get('limpiar_con') is not None or clave(h) not in conocidas]


# ------------------------------------------------------------- principal
def controlar(cliente: ClienteSheets | None = None, limpiar_hallazgos: bool = True, proteger: bool = True) -> dict:
    cliente = cliente or ClienteSheets()
    salida = revisar(cliente)
    salida['hallazgos'] = hallazgos = _linea_base(salida.get('hallazgos', []))
    salida['limpiadas'] = 0
    if hallazgos:
        CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
        ruta = CARPETA_REPORTES / f'control_{datetime.now():%Y%m%d_%H%M%S}.json'
        ruta.write_text(json.dumps(hallazgos, ensure_ascii=False, indent=1, default=str), encoding='utf-8')
        salida['respaldo'] = str(ruta)
        borrables = [h for h in hallazgos if h.get('limpiar_con') is not None]
        if limpiar_hallazgos and borrables and len(borrables) <= LIMITE_LIMPIEZA:
            salida['limpiadas'] = limpiar(cliente, borrables)
        elif limpiar_hallazgos:
            salida['detalle'] = f'{len(hallazgos)} hallazgos: demasiados para borrar solos; revisar a mano.'
    if proteger:
        try:
            salida['proteccion'] = asegurar_proteccion(cliente) if PROTEGER else quitar_proteccion(cliente)
        except Exception as error:  # noqa: BLE001
            salida['proteccion'] = {'estado': 'ERROR', 'detalle': str(error)[:300]}
    return salida


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(description='Control de la hoja BASE DE DATOS')
    ap.add_argument('--solo-revisar', action='store_true', help='No borra ni protege; solo informa.')
    a = ap.parse_args()
    try:
        res = controlar(limpiar_hallazgos=not a.solo_revisar, proteger=not a.solo_revisar)
    except Exception as error:  # noqa: BLE001
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        raise SystemExit(1)
    res['fuente'] = 'Control de la hoja BASE DE DATOS'
    res['periodo'] = date.today().strftime('%Y-%m')
    print(json.dumps(res, ensure_ascii=False, indent=1, default=str))
    print('ESTADO=OK')
