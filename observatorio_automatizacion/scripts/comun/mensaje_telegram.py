"""Arma el aviso de Telegram de cada revisión a partir del resumen de Python.

n8n solo lo reenvía. El texto dice, por indicador, si hubo datos nuevos y
quedaron en el Sheet, si no había novedades, si la publicación todavía no
salió o si hubo un error y por qué.
"""
from __future__ import annotations

import re
from datetime import datetime

LIMITE = 3900


def _limpio(texto, largo=220) -> str:
    texto = ' '.join(str(texto or '').split())
    return texto if len(texto) <= largo else texto[:largo - 1] + '…'


def _motivo_error(resultado: dict) -> str:
    bruto = ' '.join(str(resultado.get(k) or '') for k in ('detalle', 'error_final', 'salida_final'))
    m = re.search(r'ERROR(?: FINAL)?\s*[:=]\s*(.+?)(?:\s+ESTADO=|\s+RESULTADO:|$)', bruto)
    return _limpio(m.group(1) if m else (resultado.get('detalle') or 'El ejecutor no informó el motivo'))


def _texto_periodo(p) -> str:
    p = str(p or '')
    m = re.match(r'(20\d{2})-(\d{2})', p)
    if m:
        meses = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
        return f'{meses[int(m.group(2)) - 1]} {m.group(1)[2:]}'
    return p


def _periodo(resultado: dict, sheet: dict | None) -> str:
    if sheet and sheet.get('periodos'):
        return ', '.join(_texto_periodo(x) for x in sheet['periodos'] if x)
    return _texto_periodo(resultado.get('periodo'))


def _conteo(sheet: dict) -> dict:
    """Cuántos datos del período tenía que haber y cuántos quedaron en el Sheet, y por qué faltan."""
    conservadas = sheet.get('conservadas') or []
    por_formula = [c for c in conservadas if 'fórmula' in c]
    manuales = [c for c in conservadas if 'fórmula' not in c]
    escritas = int(sheet.get('escritas') or 0)
    if sheet.get('verificadas') is not None:
        en_sheet = int(sheet['verificadas'])
    else:
        en_sheet = escritas + int(sheet.get('ya_estaban') or 0) + len(por_formula)
    esperadas = int(sheet.get('esperadas') or 0)
    motivos = []
    if manuales:
        motivos.append(f"{len(manuales)} con un valor cargado a mano distinto al de la fuente; no se pisaron "
                       f"por seguridad, conviene revisarlas ({_limpio(', '.join(c.split(' ')[0] for c in manuales[:6]), 120)})")
    if sheet.get('no_encontradas'):
        motivos.append(f"{len(sheet['no_encontradas'])} sin columna en el Sheet (el título no coincide): "
                       + _limpio('; '.join(sheet['no_encontradas']), 250))
    errores_celda = [e for e in (sheet.get('errores') or [])]
    if errores_celda:
        motivos.append('error al escribir: ' + _limpio('; '.join(errores_celda), 250))
    return {'esperadas': esperadas, 'en_sheet': min(en_sheet, esperadas) if esperadas else en_sheet,
            'escritas': escritas, 'formula': len(por_formula), 'motivos': motivos}


def armar_mensaje(resumen: dict, extra_por_fuente: dict | None = None) -> str:
    """resumen: el de cualquiera de los tres ejecutores, con 'sheet_directo' si existe.

    Por indicador dice: si hubo datos nuevos, cuántos datos del período están en el
    Sheet (por ejemplo 10 de 10) y, si falta alguno, por qué.
    """
    extra_por_fuente = extra_por_fuente or {}
    por_fuente = (resumen.get('sheet_directo') or {}).get('por_fuente', {})
    lineas, avisos = [], []
    cuenta = {'nuevos': 0, 'sin': 0, 'pend': 0, 'err': 0, 'incompletos': 0}
    for r in resumen.get('resultados', []):
        nombre = _limpio(r.get('nombre') or r.get('fuente'), 80)
        fuente = r.get('fuente')
        sheet = por_fuente.get(fuente)
        estado = r.get('estado')
        escritas = int(r.get('celdas_actualizadas') or 0)
        periodo = _periodo(r, sheet)
        de_periodo = f' ({periodo})' if periodo else ''
        if estado == 'ERROR':
            cuenta['err'] += 1
            linea = f'❌ {nombre}: no se actualizó. Motivo: {_motivo_error(r)}'
        elif estado == 'PENDIENTE':
            cuenta['pend'] += 1
            linea = f'⏳ {nombre}: el organismo todavía no publicó el período siguiente.'
        elif sheet is not None and (sheet.get('esperadas') or sheet.get('verificadas') is not None):
            c = _conteo(sheet)
            total = f"{c['en_sheet']} de {c['esperadas']} datos en el Sheet" if c['esperadas'] else 'datos verificados en el Sheet'
            if c['escritas']:
                cuenta['nuevos'] += 1
                linea = f"✅ {nombre}{de_periodo}: datos nuevos cargados — {total} ({c['escritas']} escritos ahora)."
            else:
                cuenta['sin'] += 1
                linea = f'🔎 {nombre}{de_periodo}: sin datos nuevos — {total}.'
            if c['formula']:
                linea += f" {c['formula']} los calcula el Sheet con su fórmula."
            if c['motivos'] or (c['esperadas'] and c['en_sheet'] < c['esperadas']):
                cuenta['incompletos'] += 1
                faltan = c['esperadas'] - c['en_sheet']
                linea += f'\n   ⚠️ Faltan {faltan}: ' + ('; '.join(c['motivos']) if c['motivos'] else 'motivo no informado por el ejecutor') + '.'
        elif escritas > 0:
            cuenta['nuevos'] += 1
            linea = f'✅ {nombre}: datos nuevos de {periodo} ({escritas} celdas).'
        else:
            cuenta['sin'] += 1
            linea = f'🔎 {nombre}: verificado, sin datos nuevos' + (f' (último: {periodo}).' if periodo else '.')
        if extra_por_fuente.get(fuente):
            linea += ' ' + extra_por_fuente[fuente]
        lineas.append(linea)
        if sheet and sheet.get('reparadas'):
            avisos.append(f"🛠️ {nombre}: {sheet['reparadas']} celdas de meses anteriores corregidas en el Sheet "
                          "(estaban mal ubicadas o con fórmulas rotas).")
    espejo = resumen.get('espejo_sheet') or {}
    if espejo.get('estado') == 'ERROR':
        avisos.append('⚠️ No se pudo leer el Sheet antes de revisar (se usó la copia local): ' + _limpio(espejo.get('detalle'), 200))
    control = resumen.get('control_base') or {}
    hallados = control.get('hallazgos') or []

    def _detalle(lista):
        texto = ''
        for h in lista[:10]:
            contenido = h.get('formula') or h.get('valor')
            texto += (f"\n   • {h['celda']} ({_limpio(h.get('encabezado') or 'sin encabezado', 60)}"
                      + (f", {h['periodo']}" if h.get('periodo') else '') + f'): {_limpio(contenido, 60)}')
        if len(lista) > 10:
            texto += f'\n   • … y {len(lista) - 10} más (detalle en reportes/control_base).'
        return texto
    fuera = [h for h in hallados if h.get('limpiar_con') is not None]
    dentro = [h for h in hallados if h.get('limpiar_con') is None]
    if fuera:
        borrados = control.get('limpiadas', 0)
        avisos.append(f'🚨 BASE DE DATOS: {len(fuera)} cálculo(s) o valor(es) en meses futuros o fuera de la tabla'
                      + (' — ya se borraron (hay copia):' if borrados else ' — NO se borraron, revisar:') + _detalle(fuera))
    if dentro:
        avisos.append(f'👀 BASE DE DATOS: {len(dentro)} cálculo(s) fuera de lo estipulado dentro de la tabla '
                      '(miran meses que no corresponden u otra pestaña). No se tocaron, revisar:' + _detalle(dentro))
    if control.get('estado') == 'ERROR':
        avisos.append('⚠️ No se pudo controlar la hoja BASE DE DATOS: ' + _limpio(control.get('detalle'), 200))
    proteccion = control.get('proteccion') or {}
    if proteccion.get('estado') == 'ERROR':
        avisos.append('⚠️ Protección de BASE DE DATOS: ' + _limpio(proteccion.get('detalle'), 200))
    elif proteccion.get('accion') == 'quitada':
        avisos.append('🔓 Se quitó el bloqueo de BASE DE DATOS: el equipo puede volver a cargar valores y completar los cálculos.')
    formato = resumen.get('formato_sheet') or {}
    if formato.get('estado') == 'ERROR' and formato.get('detalle'):
        avisos.append('⚠️ formato_sheet: ' + _limpio(formato['detalle'], 250))
    if cuenta['err']:
        titulo = '⚠️ REVISIÓN COMPLETADA — CON ERRORES'
    elif cuenta['incompletos']:
        titulo = '⚠️ REVISIÓN COMPLETADA — HAY DATOS SIN CARGAR'
    elif cuenta['nuevos']:
        titulo = '✅ REVISIÓN COMPLETADA — DATOS NUEVOS CARGADOS'
    else:
        titulo = '🔎 REVISIÓN COMPLETADA — SIN NOVEDADES'
    try:
        fecha = datetime.fromisoformat(str(resumen.get('fin') or resumen.get('inicio')))
    except ValueError:
        fecha = datetime.now()
    resumen_linea = (f"Revisados: {len(lineas)} · nuevos: {cuenta['nuevos']} · sin novedades: {cuenta['sin']}"
                     f" · pendientes: {cuenta['pend']} · con error: {cuenta['err']}")
    partes = [titulo, '', f'Fecha: {fecha.day}/{fecha.month}/{fecha.year}, {fecha:%H:%M:%S}', resumen_linea]
    for linea in lineas + avisos:
        partes += ['', linea]
    mensaje = '\n'.join(partes)
    if len(mensaje) > LIMITE:
        mensaje = mensaje[:LIMITE - 40] + '\n\n…mensaje abreviado.'
    # Telegram se envía en modo HTML: se escapan <, > y &.
    import html
    return html.escape(mensaje, quote=False)
