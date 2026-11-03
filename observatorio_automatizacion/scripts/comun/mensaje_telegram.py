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


PRODUCTIVOS = {'comercializacion', 'pecuaria', 'estimaciones'}
# Fuentes que se informan como "dato nuevo" solo cuando entra el mes completo.
SOLO_MES_COMPLETO = {'dnrpa'}


def _vacio(valor) -> bool:
    try:
        from sheet_directo import vacio_o_cero
        return vacio_o_cero(valor)
    except Exception:  # noqa: BLE001
        return valor in (None, '', 0, '0')
# Palabras del encabezado de la hoja de productivos -> qué se actualizó, en criollo.
TEMAS_PRODUCTIVOS = [('porcino', 'porcinos'), ('bovino', 'bovinos'), ('aves', 'aves'), ('huevo', 'aves'),
                     ('proteina', 'consumo de carnes'), ('consumo', 'consumo de carnes'),
                     ('maiz', 'maíz'), ('sorgo', 'sorgo'), ('soja', 'soja'), ('trigo', 'trigo'),
                     ('sembrada', 'superficie sembrada'), ('cosechada', 'superficie cosechada'),
                     ('produccion', 'producción por cultivo'), ('participacion', 'producción por cultivo')]


def _sin_tildes(texto) -> str:
    import unicodedata
    s = unicodedata.normalize('NFD', str(texto or '').lower())
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


def _temas(sheet: dict | None) -> str:
    """De los títulos de las celdas escritas saca 'porcinos, aves', 'maíz, soja', etc."""
    vistos = []
    for cambio in (sheet or {}).get('cambios') or []:
        titulo = _sin_tildes(cambio.get('titulo'))
        for clave, tema in TEMAS_PRODUCTIVOS:
            if clave in titulo and tema not in vistos:
                vistos.append(tema)
                break
    return ', '.join(vistos)


def _revisar(resumen: dict, extra_por_fuente: dict) -> tuple[list[dict], list[str]]:
    """Una entrada por indicador (tipo: nuevo, sin, pend, err, omit) y los avisos generales."""
    por_fuente = (resumen.get('sheet_directo') or {}).get('por_fuente', {})
    try:  # aviso si el organismo publicó y no se encontró el dato (p. ej. cambió el nombre del archivo)
        import vigilancia_fuentes
        vigilancia = vigilancia_fuentes.revisar(resumen.get('resultados', []))
    except Exception:  # noqa: BLE001 - la vigilancia nunca frena el aviso
        vigilancia = {}
    items, avisos = [], []
    for r in resumen.get('resultados', []):
        nombre = _limpio(r.get('nombre') or r.get('fuente'), 80)
        fuente = r.get('fuente')
        sheet = por_fuente.get(fuente)
        estado = r.get('estado')
        escritas = int(r.get('celdas_actualizadas') or 0)
        periodo = _periodo(r, sheet)
        it = {'fuente': fuente, 'nombre': nombre, 'periodo': periodo, 'conteo': '', 'faltan': '', 'notas': [],
              'escritas': 0, 'temas': ''}
        if estado == 'ERROR':
            it.update(tipo='err', motivo=_motivo_error(r))
        elif estado == 'OMITIDA':
            it.update(tipo='omit', motivo=_limpio(r.get('detalle'), 200))
        elif estado == 'PENDIENTE':
            it.update(tipo='pend')
        elif sheet is not None and (sheet.get('esperadas') or sheet.get('verificadas') is not None):
            c = _conteo(sheet)
            it['conteo'] = f"{c['en_sheet']} de {c['esperadas']}" if c['esperadas'] else ''
            it['escritas'] = c['escritas']
            it['tipo'] = 'nuevo' if c['escritas'] else 'sin'
            if it['tipo'] == 'nuevo' and fuente in SOLO_MES_COMPLETO:
                nuevos = [x for x in (sheet.get('cambios') or []) + (r.get('cambios') or [])
                          if _vacio(x.get('antes')) and not x.get('mismo_dato')]
                if not nuevos or (c['esperadas'] and c['en_sheet'] < c['esperadas']):
                    # DNRPA corrige 1 o 2 valores de Nación varias veces por semana: eso no es "dato nuevo".
                    it['tipo'] = 'sin'
                    hecho = (f"se ajustaron {c['escritas']} valor(es) ya publicados" if not nuevos else
                             f"por ahora hay {c['en_sheet']} de {c['esperadas']} datos del mes")
                    it['notas'].append(hecho + '; se informa (con Word y plantilla) cuando entre el mes completo')
            it['temas'] = _temas(sheet) if fuente in PRODUCTIVOS else ''
            if c['motivos'] or (c['esperadas'] and c['en_sheet'] < c['esperadas']):
                faltan = c['esperadas'] - c['en_sheet']
                it['faltan'] = (f'faltan {faltan}: ' if faltan > 0 else '') + ('; '.join(c['motivos']) if c['motivos'] else 'motivo no informado')
        elif escritas > 0:
            it.update(tipo='nuevo', escritas=escritas)
        else:
            it['tipo'] = 'sin'
        if extra_por_fuente.get(fuente):
            it['notas'].append(extra_por_fuente[fuente])
        for aviso in (sheet or {}).get('avisos') or []:
            it['notas'].append('⚠️ ' + _limpio(aviso, 200))
        for nota in re.findall(r'^NOTA_TELEGRAM=(.+)$', str(r.get('salida_final') or ''), re.M):
            it['notas'].append(_limpio(nota, 300))
        if vigilancia.get(fuente) and estado != 'ERROR':
            it['vigilancia'] = vigilancia[fuente]
        items.append(it)
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
    return items, avisos


def _fecha(resumen: dict) -> str:
    try:
        f = datetime.fromisoformat(str(resumen.get('fin') or resumen.get('inicio')))
    except ValueError:
        f = datetime.now()
    return f'{f.day}/{f.month} {f:%H:%M}'


def _e(texto) -> str:
    import html
    return html.escape(str(texto or ''), quote=False)


def _corto(it: dict) -> str:
    """Una línea por indicador con novedades o con algo que mirar (texto ya escapado para HTML)."""
    p = f' ({_e(it["periodo"])})' if it['periodo'] else ''
    if it['tipo'] == 'err':
        linea = f'❌ <b>{_e(it["nombre"])}</b>: no se actualizó — {_e(it["motivo"])}'
    elif it['tipo'] == 'pend':
        linea = f'⏳ <b>{_e(it["nombre"])}</b>: todavía no se publicó el período siguiente'
    elif it['tipo'] == 'omit':
        linea = f'↪️ <b>{_e(it["nombre"])}</b>: no se revisa en este equipo — {_e(it["motivo"])}'
    elif it['tipo'] == 'nuevo':
        linea = f'✅ <b>{_e(it["nombre"])}</b>{p}: ' + (f'{it["conteo"]} datos en el Sheet ({it["escritas"]} nuevos)' if it['conteo'] else f'{it["escritas"]} datos nuevos')
    elif it['faltan']:  # sin novedades pero incompleto
        linea = f'⚠️ <b>{_e(it["nombre"])}</b>{p}: {it["conteo"]} datos en el Sheet'
    else:  # sin novedades, con un aviso de la fuente
        linea = f'🔎 <b>{_e(it["nombre"])}</b>{p}: sin novedades' + (f' — {it["conteo"]} datos en el Sheet' if it['conteo'] else '')
    if it['faltan']:
        linea += f' — {_e(it["faltan"])}'
    if it['tipo'] in ('nuevo', 'err') or it['faltan'] or it.get('mostrar'):
        for nota in it['notas']:
            linea += f'\n   {_e(nota)}'
    if it.get('vigilancia'):
        linea += f'\n   {_e(it["vigilancia"])}'
    return linea


def _mensaje_general(resumen: dict, items: list[dict], avisos: list[str]) -> str:
    """Revisión general y complejos: solo lo que cambió o hay que mirar. Si no hay nada, no se manda ('')."""
    nuevos = [i for i in items if i['tipo'] == 'nuevo']
    errores = [i for i in items if i['tipo'] == 'err']
    incompletos = [i for i in items if i['tipo'] == 'sin' and i['faltan']]
    # Sin novedades pero con algo para mirar: vigilancia de la fuente, una nota de alerta (⚠️/❌)
    # o una corrida puntual de un solo indicador (desde n8n, siempre se avisa).
    puntual = len(items) == 1
    vigilados = []
    for i in items:
        if i['tipo'] in ('sin', 'pend', 'omit') and not i['faltan'] and (
                i.get('vigilancia') or puntual or any(n.lstrip()[:1] in '⚠❌🚨' for n in i['notas'])):
            i['mostrar'] = True
            vigilados.append(i)
    if not (nuevos or errores or incompletos or vigilados or avisos):
        return ''
    if errores:
        titulo = '⚠️ <b>Revisión con errores</b>'
    elif incompletos:
        titulo = '⚠️ <b>Hay datos sin cargar</b>'
    elif nuevos:
        titulo = '✅ <b>Datos nuevos en el Sheet</b>'
    else:
        titulo = '👀 <b>Revisión: algo para mirar</b>'
    partes = [f'{titulo} · {_fecha(resumen)}']
    partes += [_corto(i) for i in nuevos + errores + incompletos + vigilados]
    sin = sum(1 for i in items if i['tipo'] == 'sin' and not i['faltan'] and i not in vigilados)
    sin_txt = f'{sin} indicador' + ('es' if sin != 1 else '')
    pend = sum(1 for i in items if i['tipo'] == 'pend' and i not in vigilados)
    otra = sum(1 for i in items if i['tipo'] == 'omit')
    resto = [f'sin novedades: {sin_txt}'] if sin else []
    resto += [f'sin publicar aún: {pend}'] if pend else []
    resto += [f'en otra PC: {otra}'] if otra else []
    if resto:
        partes.append('🔎 ' + ' · '.join(resto).capitalize())
    # Notas informativas de las demás fuentes: solo viajan si igual se manda el mensaje.
    otras = [f'📌 {_e(i["nombre"])}: {_e(n)}' for i in items if not i.get('mostrar') and i['tipo'] == 'sin' and not i['faltan']
             for n in i['notas']]
    partes += otras[:5]
    partes += [_e(a) for a in avisos]
    return '\n'.join(partes)


def _mensaje_productivos(resumen: dict, items: list[dict], avisos: list[str]) -> str:
    """Productivos (hoja aparte, la carga otra persona del equipo): siempre completo y destacado."""
    hubo = any(i['tipo'] == 'nuevo' for i in items)
    partes = [f'🌾🐄 <b>SECTOR PRODUCTIVO</b> · {_fecha(resumen)}',
              '<i>Hoja «BASE DE DATOS - Indicadores Productivos»</i>']
    lineas = []
    for it in items:
        p = f' ({_e(it["periodo"])})' if it['periodo'] else ''
        n = f'<b>{_e(it["nombre"])}</b>{p}'
        if it['tipo'] == 'err':
            linea = f'🟥 {n}: no se actualizó — {_e(it["motivo"])}'
        elif it['tipo'] == 'pend':
            linea = f'🟨 {n}: todavía no se publicó el período siguiente'
        elif it['tipo'] == 'nuevo':
            que = f' de {_e(it["temas"])}' if it['temas'] else ''
            linea = f'🟩 {n}: <b>datos nuevos{que}</b>' + (f' · {it["conteo"]} en el Sheet' if it['conteo'] else '')
        else:
            linea = f'⬜ {n}: sin novedades' + (f' · {it["conteo"]} en el Sheet' if it['conteo'] else '')
        if it['faltan']:
            linea += f'\n   ⚠️ {_e(it["faltan"])}'
        for nota in it['notas']:
            linea += f'\n   {_e(nota)}'
        lineas.append(linea)
    partes.append('<blockquote>' + '\n'.join(lineas) + '</blockquote>')
    partes.append('✅ Hubo actualización.' if hubo else '🔎 Sin cambios en esta revisión.')
    partes += [_e(a) for a in avisos]
    return '\n'.join(partes)


def armar_mensaje(resumen: dict, extra_por_fuente: dict | None = None) -> str:
    """Aviso de Telegram (HTML) de cualquiera de los tres ejecutores.

    General y complejos: corto, solo lo que se actualizó o hay que mirar ('15 de 15 datos en el Sheet').
    Si no hubo nada, devuelve '' y n8n no manda mensaje.
    Productivos: siempre completo y destacado, se haya actualizado o no.
    """
    items, avisos = _revisar(resumen, extra_por_fuente or {})
    productivos = bool(items) and all(i['fuente'] in PRODUCTIVOS for i in items)
    mensaje = (_mensaje_productivos if productivos else _mensaje_general)(resumen, items, avisos)
    if len(mensaje) > LIMITE:
        # Se corta sin romper etiquetas: se quita el formato del tramo final.
        mensaje = re.sub(r'</?(b|i|blockquote)>', '', mensaje)[:LIMITE - 40] + '\n\n…mensaje abreviado.'
    return mensaje
