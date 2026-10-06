from __future__ import annotations
'GENERAL_CONTINUO_V1: aislamiento y diagnostico por fuente.'
import argparse
import json
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
RAIZ = Path(__file__).resolve().parent
MOTOR = Path(__file__).resolve()
DESTINO = RAIZ / 'salidas_n8n' / 'resumen_general.json'
# Fuentes que solo se corren a mano (no entran en 'todas').
FUENTES_MANUALES = ['srt_historico', 'ieric_correccion', 'correccion_transferencias', 'limpieza_motos', 'control_base']
FUENTES = ['eph', 'pbi', 'iaraf', 'atp', 'canasta', 'sipa', 'ipi', 'cammesa', 'dnrpa', 'ieric', 'ipc', 'internet', 'transferencias_fiscales', 'srt', 'combustibles']

def ahora():
    return datetime.now().astimezone().isoformat(timespec='seconds')

def guardar(resumen):
    DESTINO.parent.mkdir(parents=True, exist_ok=True)
    temporal = DESTINO.with_name(f'resumen_general_{os.getpid()}.tmp')
    temporal.write_text(json.dumps(resumen, ensure_ascii=False), encoding='utf-8')
    temporal.replace(DESTINO)

def interpretar(salida):
    for linea in reversed(salida.splitlines()):
        if linea.startswith('RESUMEN_N8N='):
            return json.loads(linea.split('=', 1)[1])
    raise RuntimeError('El motor termino sin emitir RESUMEN_N8N. Consultar el registro de esta fuente.')

def finalizar_proceso(proceso):
    if os.name == 'nt':
        subprocess.run(['taskkill', '/PID', str(proceso.pid), '/T', '/F'], capture_output=True, timeout=30, check=False)
    if proceso.poll() is None:
        proceso.kill()

def ejecutar(fuente, carpeta_logs):
    registro = carpeta_logs / f'{fuente}.log'
    salida = ''
    errores = ''
    proceso = None
    try:
        entorno = os.environ.copy()
        entorno['PYTHONIOENCODING'] = 'utf-8'
        entorno['PYTHONUTF8'] = '1'
        proceso = subprocess.Popen([sys.executable, '-u', str(MOTOR), '--motor-interno', '--fuentes', fuente, '--estricto', '--forzar-google'], cwd=RAIZ, env=entorno, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='replace')
        try:
            salida, errores = proceso.communicate(timeout=2400 if fuente in FUENTES_MANUALES else 600)
        except subprocess.TimeoutExpired:
            finalizar_proceso(proceso)
            salida, errores = proceso.communicate(timeout=30)
            raise RuntimeError('La fuente supero 10 minutos, incluyendo extraccion y finalizacion. Se detuvo su proceso para continuar con la siguiente.')
        registro.write_text(salida + '\n\nSTDERR:\n' + errores, encoding='utf-8')
        diagnostico = interpretar(salida)
        resultados = diagnostico.get('resultados') or []
        resultado = next((item for item in resultados if item.get('fuente') == fuente), None)
        if resultado is None:
            raise RuntimeError(f'El resumen no contiene el resultado de la fuente solicitada: {fuente}')
        resultado = dict(resultado)
        resultado['log'] = str(registro)
        if diagnostico.get('estado') != 'OK' or resultado.get('estado') != 'OK' or proceso.returncode != 0:
            resultado['estado'] = 'ERROR'
            resultado['detalle'] = resultado.get('detalle') or resultado.get('error_final') or 'La fuente o su finalizacion fallo. Errores: ' + str(diagnostico.get('errores', []))
            return (resultado, [])
        resultado['detalle'] = 'Extraccion y preparacion terminadas. Escritura en Google Sheets pendiente.'
        return (resultado, diagnostico.get('actualizaciones_google', []))
    except Exception as error:
        try:
            registro.write_text(salida + '\n\nSTDERR:\n' + errores + '\n\nERROR:\n' + str(error), encoding='utf-8')
        except OSError:
            pass
        return ({'fuente': fuente, 'nombre': fuente.upper(), 'estado': 'ERROR', 'detalle': str(error), 'error_final': errores[-1500:] or str(error), 'log': str(registro), 'celdas_actualizadas': 0, 'hubo_cambios': False}, [])

def nota_inflacion_atp(salida):
    """Informa con qué inflación estimada de IARAF se calcularon los valores reales de ATP."""
    tasa = re.search(r'"tasa_estimada":\s*([0-9.]+)', salida)
    anterior = re.search(r'"tasa_anterior":\s*([0-9.]+)', salida)
    if tasa:
        t = float(tasa.group(1))
        texto = f'Real calculado con inflación estimada IARAF {t:g}%'.replace('.', ',')
        if anterior and float(anterior.group(1)) != t:
            texto += f' (antes {float(anterior.group(1)):g}%)'.replace('.', ',')
        return texto
    pendiente = re.search(r'"estado":\s*"PENDIENTE",\s*"detalle":\s*"([^"]{0,160})', salida)
    if pendiente:
        return '⏳ Valores reales pendientes: ' + pendiente.group(1)
    return 'No se pudo confirmar la inflación IARAF usada para los valores reales'


def escribir_en_sheet(resumen):
    """Escribe directo en Google Sheets y deja en cada resultado lo que realmente pasó.

    n8n ya no escribe: recibe actualizaciones_google vacío y solo envía el aviso.
    """
    import re as _re
    sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
    from sheet_directo import escribir_actualizaciones, describir
    cargas = resumen.get('actualizaciones_google') or []
    try:
        escritura = escribir_actualizaciones(cargas, RAIZ / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx')
    except Exception as error:
        escritura = {'estado': 'ERROR', 'por_fuente': {}, 'errores': ['sheet'], 'detalle': str(error)}
        resumen['errores'].append('google_sheet')
        resumen['resultados'].append({'fuente': 'google_sheet', 'nombre': 'Escritura en Google Sheets', 'estado': 'ERROR',
                                      'detalle': f'No se pudo escribir en el Sheet: {error}', 'celdas_actualizadas': 0})
    resumen['sheet_directo'] = escritura
    total = 0
    modificadas = []
    for resultado in resumen['resultados']:
        fuente = resultado.get('fuente')
        if resultado.get('estado') != 'OK' or fuente == 'google_sheet':
            continue
        datos = escritura.get('por_fuente', {}).get(fuente)
        if fuente == 'dnrpa' and datos is None:
            texto = resultado.get('salida_final') or ''
            m = _re.search(r'"celdas_escritas":\s*(\d+)', texto)
            escritas = int(m.group(1)) if m else 0
            datos = {'periodos': [resultado.get('periodo') or ''], 'escritas': escritas, 'no_encontradas': [], 'conservadas': [], 'errores': []}
            m_e = _re.search(r'"datos_esperados":\s*(\d+)', texto)
            m_v = _re.search(r'"datos_verificados":\s*(\d+)', texto)
            if m_e and m_v:
                datos['esperadas'] = int(m_e.group(1))
                datos['verificadas'] = int(m_v.group(1))
            escritura.setdefault('por_fuente', {})['dnrpa'] = datos
        if datos is None:
            resultado['detalle'] = 'Sin datos nuevos para enviar al Sheet.'
            resultado['celdas_actualizadas'] = 0
            continue
        resultado['detalle'] = describir(datos)
        resultado['celdas_actualizadas'] = datos.get('escritas', 0)
        resultado['hubo_cambios'] = bool(datos.get('escritas'))
        if datos.get('errores'):
            resultado['estado'] = 'ERROR'
            if fuente not in resumen['errores']:
                resumen['errores'].append(fuente)
        if datos.get('escritas'):
            total += datos['escritas']
            modificadas.append(fuente)
    resumen['notificacion'] = {'hubo_cambios': bool(modificadas), 'total_celdas_actualizadas': total,
                               'fuentes_modificadas': modificadas, 'cambios_locales_no_confirmados_en_sheet': False}
    resumen['actualizaciones_google'] = []


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--fuentes', default='todas')
    parser.add_argument('--estricto', action='store_true')
    parser.add_argument('--forzar-google', action='store_true')
    parser.add_argument('--listar', action='store_true')
    parser.add_argument('--solo-finalizar', action='store_true')
    args = parser.parse_args()
    if args.listar:
        print(json.dumps(FUENTES, ensure_ascii=False, indent=2))
        return 0
    if args.solo_finalizar:
        return subprocess.call([sys.executable, '-u', str(MOTOR), '--motor-interno', '--solo-finalizar'], cwd=RAIZ)
    seleccion = list(FUENTES) if args.fuentes.strip().lower() == 'todas' else list(dict.fromkeys((x.strip().lower() for x in args.fuentes.split(',') if x.strip())))
    resumen = {'version_ejecutor': 'GENERAL_CONTINUO_V1', 'inicio': ahora(), 'fin': None, 'estado': 'EN_CURSO', 'solicitadas': seleccion, 'ejecutadas': [], 'errores': [], 'omitidas': [], 'resultados': [], 'actualizaciones_google': [], 'notificacion': {'hubo_cambios': False, 'total_celdas_actualizadas': 0, 'fuentes_modificadas': [], 'cambios_locales_no_confirmados_en_sheet': True}}
    logs = RAIZ / 'logs' / 'general_continuo' / datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    logs.mkdir(parents=True, exist_ok=True)
    # Control de BASE DE DATOS: cálculos o valores fuera de lugar (meses futuros,
    # fuera de la tabla) se guardan, se borran y se avisan; la hoja queda protegida.
    if any(f not in FUENTES_MANUALES for f in seleccion):
        try:
            sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
            from control_base import controlar
            resumen['control_base'] = controlar()
        except Exception as error:
            resumen['control_base'] = {'estado': 'ERROR', 'detalle': str(error)[:300]}
        print('CONTROL_BASE=' + json.dumps(resumen['control_base'], ensure_ascii=False, default=str)[:3000], flush=True)
    # El Sheet de la plataforma manda: antes de revisar, el Excel local se pone al día con él.
    try:
        sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
        from sheet_directo import espejar_desde_sheet
        resumen['espejo_sheet'] = espejar_desde_sheet(RAIZ / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx')
    except Exception as error:
        resumen['espejo_sheet'] = {'estado': 'ERROR', 'detalle': str(error)}
    print('ESPEJO_SHEET=' + json.dumps(resumen['espejo_sheet'], ensure_ascii=False, default=str)[:2000], flush=True)
    guardar(resumen)
    for fuente in seleccion:
        print(f'INICIANDO={fuente}', flush=True)
        if fuente not in FUENTES and fuente not in FUENTES_MANUALES:
            resultado = {'fuente': fuente, 'estado': 'ERROR', 'detalle': 'Fuente desconocida en el registro.', 'celdas_actualizadas': 0}
            cargas = []
        else:
            resultado, cargas = ejecutar(fuente, logs)
        texto_error = ' '.join(str(resultado.get(k) or '') for k in ('detalle', 'salida_final', 'error_final'))
        if resultado.get('estado') == 'ERROR' and re.search(r'todav[ií]a no (?:est[aá] disponible|fue publicad|se public)', texto_error, re.IGNORECASE):
            # Publicación oficial aún no disponible: no es una falla.
            resultado['estado'] = 'PENDIENTE'
            m = re.search(r'[ÚU]ltimo per[ií]odo oficial detectado: ([0-9-]+)', texto_error)
            resultado['detalle'] = 'Todavía no se publicó el período siguiente' + (f'; el último publicado ({m.group(1)}) ya está cargado.' if m else '.')
        resumen['resultados'].append(resultado)
        resumen['ejecutadas'].append(fuente)
        if resultado.get('estado') == 'PENDIENTE':
            pass
        elif resultado.get('estado') == 'OK':
            for carga in cargas:
                carga['fuente'] = fuente
            resumen['actualizaciones_google'].extend(cargas)
        else:
            resumen['errores'].append(fuente)
        resumen['fin'] = ahora()
        guardar(resumen)
        print(f"TERMINADO={fuente}; ESTADO={resultado.get('estado')}; DETALLE={resultado.get('detalle', '')}", flush=True)
    escribir_en_sheet(resumen)
    resumen['estado'] = 'ERROR' if resumen['errores'] else 'OK'
    resumen['fin'] = ahora()
    try:
        from mensaje_telegram import armar_mensaje
        extra = {r['fuente']: nota_inflacion_atp(r.get('salida_final') or '') for r in resumen['resultados'] if r.get('fuente') == 'atp' and r.get('estado') == 'OK'}
        resumen['notificacion']['mensaje_telegram'] = armar_mensaje(resumen, extra)
    except Exception as error:
        resumen['notificacion']['mensaje_telegram'] = f'⚠️ No se pudo armar el aviso: {error}'
    guardar(resumen)
    print('RESUMEN_N8N=' + json.dumps(resumen, ensure_ascii=False), flush=True)
    print('RESUMEN_ARCHIVO=' + str(DESTINO), flush=True)
    return 0

def ejecutar_motor_integrado():
    """Punto único de entrada para n8n.

Ejecuta los procesos de cada fuente desde un único comando y devuelve
un resumen JSON estable para n8n.
"""
    import argparse
    import json
    import os
    import re
    import subprocess
    import sys
    import time
    from dataclasses import dataclass
    from datetime import datetime
    from pathlib import Path
    from openpyxl import load_workbook
    RAIZ = Path(__file__).resolve().parent
    ARCHIVO_EXCEL = RAIZ / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx'
    HOJA_DESTINO = 'BASE DE DATOS'
    FUENTES_DIRECTAS_GOOGLE = {'eph', 'transferencias_fiscales'}
    sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
    from control_resultados import ErrorControlFuente, validar_emision_basica
    CAMPOS_GOOGLE_POR_FUENTE: dict[str, tuple[str, ...]] = {'ipc': ('Indice de precios al consumidor- Nivel general (Región NEA)', 'Indice de precios al consumidor- Nivel general N', 'IPC nivel general con respecto al mes anterior (Region NEA) ', 'IPC nivel general con respecto al mes anterior N', 'IPC nivel general con respecto a diciembre del año anterior (IPC Acumulado) Region NEA ', 'IPC nivel general con respecto a diciembre del año anterior (IPC acumulado) N', 'IPC nivel general con respecto al mismo mes del año anterior (Region NEA)', 'IPC nivel general con respecto al mismo mes del año anterior N', 'Var % mensual IPC nivel general  segun divisiones (Region NEA): Alimentos y bebidas no alcohólicas', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Bebidas alcohólicas y tabaco', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Prendas de vestir y calzado', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Vivienda, agua, electricidad, gas y otros combustibles', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Equipamiento y mantenimiento del hogar', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Salud', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Transporte', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Comunicación', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Recreación y cultura', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Educación', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Restaurantes y hoteles', 'Var. % mensual IPC nivel general segun divisiones (Region NEA): Bienes y servicios varios', 'Var % interanual IPC nivel general según divisiones (Region NEA): Alimentos y bebidas no alcohólicas', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Bebidas alcohólicas y tabaco', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Prendas de vestir y calzado', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Vivienda, agua, electricidad, gas y otros combustibles', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Equipamiento y mantenimiento del hogar', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Salud', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Transporte', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Comunicación', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Recreación y cultura', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Educación', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Restaurantes y hoteles', 'Var. % interanual IPC nivel general segun divisiones (Region NEA): Bienes y servicios varios'), 'canasta': ('CBA. Valores mensuales por adulto equivalente. Línea de indigencia. (Valor en $)\n', 'VAR% mensual CBA adulto equivalente', 'VAR% anual CBA adulto equivalente', 'Acumulado CBA adulto equivalente año en curso', 'CBT. Valores mensuales por adulto equivalente. Línea de pobreza (Valor en $)', 'VAR% mensual CBT adulto equivalente', 'VAR% anual CBT adulto equivalente', 'Acumulado CBT adulto equivalente año en curso', 'Cálculo de la CBA para un hogar de ejemplo. Hogar 1', 'Cálculo de la CBA para un hogar de ejemplo. Hogar 2', 'VAR% mensual CBA hogar 2', 'VAR% anual CBA hogar 2', 'Cálculo de la CBA para un hogar de ejemplo. Hogar 3', 'VAR% mensual CBA hogar 3', 'VAR% anual CBA hogar 3', 'Cálculo de la CBT para un hogar de ejemplo. Hogar 1', 'Cálculo de la CBT para un hogar de ejemplo. Hogar 2', 'VAR% mensual CBT hogar 2', 'VAR% anual CBT hogar 2', 'Cálculo de la CBT para un hogar de ejemplo. Hogar 3', 'VAR% mensual CBT hogar 3', 'VAR% anual CBT hogar 3'), 'atp': ('Recaudación tributaria F', 'Recaudación tributaria acumulada F', 'Variación intermensual recaudación tributaria F Nominal', 'Variación interanual recaudación tributaria F Nominal'), 'dnrpa': ('Patentamiento de automóviles F', 'Transferencias de automoviles F', 'Patentamiento de automóviles N', 'Transferencias de automoviles N', 'Patentamiento de Motos F', 'Transferencias de Motos F', 'Patentamiento de Motos N', 'Transferencias de Motos N'), 'iaraf': ('Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ F', 'Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Nominal F', 'Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Var% Anual Real F', 'Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Acumulado anual. En millones de $ F', 'Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ N', 'Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Nominal N', 'Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Real N', 'Transferencias automáticas por coparticipación, leyes compl. y compensac. Acumulado anual. En millones de $ N'), 'internet': ('Cantidad de accesos a internet F', 'Cantidad de accesos a internet - Var. Interanual (%) F', 'Cantidad de accesos a internet N', 'Cantidad de accesos a internet - Var. Interanual (%) N ', 'Cantidad de accesos fijos a internet F', 'Cantidad de accesos fijos a internet - Var. Interanual (%) F', 'Cantidad de accesos fijos a internet N', 'Cantidad de accesos fijos a internet - Var. Interanual (%) N', 'Cantidad de accesos moviles a internet F', 'Cantidad de accesos moviles a internet - Var. Interanual (%) F', 'Cantidad de accesos moviles a internet N', 'Cantidad de accesos moviles a internet - Var. Interanual (%) N')}
    PATRONES_GOOGLE_POR_FUENTE = {'pbi': ('^PBI ',), 'sipa': ('Personas con trabajo registrado', 'asalariados registrados del sector privado', 'Remuneración .* trabajadores registrados del sector privado'), 'ipc': ('^IPC ', '^Indice de precios al consumidor', '^Var.*IPC nivel general')}

    def capturar_hoja() -> dict:
        """Captura valores y fórmulas para detectar todo cambio real."""
        libro = load_workbook(ARCHIVO_EXCEL, read_only=True, data_only=False)
        try:
            hoja = libro[HOJA_DESTINO]
            filas = list(hoja.iter_rows(values_only=True))
            if not filas:
                return {'encabezados': {}, 'valores': {}}
            return {'encabezados': {c: valor for c, valor in enumerate(filas[0], start=1)}, 'valores': {(f, c): valor for f, fila in enumerate(filas[1:], start=2) for c, valor in enumerate(fila, start=1)}}
        finally:
            libro.close()

    def normalizar_periodo(valor) -> str:
        meses = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
        if isinstance(valor, datetime):
            return f'{meses[valor.month - 1]} {valor.year % 100:02d}'
        texto = str(valor).strip()
        coincidencia = re.fullmatch('(20\\d{2})-(0[1-9]|1[0-2])(?:-\\d{2})?', texto)
        if coincidencia:
            anio = int(coincidencia.group(1))
            mes = int(coincidencia.group(2))
            return f'{meses[mes - 1]} {anio % 100:02d}'
        return texto

    def preparar_actualizaciones_google(antes: dict, despues: dict) -> list[dict]:
        encabezados = despues['encabezados']
        cp = next((c for c, t in encabezados.items() if str(t or '').strip().casefold() == 'periodo'), None)
        if cp is None:
            raise RuntimeError(f'No se encontró el encabezado PERIODO en {HOJA_DESTINO}')
        filas = {}
        for f, c in set(antes['valores']) | set(despues['valores']):
            previo = antes['valores'].get((f, c))
            nuevo = despues['valores'].get((f, c))
            if previo == nuevo or c == cp:
                continue
            titulo = encabezados.get(c)
            periodo = despues['valores'].get((f, cp))
            if titulo in (None, '') or periodo in (None, ''):
                continue
            item = filas.setdefault(f, {'hoja': HOJA_DESTINO, 'row_number': f, 'periodo': normalizar_periodo(periodo), 'valores': {}})
            item['valores'][str(titulo)] = nuevo
        return list(filas.values())

    def preparar_actualizaciones_forzadas(despues: dict, seleccion: list[str], resultados: list[dict], forzar: bool=False) -> list[dict]:
        """Exporta campos seleccionados aunque el respaldo local ya coincida."""
        encabezados = despues['encabezados']
        columna_periodo = next((columna for columna, titulo in encabezados.items() if str(titulo or '').strip().casefold() == 'periodo'), None)
        if columna_periodo is None:
            raise RuntimeError(f'No se encontró el encabezado PERIODO en {HOJA_DESTINO}')
        columnas_por_titulo = {str(titulo): columna for columna, titulo in encabezados.items() if titulo not in (None, '')}
        periodos_por_fuente = {resultado['fuente']: resultado.get('periodos') or ([resultado.get('periodo')] if resultado.get('periodo') else []) for resultado in resultados if resultado.get('periodo') or resultado.get('periodos')}
        forzadas = []
        for resultado in resultados:
            if resultado.get('fuente') not in seleccion or resultado.get('estado') != 'OK':
                continue
            for carga in resultado.get('cargas_google', []):
                periodo = normalizar_periodo(carga.get('periodo'))
                filas = [fila for (fila, columna), valor in despues['valores'].items() if columna == columna_periodo and normalizar_periodo(valor) == periodo]
                valores = carga.get('valores', {})
                if len(filas) != 1 or filas[0] != carga.get('row_number') or (not valores):
                    raise RuntimeError(f'Carga con fila o período inválido: {periodo}')
                faltantes = [titulo for titulo in valores if titulo not in columnas_por_titulo]
                if faltantes:
                    raise RuntimeError('Google Sheets no contiene estos encabezados: ' + ', '.join(faltantes))
                forzadas.append({'hoja': HOJA_DESTINO, 'row_number': filas[0], 'periodo': periodo, 'valores': valores, 'permitir_reemplazos': bool(carga.get('permitir_reemplazos'))})
        for resultado in resultados:
            if resultado.get('fuente') not in seleccion or resultado.get('estado') != 'OK':
                continue
            if resultado.get('fuente') in PATRONES_GOOGLE_POR_FUENTE:
                continue
            if resultado.get('cargas_google'):
                continue
            detalle = resultado.get('detalle_completo')
            periodo = resultado.get('periodo')
            if not isinstance(detalle, list) or not detalle or (not periodo):
                continue
            valores = {str(item['titulo']): item.get('despues') for item in detalle if isinstance(item, dict) and item.get('titulo')}
            faltantes = [titulo for titulo in valores if titulo not in columnas_por_titulo]
            if faltantes:
                raise RuntimeError('Google Sheets no contiene estos encabezados: ' + ', '.join(faltantes))
            periodo_normalizado = normalizar_periodo(periodo)
            filas = [fila for (fila, columna), valor in despues['valores'].items() if columna == columna_periodo and normalizar_periodo(valor) == periodo_normalizado]
            if len(filas) != 1:
                raise RuntimeError(f"No se pudo localizar una única fila para {resultado.get('fuente')} {periodo}.")
            forzadas.append({'hoja': HOJA_DESTINO, 'row_number': filas[0], 'periodo': periodo_normalizado, 'valores': valores})
        fuentes_con_carga_completa = {resultado.get('fuente') for resultado in resultados if resultado.get('estado') == 'OK' and isinstance(resultado.get('detalle_completo'), list) and resultado.get('detalle_completo') and (resultado.get('fuente') not in PATRONES_GOOGLE_POR_FUENTE)}
        exitosas = {resultado.get('fuente') for resultado in resultados if resultado.get('estado') == 'OK'}
        for fuente in seleccion:
            if fuente not in exitosas:
                continue
            if fuente in fuentes_con_carga_completa:
                continue
            campos = CAMPOS_GOOGLE_POR_FUENTE.get(fuente)
            if fuente == 'sipa':
                inicio = next((c for c, t in encabezados.items() if str(t or '').strip() == 'Cant. Personas con trabajo registrado. Total país. Serie Original (en miles)'), None)
                fin = next((c for c, t in encabezados.items() if str(t or '').strip() == 'Var. absoluta i.a. asalariados registrados del sector privado S/E F'), None)
                if inicio is None or fin is None or fin - inicio + 1 != 70:
                    raise RuntimeError('No se pudo delimitar el bloque SIPA de 70 columnas.')
                campos = tuple((str(encabezados[c]) for c in range(inicio, fin + 1) if encabezados.get(c) not in (None, '')))
            elif fuente == 'atp':
                # Bloque ATP completo: datos de ATP + columnas calculadas (fórmulas del Excel local).
                inicio = next((c for c, t in encabezados.items() if str(t or '').strip() == 'Recaudación tributaria F'), None)
                fin = next((c for c, t in encabezados.items() if str(t or '').strip().startswith('Total recaudacion "otros impuestos')), None)
                if inicio is None or fin is None or fin < inicio:
                    raise RuntimeError('No se pudo delimitar el bloque ATP.')
                campos = tuple((str(encabezados[c]) for c in range(inicio, fin + 1) if encabezados.get(c) not in (None, '')))
            elif fuente in PATRONES_GOOGLE_POR_FUENTE:
                patrones = tuple((re.compile(p, re.IGNORECASE) for p in PATRONES_GOOGLE_POR_FUENTE[fuente]))
                campos = tuple((str(titulo) for titulo in encabezados.values() if titulo not in (None, '') and any((p.search(str(titulo)) for p in patrones))))
            periodos = list(periodos_por_fuente.get(fuente, []))
            if fuente == 'atp' and periodos:
                # También el mes anterior: ATP revisa cifras y así se corrigen sus fórmulas.
                coincidencia = re.match('(20\\d{2})-(\\d{2})', str(periodos[-1]))
                if coincidencia:
                    anio, mes = int(coincidencia.group(1)), int(coincidencia.group(2))
                    anterior = f'{anio - 1}-12' if mes == 1 else f'{anio}-{mes - 1:02d}'
                    periodos = [anterior] + periodos
            if not campos or not periodos:
                continue
            for periodo in periodos:
                periodo_normalizado = normalizar_periodo(periodo)
                filas = [fila for (fila, columna), valor in despues['valores'].items() if columna == columna_periodo and normalizar_periodo(valor) == periodo_normalizado]
                if len(filas) != 1:
                    raise RuntimeError(f'No se pudo localizar una única fila para {fuente} {periodo}.')
                fila = filas[0]
                valores = {}
                for titulo in campos:
                    columna = columnas_por_titulo.get(titulo)
                    if columna is None:
                        raise RuntimeError(f'No se encontró el encabezado {titulo!r}.')
                    valores[titulo] = despues['valores'].get((fila, columna))
                forzadas.append({'hoja': HOJA_DESTINO, 'row_number': fila, 'periodo': normalizar_periodo(despues['valores'].get((fila, columna_periodo))), 'valores': valores})
        return forzadas

    @dataclass(frozen=True)
    class Fuente:
        """Configuración de una fuente de datos."""
        nombre: str
        ruta_relativa: str
        argumentos: tuple[str, ...] = ()

        def obtener_ruta(self) -> Path:
            """Devuelve la ruta absoluta del ejecutor de la fuente."""
            return RAIZ / self.ruta_relativa

        def validar_ruta(self) -> str | None:
            """Valida que el ejecutor exista y no esté vacío."""
            ruta = self.obtener_ruta()
            if not ruta.is_file():
                return f'No existe el extractor: {self.ruta_relativa}'
            if ruta.stat().st_size == 0:
                return f'El extractor está vacío: {self.ruta_relativa}'
            return None
    FUENTES: dict[str, Fuente] = {'eph': Fuente(nombre='Mercado de trabajo - EPH', ruta_relativa='scripts/mercado_trabajo/ejecutar_eph.py', argumentos=('--solo-extraer',)), 'pbi': Fuente(nombre='PBI', ruta_relativa='scripts/pbi/ejecutar_pbi.py'), 'iaraf': Fuente(nombre='IARAF', ruta_relativa='scripts/iaraf/ejecutar_iaraf.py'), 'atp': Fuente(nombre='ATP - Recaudación Tributaria Provincial', ruta_relativa='scripts/atp/ejecutar_atp.py'), 'canasta': Fuente(nombre='CBA y CBT', ruta_relativa='scripts/canasta/ejecutar_canasta.py'), 'sipa': Fuente(nombre='SIPA - Trabajo registrado', ruta_relativa='scripts/sipa/ejecutar_sipa.py'), 'ipi': Fuente(nombre='IPI manufacturero', ruta_relativa='scripts/ipi/ejecutar_ipi.py'), 'cammesa': Fuente(nombre='CAMMESA - Demanda eléctrica de Formosa', ruta_relativa='scripts/cammesa/ejecutar_cammesa.py'), 'dnrpa': Fuente(nombre='DNRPA', ruta_relativa='scripts/dnrpa/ejecutar_dnrpa.py'), 'ieric': Fuente(nombre='IERIC', ruta_relativa='scripts/ieric/ejecutar_ieric.py'), 'ipc': Fuente(nombre='IPC', ruta_relativa='scripts/ipc/ejecutar_ipc.py'), 'internet': Fuente(nombre='INDEC - Accesos a internet', ruta_relativa='scripts/internet/ejecutar_internet.py'), 'transferencias_fiscales': Fuente(nombre='Transferencias automaticas - IARAF', ruta_relativa='scripts/transferencias_fiscales/ejecutar_transferencias_fiscales.py', argumentos=('--solo-extraer',)), 'srt': Fuente(nombre='SRT - Trabajadores y empleadores (Formosa)', ruta_relativa='scripts/srt/ejecutar_srt.py'), 'srt_historico': Fuente(nombre='SRT - Carga histórica desde ene 2023', ruta_relativa='scripts/srt/ejecutar_srt.py', argumentos=('--desde', '2023-01')), 'combustibles': Fuente(nombre='Precio de combustibles - CECHA', ruta_relativa='scripts/combustibles/ejecutar_combustibles.py'), 'ieric_correccion': Fuente(nombre='IERIC - Corrección verificada con informes oficiales', ruta_relativa='scripts/ieric/ejecutar_ieric.py', argumentos=('--aplicar-correccion', 'correccion_20261002.json')), 'correccion_transferencias': Fuente(nombre='Transferencias IARAF - Corrección del acumulado Nación 2026', ruta_relativa='scripts/correcciones/aplicar_correccion_sheet.py', argumentos=('--archivo', 'transferencias_acumulado_2026.json')), 'limpieza_motos': Fuente(nombre='DNRPA - Limpieza de celdas cargadas por error (motos, nov/dic 26)', ruta_relativa='scripts/correcciones/aplicar_correccion_sheet.py', argumentos=('--archivo', 'limpieza_motos_filas_futuras.json')), 'control_base': Fuente(nombre='Control de la hoja BASE DE DATOS (solo revisar)', ruta_relativa='scripts/comun/control_base.py', argumentos=('--solo-revisar',))}

    def configurar_salida() -> None:
        """Configura la salida del ejecutor general en UTF-8."""
        for flujo in (sys.stdout, sys.stderr):
            try:
                flujo.reconfigure(encoding='utf-8', errors='replace')
            except (AttributeError, OSError):
                pass

    def extraer_resultado_json(salida: str) -> dict | None:
        """Recupera el objeto JSON principal impreso por un ejecutor."""
        decodificador = json.JSONDecoder()
        candidatos: list[dict] = []
        for posicion, caracter in enumerate(salida):
            if caracter != '{':
                continue
            try:
                objeto, _ = decodificador.raw_decode(salida[posicion:])
            except json.JSONDecodeError:
                continue
            if isinstance(objeto, dict):
                candidatos.append(objeto)
        if not candidatos:
            return None
        return next((objeto for objeto in candidatos if 'fuente' in objeto and ('actualizacion' in objeto or 'periodo' in objeto or 'series' in objeto)), candidatos[0])

    def periodo_resultado(datos: dict | None) -> str | None:
        """Obtiene un período uniforme, incluido el formato propio de DNRPA."""
        if not datos:
            return None
        periodo = datos.get('periodo')
        if periodo:
            return str(periodo)
        anio = datos.get('anio')
        mes = datos.get('mes_numero')
        try:
            if anio is not None and mes is not None:
                return f'{int(anio):04d}-{int(mes):02d}'
        except (TypeError, ValueError):
            pass
        return None

    def resumir_cambios(datos: dict | None) -> dict:
        """Normaliza los cambios informados por un ejecutor compatible."""
        if not datos:
            return {'cambios_informados': False, 'hubo_cambios': False, 'celdas_actualizadas': 0, 'periodo': None, 'cambios': [], 'valores_extraidos': {}, 'limpiar_periodo_erroneo': None}
        actualizacion = datos.get('actualizacion')
        if not isinstance(actualizacion, dict):
            return {'cambios_informados': False, 'hubo_cambios': False, 'celdas_actualizadas': 0, 'periodo': periodo_resultado(datos), 'cambios': [], 'valores_extraidos': datos.get('valores', {}), 'limpiar_periodo_erroneo': datos.get('limpiar_periodo_erroneo')}
        cantidad = actualizacion.get('celdas_actualizadas', 0)
        try:
            cantidad = int(cantidad)
        except (TypeError, ValueError):
            cantidad = 0
        cambios = actualizacion.get('cambios', [])
        if not isinstance(cambios, list):
            cambios = []
        return {'cambios_informados': True, 'hubo_cambios': cantidad > 0, 'celdas_actualizadas': cantidad, 'periodo': periodo_resultado(datos), 'periodos': datos.get('periodos', []), 'cambios': cambios, 'detalle_completo': actualizacion.get('detalle_completo', []), 'cargas_google': actualizacion.get('cargas_google', []), 'reporte': datos.get('reporte'), 'power_bi': datos.get('power_bi'), 'valores_extraidos': datos.get('valores', {}), 'limpiar_periodo_erroneo': datos.get('limpiar_periodo_erroneo')}

    def ejecutar_fuente(clave: str, fuente: Fuente, estricto: bool) -> dict:
        """Ejecuta una fuente y devuelve un resultado normalizado."""
        error_ruta = fuente.validar_ruta()
        if error_ruta:
            return {'fuente': clave, 'nombre': fuente.nombre, 'estado': 'ERROR' if estricto else 'OMITIDA', 'detalle': error_ruta}
        ruta = fuente.obtener_ruta()
        inicio = time.monotonic()
        entorno = os.environ.copy()
        entorno['PYTHONIOENCODING'] = 'utf-8'
        entorno['PYTHONUTF8'] = '1'
        try:
            proceso = subprocess.run([sys.executable, '-u', str(ruta), *fuente.argumentos], cwd=RAIZ, env=entorno, text=True, capture_output=True, encoding='utf-8', errors='replace', check=False, timeout=1800 if clave in FUENTES_MANUALES else 300)
        except Exception as error:
            return {'fuente': clave, 'nombre': fuente.nombre, 'estado': 'ERROR', 'duracion_segundos': round(time.monotonic() - inicio, 2), 'detalle': f'No se pudo iniciar el extractor: {error}'}
        salida = proceso.stdout.strip()
        error = proceso.stderr.strip()
        datos = extraer_resultado_json(salida)
        cambios = resumir_cambios(datos)
        error_control = ''
        if proceso.returncode == 0:
            try:
                validar_emision_basica(datos, exigir_fuente=False)
            except ErrorControlFuente as error_calidad:
                error_control = str(error_calidad)
        return {'fuente': clave, 'nombre': fuente.nombre, 'estado': 'OK' if proceso.returncode == 0 and (not error_control) else 'ERROR', 'codigo': proceso.returncode, 'duracion_segundos': round(time.monotonic() - inicio, 2), 'salida_final': salida[-4000:], 'error_final': (error[-4000:] + ('\nCONTROL_CALIDAD=' + error_control if error_control else '')).strip(), 'detalle': error_control, **cambios}

    def seleccionar_fuentes(texto: str) -> list[str]:
        """Interpreta 'todas' o una lista de fuentes separadas por comas."""
        texto = texto.strip().casefold()
        if texto in {'todas', 'todo', 'all'}:
            return [clave for clave in FUENTES if clave not in FUENTES_MANUALES]
        seleccion: list[str] = []
        for item in texto.split(','):
            clave = item.strip().casefold()
            if clave and clave not in seleccion:
                seleccion.append(clave)
        if not seleccion:
            raise ValueError('No se indicó ninguna fuente para ejecutar.')
        desconocidas = [clave for clave in seleccion if clave not in FUENTES]
        if desconocidas:
            disponibles = ', '.join(FUENTES)
            raise ValueError('Fuentes desconocidas: ' + ', '.join(desconocidas) + f'. Fuentes disponibles: {disponibles}')
        return seleccion

    def finalizar_excel(fuentes: list[str] | None=None) -> dict:
        """Completa fórmulas y recalcula el libro una sola vez al final."""
        ruta = RAIZ / 'scripts' / 'comun' / 'finalizar_excel.py'
        inicio = time.monotonic()
        comando = [sys.executable, '-u', str(ruta)]
        if fuentes:
            comando.extend(['--fuentes', ','.join(fuentes)])
        proceso = subprocess.run(comando, cwd=RAIZ, text=True, capture_output=True, encoding='utf-8', errors='replace', check=False)
        salida = proceso.stdout.strip()
        detalle = None
        for linea in salida.splitlines():
            if linea.startswith('FINALIZACION_EXCEL='):
                try:
                    detalle = json.loads(linea.split('=', 1)[1])
                except json.JSONDecodeError:
                    detalle = None
        return {'estado': 'OK' if proceso.returncode == 0 else 'ERROR', 'codigo': proceso.returncode, 'duracion_segundos': round(time.monotonic() - inicio, 2), 'detalle': detalle, 'salida_final': salida[-4000:], 'error_final': proceso.stderr.strip()[-4000:]}

    def mostrar_fuentes() -> None:
        """Muestra las fuentes configuradas y sus rutas."""
        listado = {}
        for clave, fuente in FUENTES.items():
            ruta = fuente.obtener_ruta()
            error_ruta = fuente.validar_ruta()
            listado[clave] = {'nombre': fuente.nombre, 'ruta': fuente.ruta_relativa, 'existe': ruta.is_file(), 'valida': error_ruta is None, 'detalle': error_ruta}
        print(json.dumps(listado, ensure_ascii=False, indent=2))
        print('ESTADO=' + ('OK' if all((item['valida'] for item in listado.values())) else 'ERROR'))

    def main() -> int:
        """Ejecuta las fuentes solicitadas y genera el resumen para n8n."""
        configurar_salida()
        parser = argparse.ArgumentParser(description='Ejecutor general de indicadores para n8n.')
        parser.add_argument('--fuentes', default='todas', help="Usar 'todas' o una lista separada por comas. Ejemplo: iaraf,ipi")
        parser.add_argument('--listar', action='store_true', help='Muestra las fuentes configuradas sin ejecutarlas.')
        parser.add_argument('--solo-finalizar', action='store_true', help='Completa fórmulas y recalcula Excel sin ejecutar fuentes.')
        parser.add_argument('--forzar-google', action='store_true', help='Opción conservada por compatibilidad; las fuentes OK se preparan siempre para Sheets.')
        parser.add_argument('--estricto', action='store_true', help='Devuelve error si falta o está vacío el ejecutor de una fuente.')
        args = parser.parse_args()
        if args.listar:
            mostrar_fuentes()
            return 0
        if args.solo_finalizar:
            resultado = finalizar_excel()
            print('FINALIZACION_GENERAL=' + json.dumps(resultado, ensure_ascii=False))
            print(f"ESTADO={resultado['estado']}")
            return 0 if resultado['estado'] == 'OK' else 1
        try:
            seleccion = seleccionar_fuentes(args.fuentes)
        except ValueError as error:
            print(f'ERROR={error}')
            print('ESTADO=ERROR')
            return 2
        inicio = datetime.now().astimezone()
        seleccion_local = [clave for clave in seleccion if clave not in FUENTES_DIRECTAS_GOOGLE]
        estado_antes = capturar_hoja() if seleccion_local else {'encabezados': {}, 'valores': {}}
        resultados = [ejecutar_fuente(clave=clave, fuente=FUENTES[clave], estricto=args.estricto) for clave in seleccion]
        errores = [resultado['fuente'] for resultado in resultados if resultado['estado'] == 'ERROR']
        omitidas = [resultado['fuente'] for resultado in resultados if resultado['estado'] == 'OMITIDA']
        finalizacion = None
        if not errores and seleccion_local:
            finalizacion = finalizar_excel(seleccion_local)
            if finalizacion['estado'] == 'ERROR':
                errores.append('formulas_excel')
        elif not errores:
            finalizacion = {'estado': 'NO_REQUERIDO', 'detalle': 'La fuente se carga directamente en Google Sheets por encabezados'}
        actualizaciones_google = preparar_actualizaciones_google(estado_antes, capturar_hoja()) if seleccion_local and (not errores) else []
        if 'dnrpa' in seleccion_local:
            originales_dnrpa = set(CAMPOS_GOOGLE_POR_FUENTE['dnrpa'])
            prefijos_dnrpa = ('patentamiento de automóviles', 'transferencias de automoviles', 'transferencias de automóviles', 'patentamiento de motos', 'transferencias de motos', 'trasnferencias de motos')
            for item in actualizaciones_google:
                item['valores'] = {titulo: valor for titulo, valor in item['valores'].items() if not str(titulo).casefold().startswith(prefijos_dnrpa) or titulo in originales_dnrpa}
            actualizaciones_google = [x for x in actualizaciones_google if x['valores']]
        for resultado in resultados:
            if resultado.get('fuente') in FUENTES_DIRECTAS_GOOGLE and resultado.get('estado') == 'OK' and resultado.get('periodo') and isinstance(resultado.get('valores_extraidos'), dict) and resultado['valores_extraidos']:
                periodo_erroneo = resultado.get('limpiar_periodo_erroneo')
                if periodo_erroneo:
                    actualizaciones_google.append({'hoja': HOJA_DESTINO, 'periodo': periodo_erroneo, 'valores': {titulo: '' for titulo in resultado['valores_extraidos']}})
                actualizaciones_google.append({'hoja': HOJA_DESTINO, 'periodo': resultado['periodo'], 'valores': resultado['valores_extraidos']})
        if seleccion_local:
            estado_despues = capturar_hoja()
            for item in preparar_actualizaciones_forzadas(estado_despues, seleccion_local, resultados, args.forzar_google):
                existente = next((actualizacion for actualizacion in actualizaciones_google if actualizacion['hoja'] == item['hoja'] and actualizacion['periodo'] == item['periodo']), None)
                if existente:
                    existente['valores'].update(item['valores'])
                    if item.get('permitir_reemplazos'):
                        existente['permitir_reemplazos'] = True
                else:
                    actualizaciones_google.append(item)
        # DNRPA escribe directo en Google Sheets desde scripts/dnrpa/sheet_dnrpa.py;
        # no se reenvía a n8n para evitar una doble escritura.
        prefijos_dnrpa_sheet = ('patentamiento de automóviles', 'transferencias de automoviles', 'transferencias de automóviles', 'patentamiento de motos', 'transferencias de motos', 'trasnferencias de motos')
        for item in actualizaciones_google:
            item['valores'] = {titulo: valor for titulo, valor in item['valores'].items() if not str(titulo).strip().casefold().startswith(prefijos_dnrpa_sheet)}
        actualizaciones_google = [item for item in actualizaciones_google if item['valores']]
        if 'eph' in seleccion:
            for item in actualizaciones_google:
                item['valores'] = {titulo: valor for titulo, valor in item['valores'].items() if not str(titulo).strip().casefold().startswith('dif pp. tasa de ')}
            actualizaciones_google = [item for item in actualizaciones_google if item['valores']]
        formato_sheet = {'estado': 'NO_CONFIGURADO', 'celdas': 0}
        if actualizaciones_google and os.environ.get('GOOGLE_APPLICATION_CREDENTIALS'):
            try:
                from scripts.comun.formato_observatorio import preparar_formatos
                formato_sheet = preparar_formatos(actualizaciones_google)
            except Exception as error:
                formato_sheet = {'estado': 'ERROR', 'detalle': str(error)}
                errores.append('formato_sheet')
        resultados_con_cambios = [resultado for resultado in resultados if resultado.get('hubo_cambios')]
        notificacion = {'hubo_cambios': bool(resultados_con_cambios), 'fuentes_modificadas': [resultado['fuente'] for resultado in resultados_con_cambios], 'total_celdas_actualizadas': sum((resultado.get('celdas_actualizadas', 0) for resultado in resultados_con_cambios)), 'detalle': [{'fuente': resultado['fuente'], 'nombre': resultado['nombre'], 'periodo': resultado.get('periodo'), 'celdas_actualizadas': resultado.get('celdas_actualizadas', 0), 'cambios': resultado.get('cambios', [])} for resultado in resultados_con_cambios], 'cambios_locales_no_confirmados_en_sheet': True}
        resumen = {'inicio': inicio.isoformat(timespec='seconds'), 'fin': datetime.now().astimezone().isoformat(timespec='seconds'), 'estado': 'ERROR' if errores else 'OK', 'solicitadas': len(seleccion), 'ejecutadas': len(resultados) - len(omitidas), 'errores': errores, 'omitidas': omitidas, 'notificacion': notificacion, 'actualizaciones_google': actualizaciones_google, 'formato_sheet': formato_sheet, 'finalizacion_excel': finalizacion, 'resultados': resultados}
        carpeta_n8n = RAIZ / 'salidas_n8n'
        carpeta_n8n.mkdir(parents=True, exist_ok=True)
        destino_n8n = carpeta_n8n / 'resumen_general.json'
        temporal_n8n = carpeta_n8n / 'resumen_general.tmp'
        temporal_n8n.write_text(json.dumps(resumen, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
        temporal_n8n.replace(destino_n8n)
        print('RESUMEN_N8N=' + json.dumps(resumen, ensure_ascii=False, separators=(',', ':')))
        print(f"ESTADO={resumen['estado']}")
        return 0
    return main()
if __name__ == '__main__':
    if '--motor-interno' in sys.argv[1:]:
        sys.argv.remove('--motor-interno')
        raise SystemExit(ejecutar_motor_integrado())
    raise SystemExit(main())
