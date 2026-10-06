"""Infraestructura exclusiva de PBI/EPH; no modifica otras fuentes."""
from __future__ import annotations
import argparse
import io
import json
import math
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import urllib.parse
import urllib.request
from datetime import date, datetime
from pathlib import Path
import pdfplumber
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parents[2]
HOJA = 'BASE DE DATOS'

def normalizar(valor):
    t = unicodedata.normalize('NFKD', str(valor or '')).casefold()
    return re.sub(r'[^a-z0-9]+', ' ', ''.join(c for c in t if not unicodedata.combining(c))).strip()

def numero(valor):
    if isinstance(valor, bool):
        raise ValueError('Booleano recibido como dato')
    t = str(valor).strip().replace('−', '-').replace('–', '-').replace('%', '')
    t = re.sub(r'\s*\([1-9]\)\s*$', '', t)
    if t == '-':
        return 0.0  # INDEC: cero absoluto, no dato faltante.
    if not re.fullmatch(r'-?\d[\d.]*(?:,\d+)?', t):
        raise ValueError(f'Dato no numérico o no disponible: {valor!r}')
    n = float(t.replace('.', '').replace(',', '.'))
    if not math.isfinite(n):
        raise ValueError('Dato no finito')
    return n

def destino(anio, trimestre):
    if type(anio) is not int or type(trimestre) is not int or trimestre not in (1, 2, 3, 4):
        raise ValueError('Año/trimestre inválido')
    return date(anio + (trimestre == 4), 3 if trimestre == 4 else 3 * (trimestre + 1), 1)

def trimestre_objetivo(hoy=None):
    hoy = hoy or date.today()
    if hoy.month < 3:
        return hoy.year - 1, 3
    if hoy.month < 6:
        return hoy.year - 1, 4
    return hoy.year, hoy.month // 3 - 1

def periodo_pdf(textos, tipo):
    # La segunda página incluye índices y series históricas; el período del
    # informe se toma exclusivamente de la portada.
    portada = normalizar(textos[0] if textos else '')
    requerido = 'informe de avance' if tipo == 'pbi' else 'mercado de trabajo'
    if requerido not in portada:
        raise ValueError(f'El PDF no es el informe de {tipo}: no se admite otro indicador')
    matches = re.findall(r'(primer|primero|segundo|tercer|tercero|cuarto|[1-4]) trimestre (?:de )?(20\d{2})', portada)
    periodos = {({'primer':1,'primero':1,'segundo':2,'tercer':3,'tercero':3,'cuarto':4}.get(q, int(q) if q.isdigit() else 0), int(y)) for q,y in matches}
    if len(periodos) != 1:
        raise ValueError(f'Período de portada no encontrado o ambiguo: {sorted(periodos)}')
    q,y = periodos.pop()
    return y,q

def pedir(url):
    p = urllib.parse.urlparse(url)
    if p.scheme != 'https' or p.hostname not in ('www.indec.gob.ar', 'indec.gob.ar'):
        raise ValueError(f'Enlace no oficial INDEC: {url}')
    req = urllib.request.Request(url, headers={'User-Agent':'Mozilla/5.0 (indicadores-PBI-EPH)'})
    with urllib.request.urlopen(req, timeout=60) as r:
        final = urllib.parse.urlparse(r.geturl())
        if final.hostname not in ('www.indec.gob.ar', 'indec.gob.ar'):
            raise ValueError('INDEC redirigió fuera de su dominio')
        return r.read()

def descubrir(pagina, patron):
    html = pedir(pagina).decode('utf-8', errors='replace')
    enlaces = re.findall(r'''href\s*=\s*["']([^"']+\.pdf(?:\?[^"']*)?)["']''', html, re.I)
    urls = {urllib.parse.urljoin(pagina, e.replace('&amp;', '&')) for e in enlaces if re.search(patron, e, re.I)}
    def orden_url(url):
        nombre=urllib.parse.urlparse(url).path.rsplit('/',1)[-1]
        m=re.search(r'pib_(\d{2})_(\d{2})',nombre,re.I)
        if m: return 2000+int(m[2]),int(m[1])
        m=re.search(r'eph_([1-4])trim[_-]?(\d{2})',nombre,re.I)
        if m: return 2000+int(m[2]),3*int(m[1])
        return 0,0
    urls=sorted(urls,key=orden_url,reverse=True)
    if not urls:
        raise ValueError('No se encontraron enlaces al informe en la página INDEC; no se usó una URL adivinada')
    candidatos = []
    for url in urls[:4]:
        b = pedir(url)
        if not b.startswith(b'%PDF'):
            raise ValueError(f'El enlace no devuelve PDF: {url}')
        with pdfplumber.open(io.BytesIO(b)) as p:
            textos = [x.extract_text() or '' for x in p.pages[:2]]
        tipo = 'eph' if 'mercado_trabajo' in patron else 'pbi'
        y,q = periodo_pdf(textos, tipo)
        candidatos.append(((y,q),url,b))
    return max(candidatos, key=lambda x: x[0])

def indice_columnas(ws):
    indice = {}
    for c in range(1, ws.max_column + 1):
        k = normalizar(ws.cell(1,c).value)
        if k:
            indice.setdefault(k, []).append(c)
    return indice

def columna(indice, titulo):
    cs = indice.get(normalizar(titulo), [])
    if len(cs) != 1:
        raise ValueError(f'Encabezado ausente o duplicado: {titulo}')
    return cs[0]

def mes_celda(v):
    if isinstance(v, (date, datetime)):
        return v.year, v.month
    t = normalizar(v)
    m = re.fullmatch(r'(20\d{2}) (\d{1,2})(?: \d{1,2})?', t)
    if m:
        return int(m[1]),int(m[2])
    meses = {'ene':1,'feb':2,'mar':3,'abr':4,'may':5,'jun':6,'jul':7,'ago':8,'sep':9,'sept':9,'oct':10,'nov':11,'dic':12}
    m = re.fullmatch(r'([a-z]+) (\d{2}|\d{4})', t)
    if m and m[1] in meses:
        return int(m[2])+(2000 if len(m[2])==2 else 0),meses[m[1]]
    return None

def fila_mes(ws, indice, mes):
    c = columna(indice, 'PERIODO')
    filas = [r for r in range(2, ws.max_row+1) if mes_celda(ws.cell(r,c).value)==(mes.year,mes.month)]
    if len(filas)!=1:
        raise ValueError(f'Fila mensual {mes:%Y-%m} ausente o duplicada')
    return filas[0]

def actualizar(resultado, base, formulas=()):
    if not base.is_file():
        raise FileNotFoundError(f'No existe la base: {base}')
    wb = load_workbook(base, data_only=False)
    try:
        ws = wb[HOJA]
        ind = indice_columnas(ws)
        mes = destino(resultado['anio'], resultado['trimestre'])
        fila = fila_mes(ws, ind, mes)
        plan, detalle, conservadas = [], [], []
        usados = set()
        for titulo,nuevo in resultado['valores'].items():
            c = columna(ind,titulo)
            if c in usados or c == columna(ind,'PERIODO'):
                raise ValueError(f'Destino repetido o reservado: {titulo}')
            usados.add(c)
            celda = ws.cell(fila,c)
            if celda.data_type == 'f':
                raise ValueError(f'Un dato original apunta a una fórmula: {celda.coordinate}')
            if isinstance(nuevo,bool) or not isinstance(nuevo,(float,int)) or not math.isfinite(nuevo):
                raise ValueError(f'Dato inválido: {titulo}')
            detalle.append({'titulo':ws.cell(1,c).value,'valor':nuevo})
            if celda.value != nuevo:
                plan.append({'celda':celda.coordinate,'titulo':ws.cell(1,c).value,'antes':celda.value,'despues':nuevo})
        # Solo cálculos EPH identificados explícitamente. No se toca una fórmula existente.
        for titulo,fuente,retroceso in formulas:
            c,orig = columna(ind,titulo),columna(ind,fuente)
            celda = ws.cell(fila,c)
            if celda.data_type == 'f':
                conservadas.append(celda.coordinate)
                detalle.append({'titulo':ws.cell(1,c).value,'valor':celda.value})
                continue
            anterior = date(mes.year-(mes.month<=retroceso), (mes.month-retroceso-1)%12+1,1) if retroceso==3 else date(mes.year-1,mes.month,1)
            previa = fila_mes(ws,ind,anterior)
            a,b = f'{get_column_letter(orig)}{fila}',f'{get_column_letter(orig)}{previa}'
            nuevo = f'=IF(AND(ISNUMBER({a}),ISNUMBER({b})),{a}-{b},"")'
            detalle.append({'titulo':ws.cell(1,c).value,'valor':nuevo})
            if celda.value != nuevo:
                plan.append({'celda':celda.coordinate,'titulo':ws.cell(1,c).value,'antes':celda.value,'despues':nuevo})
        respaldo = None
        if plan:
            carpeta = RAIZ / 'respaldos' / 'pbi_eph'
            carpeta.mkdir(parents=True,exist_ok=True)
            respaldo = carpeta / f'{base.stem}_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx'
            shutil.copy2(base,respaldo)
            for cambio in plan:
                ws[cambio['celda']] = cambio['despues']
            fd,nombre = tempfile.mkstemp(prefix='pbi_eph_',suffix='.xlsx',dir=base.parent)
            os.close(fd)
            try:
                wb.save(nombre)
                os.replace(nombre,base)
            finally:
                if os.path.exists(nombre):
                    os.unlink(nombre)
        return {'archivo_excel':str(base.resolve()),'hoja':HOJA,'fila_localizada':fila,
                'celdas_actualizadas':len(plan),'hubo_cambios':bool(plan),'cambios':plan,
                'detalle_completo':detalle,'formulas_conservadas':conservadas,
                'respaldo':str(respaldo) if respaldo else None}
    finally:
        wb.close()

def ejecutar(tipo, pagina, patron, extractor, formulas=()):
    for flujo in (sys.stdout,sys.stderr):
        if hasattr(flujo,'reconfigure'):
            flujo.reconfigure(encoding='utf-8',errors='replace')
    parser = argparse.ArgumentParser(description=f'Extrae y actualiza {tipo} sin modificar otros indicadores')
    parser.add_argument('--base',type=Path,default=RAIZ/'base'/'BASE DE DATOS INDICADORES PROPIOS.xlsx')
    parser.add_argument('--pdf',type=Path)
    parser.add_argument('--solo-extraer',action='store_true')
    parser.add_argument('--periodo',help='Trimestre estadístico explícito, ejemplo: 2026-T2')
    args = parser.parse_args()
    try:
        objetivo = trimestre_objetivo()
        if args.periodo:
            m = re.fullmatch(r'(20\d{2})-T([1-4])', args.periodo)
            if not m:
                raise ValueError('--periodo debe tener formato 2026-T2')
            objetivo = int(m[1]),int(m[2])
        if args.pdf:
            url,b = args.pdf.resolve().as_uri(),args.pdf.read_bytes()
        else:
            periodo,url,b = descubrir(pagina,patron)
            if periodo < objetivo:
                resultado = {'fuente':tipo,'periodo':None,'periodo_estadistico':f'{objetivo[0]}-T{objetivo[1]}',
                    'estado_publicacion':'PENDIENTE','diagnostico':'La página todavía no enlaza el informe del trimestre solicitado; no se cargó el anterior en el mes nuevo.',
                    'ultimo_publicado':f'{periodo[0]}-T{periodo[1]}',
                    'actualizacion':{'celdas_actualizadas':0,'cambios':[],'detalle_completo':[]}}
                print(json.dumps(resultado,ensure_ascii=False)); print('ESTADO=SIN_PUBLICACION_NUEVA'); return 0
            if periodo != objetivo:
                raise ValueError(f'La publicación {periodo} no coincide con el objetivo {objetivo}')
        resultado = extractor(b,url)
        if args.periodo and (resultado['anio'],resultado['trimestre']) != objetivo:
            raise ValueError('El PDF no coincide con --periodo')
        if not args.solo_extraer:
            resultado['actualizacion'] = actualizar(resultado,args.base,formulas)
        carpeta = RAIZ/'scripts'/'reportes'/tipo
        carpeta.mkdir(parents=True,exist_ok=True)
        reporte = carpeta/f'{tipo}_{resultado["anio"]}_T{resultado["trimestre"]}.json'
        resultado['reporte'] = str(reporte.resolve())
        reporte.write_text(json.dumps(resultado,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(resultado,ensure_ascii=False,indent=2)); print('ESTADO=OK'); return 0
    except Exception as error:
        print(json.dumps({'fuente':tipo,'periodo':None,'estado':'ERROR','diagnostico':str(error)},ensure_ascii=False))
        print(f'ERROR={error}',file=sys.stderr); print('ESTADO=ERROR'); return 1
