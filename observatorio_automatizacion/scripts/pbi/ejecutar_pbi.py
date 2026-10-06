"""Ejecutor PBI independiente. Guardar en scripts/pbi/ejecutar_pbi.py."""
from __future__ import annotations
import argparse
import io
import json
import logging
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
logging.getLogger('pdfminer').setLevel(logging.ERROR)
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
    # El dato se guarda en el último mes del trimestre estadístico,
    # no en el mes de publicación del informe.
    return date(anio, 3 * trimestre, 1)

def trimestre_objetivo(hoy=None):
    hoy = hoy or date.today()
    if hoy.month < 3:
        return hoy.year - 1, 3
    if hoy.month < 6:
        return hoy.year - 1, 4
    return hoy.year, hoy.month // 3 - 1

def periodo_pdf(textos, tipo):
    # El período oficial está en la portada. La página 2 contiene un índice
    # con referencias históricas que no deben considerarse períodos candidatos.
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

def resolver_enlace(pagina, enlace):
    enlace = enlace.replace('&amp;', '&').strip()
    # INDEC usa rutas ftp/... y uploads/... referidas a la raíz del sitio.
    if re.match(r'^(?:\./)?(?:ftp|uploads)/', enlace, re.I):
        origen = urllib.parse.urlsplit(pagina)
        return urllib.parse.urljoin(f'{origen.scheme}://{origen.netloc}/', enlace.removeprefix('./'))
    return urllib.parse.urljoin(pagina, enlace)

def enlaces_informes(html, pagina):
    enlaces = re.findall(r'''href\s*=\s*["']([^"']+\.pdf(?:\?[^"']*)?)["']''', html, re.I)
    urls = set()
    for enlace in enlaces:
        url = resolver_enlace(pagina,enlace)
        p=urllib.parse.urlsplit(url)
        nombre=p.path.rsplit('/',1)[-1]
        # Solo informes fechados: excluye pib_revision_*, metodologías y anexos.
        if p.hostname in ('www.indec.gob.ar','indec.gob.ar') and re.fullmatch(r'pib_(0[1-9]|1[0-2])_\d{2}[A-Za-z0-9_-]*\.pdf',nombre,re.I):
            urls.add(url)
    return sorted(urls,key=lambda u:(int(re.search(r'pib_\d{2}_(\d{2})',u,re.I)[1]),int(re.search(r'pib_(\d{2})_',u,re.I)[1])),reverse=True)

def descubrir(pagina, patron):
    html = pedir(pagina).decode('utf-8', errors='replace')
    urls = enlaces_informes(html,pagina)
    if not urls:
        raise ValueError('No se encontraron enlaces al informe en la página INDEC; no se usó una URL adivinada')
    candidatos = []
    for url in urls[:1]:
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
TITULOS = (
    'PBI Var. % desestacionalizada respecto del trimestre anterior.',
    'PBI Var. % interanual. Igual trimestre del año anterior.',
    'PBI Acumulado del año respecto a igual acumulado del año anterior.',
    'PBI Consumo Privado desestacionalizado N',
    'PBI Consumo Público desestacionalizado N',
    'PBI Inversión desestacionalizado N',
    'PBI Exportaciones desestacionalizado N',
    'PBI Importaciones desestacionalizado N',
    'PBI Consumo Privado Interanual N',
    'PBI Consumo Público Interanual N',
    'PBI Inversión Interanual N',
    'PBI Exportaciones Interanual N',
    'PBI Importaciones Interanual N',
)

def variacion_trimestral_resumen(resumen):
    """Obtiene la variación desestacionalizada sin depender de una frase fija."""
    texto = resumen.replace('\u00ad', '').replace('\u00a0', ' ')
    # Primero se usa la oración narrativa: evita confundir el dato del PIB con
    # porcentajes de gráficos o de los componentes de la demanda.
    texto_oraciones = re.sub(r'\s+', ' ', texto)
    for oracion in re.split(r'(?<=[.!?])\s+', texto_oraciones):
        clave = normalizar(oracion)
        if (
            ('producto interno bruto' in clave or re.search(r'\bpib\b', clave))
            and 'desestacionalizad' in clave
            and 'trimestre' in clave
        ):
            marcas_pib = list(re.finditer(r'producto\s+interno\s+bruto|\bPIB\b', oracion, re.I))
            segmento = oracion[marcas_pib[-1].end():] if marcas_pib else oracion
            cifras = re.findall(r'(?<!\d)([\-\u2212\u2013]?\s*\d{1,3},\d+)\s*%', segmento)
            if len(cifras) == 1:
                valor = numero(re.sub(r'\s+', '', cifras[0]))
                if valor > 0 and any(v in clave for v in (
                    'cayo','caida','descendio','descenso','disminuyo',
                    'disminucion','bajo','redujo','contrajo','contraccion'
                )):
                    valor = -valor
                return valor
    referencias_pib = [
        m.start() for m in re.finditer(r'producto\s+interno\s+bruto|\bPIB\b', texto, re.I)
    ]
    referencias_trimestre = [
        m.start() for m in re.finditer(
            r'respecto\s+(?:del|al)\s+trimestre\s+anterior|'
            r'(?:serie|t[eé]rminos?)\s+desestacionalizad[oa]s?',
            texto,
            re.I,
        )
    ]
    porcentajes = list(re.finditer(r'(?<!\d)(-?\d{1,3},\d+)\s*%', texto))
    candidatos = []
    for porcentaje in porcentajes:
        if not referencias_pib or not referencias_trimestre:
            continue
        distancia_pib = min(abs(porcentaje.start() - pos) for pos in referencias_pib)
        distancia_trim = min(abs(porcentaje.start() - pos) for pos in referencias_trimestre)
        if distancia_pib <= 550 and distancia_trim <= 350:
            inicio_oracion = max(
                texto.rfind('.', 0, porcentaje.start()),
                texto.rfind('\n', 0, porcentaje.start()),
            ) + 1
            finales = [
                posicion for posicion in (
                    texto.find('.', porcentaje.end()),
                    texto.find('\n', porcentaje.end()),
                ) if posicion >= 0
            ]
            fin_oracion = min(finales) if finales else len(texto)
            oracion = normalizar(texto[inicio_oracion:fin_oracion])
            es_trimestral = (
                'trimestre anterior' in oracion
                or 'desestacionalizad' in oracion
            )
            penalizacion = 1000 if 'interanual' in oracion and not es_trimestral else 0
            if es_trimestral:
                penalizacion -= 500
            # La referencia a trimestre anterior/desestacionalización es la
            # señal principal. "PIB" puede aparecer una sola vez encabezando
            # varias cifras distintas en una infografía.
            candidatos.append((penalizacion + distancia_trim, distancia_pib, porcentaje))
    if not candidatos:
        raise ValueError(
            'No se reconoció la variación trimestral del PIB en el resumen ejecutivo'
        )
    _, _, elegido = min(candidatos, key=lambda item: (item[0], item[1]))
    valor = numero(elegido.group(1))
    contexto = normalizar(texto[max(0, elegido.start()-180):elegido.end()+180])
    if valor > 0 and any(v in contexto for v in (
        'cayo','caida','descendio','descenso','disminuyo','disminucion',
        'bajo','redujo','contrajo','contraccion'
    )):
        valor = -valor
    return valor


def fila_variaciones_pib(textos, anio, trimestre):
    """Lee las cuatro variaciones publicadas en la tabla principal del PIB."""
    etiquetas = {
        1: 'Primer trimestre',
        2: 'Segundo trimestre',
        3: 'Tercer trimestre',
        4: 'Cuarto trimestre',
    }
    candidatos = []
    patron_numero = r'[\-\u2212\u2013]?\s*\d+(?:[.,]\d+)?'
    for texto in textos:
        limpio = normalizar(texto)
        if (
            'estimacion preliminar del pib' not in limpio
            or 'acumulado del ano' not in limpio
            or str(anio) not in limpio
        ):
            continue
        etiqueta_normalizada = normalizar(etiquetas[trimestre])
        anio_bloque = None
        for linea in texto.splitlines():
            marca_anio = re.fullmatch(r'\s*(20\d{2})\s*(?:\(\d+\))?\s*', linea)
            if marca_anio:
                anio_bloque = int(marca_anio.group(1))
                continue
            if etiqueta_normalizada not in normalizar(linea):
                continue
            if anio_bloque != anio:
                continue
            resto = re.sub(
                rf'^.*?{re.escape(etiquetas[trimestre])}', '', linea,
                count=1, flags=re.I,
            )
            tokens = re.findall(patron_numero, resto)
            if len(tokens) == 4:
                candidatos.append([
                    numero(re.sub(r'\s+', '', valor)) for valor in tokens
                ])
    candidatos_unicos = []
    for candidato in candidatos:
        if candidato not in candidatos_unicos:
            candidatos_unicos.append(candidato)
    if len(candidatos_unicos) != 1:
        raise ValueError(
            'No se reconoció de forma única la fila oficial de variaciones del PIB '
            f'(candidatas: {candidatos_unicos})'
        )
    # desestacionalizada, tendencia-ciclo, interanual y acumulada.
    return candidatos_unicos[0]

def extraer(b,url):
    with pdfplumber.open(io.BytesIO(b)) as p:
        textos = [x.extract_text() or '' for x in p.pages]
    y,q = periodo_pdf(textos,'pbi')
    resumenes = [t for t in textos[2:] if 'resumen ejecutivo' in normalizar(t) and 'informe de avance' in normalizar(t)]
    if len(resumenes)!=1:
        raise ValueError('Resumen ejecutivo PBI ausente o ambiguo')
    resumen = resumenes[0]
    filas = []
    for linea in resumen.splitlines():
        tokens = linea.split()
        if normalizar(linea).startswith(('igual trimestre ', 'respecto al ')):
            tokens=tokens[-6:]
        if len(tokens)==6 and tokens[3]=='///':
            try:
                filas.append([numero(v) for i,v in enumerate(tokens) if i!=3])
            except ValueError:
                pass
    if len(filas)!=2:
        raise ValueError('No se reconocieron las dos filas de cinco componentes de demanda en el resumen PBI')
    # INDEC presenta primero la variación interanual y luego la desestacionalizada.
    if normalizar(resumen).find('var interanual') > normalizar(resumen).find('var desestacionalizada'):
        raise ValueError('Cambió el orden de filas en el resumen ejecutivo')
    original,desest = filas
    # INDEC modifica con frecuencia la redacción y el orden visual de esta
    # oración. Se relaciona el porcentaje con los conceptos PIB y trimestre
    # anterior, en lugar de exigir una frase literal.
    sa_narrativa = variacion_trimestral_resumen(resumen)
    sa, _tendencia, interanual_tabla, acumulado = fila_variaciones_pib(textos, y, q)
    if not math.isclose(sa, sa_narrativa, rel_tol=0, abs_tol=0.11):
        raise ValueError('No concilia la variación trimestral entre resumen y tabla PIB')
    cuadros1 = []
    cuadros2 = []
    for t in textos:
        m = re.search(r'Cuadro 1\.\s*Oferta y demanda globales\.(.*)', t,re.S)
        if m and 'precios de 2004' in m[1][:160] and re.search(r'^B1b\s+Producto interno bruto',m[1],re.M):
            cuadros1.append(m[1].split('Fuente:')[0])
        m = re.search(r'Cuadro 2\.\s*Oferta y demanda globales\.(.*?)(?:Cuadro 3\.|\Z)',t,re.S)
        if m and 'porcentual' in m[1][:160] and re.search(r'^B1b\s+Producto interno bruto',m[1],re.M):
            cuadros2.append(m[1])
    if len(cuadros1)!=1 or len(cuadros2)!=1:
        raise ValueError('Cuadros originales 1 y 2 no encontrados o duplicados')
    def pib_fila(t):
        m=re.search(r'^B1b\s+Producto interno bruto\s+(.+)$',t,re.M)
        if not m:
            raise ValueError('No se reconoció B1b Producto interno bruto')
        return [numero(v) for v in m[1].split()]
    nivel = pib_fila(cuadros1[0])
    variaciones = pib_fila(cuadros2[0])
    # El cuadro nuevo usa una ventana móvil de seis columnas: trimestres del
    # año anterior desde el comparable, Total y trimestres corrientes.
    if len(nivel)!=6 or len(variaciones)!=6:
        raise ValueError('Cambió la estructura trimestral del cuadro PBI; no se adivinó la columna')
    anios = re.findall(r'\b(20\d{2})\s*\(1\)',cuadros1[0].split('Código')[0])
    if anios != [str(y-1),str(y)]:
        raise ValueError(f'Años de Cuadro 1 incompatibles: {anios}')
    if not math.isclose(variaciones[-1], interanual_tabla, rel_tol=0, abs_tol=0.11):
        raise ValueError('No concilia la variación interanual entre los cuadros PIB')
    valores = dict(zip(TITULOS,[sa,variaciones[-1],acumulado,*desest,*original]))
    if any(abs(v)>300 for v in valores.values()):
        raise ValueError('Variación PBI fuera del rango de control; requiere diagnóstico')
    return {'fuente':'pbi','anio':y,'trimestre':q,'periodo_estadistico':f'{y}-T{q}',
        'periodo':destino(y,q).isoformat(),'url_informe':url,'estado_publicacion':'PUBLICADO',
        'valores':valores,'validaciones':{'campos':13,'unidad':'variación porcentual',
        'acumulado':'valor publicado en la tabla oficial de variaciones del PIB; conciliado con los Cuadros 1 y 2'}}

if __name__=='__main__':
    raise SystemExit(ejecutar('pbi','https://www.indec.gob.ar/Nivel4/Tema/3/9/47',r'/pib_',extraer))
