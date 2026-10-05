"""Tasas y poblaciones oficiales EPH. No procesa microdatos ni salarios."""
from __future__ import annotations
import io
import re
import sys
from pathlib import Path
import pdfplumber
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from trimestrales.comun import destino, ejecutar, normalizar, numero, periodo_pdf

REGIONES = ('Formosa','NEA','N')
TASAS = ('empleo','desocupación','actividad','subocupación')

def titulo_tasa(tasa,region):
    prefijo = 'tasa' if tasa=='desocupación' else 'Tasa'
    if region=='Formosa':
        return f'{prefijo} de {tasa} (ciudad de Formosa) F'
    if region=='NEA':
        nombre = 'Región NEA' if tasa in ('empleo','desocupación') else 'Region NEA'
        return f'{prefijo} de {tasa} ({nombre})'
    return f'{prefijo} de {tasa} N'

def formulas():
    salida=[]
    for tasa in TASAS:
        t = normalizar(tasa)
        for region in REGIONES:
            if region=='Formosa':
                base=f'Dif pp. tasa de {t} (ciudad de Formosa)'; fin=' F'
            elif region=='NEA':
                nombre='Región NEA' if tasa in ('empleo','desocupación') else 'Region NEA'
                base=f'Dif pp. tasa de {t} ({nombre})'; fin=''
            else:
                base=f'Dif pp. tasa de {t}'; fin=' N'
            for meses,comparacion in ((3,'trimestre anterior'),(12,'mismo trimestre del año anterior')):
                titulo=f'{base} con respecto al {comparacion}{fin}'
                # Error ortográfico conservado en el encabezado real del Observatorio.
                if tasa=='desocupación' and region=='NEA' and meses==12:
                    titulo='Dif pp. tasa de desocupacion Región NEA) con respecto al mismo trimestre del año anterior'
                salida.append((titulo,titulo_tasa(tasa,region),meses))
    return tuple(salida)

def titulo_poblacion(nombre,region):
    if nombre=='PEA':
        return {'Formosa':'PEA (ciudad de Formosa -en miles) F','NEA':'PEA (Region NEA -en miles)','N':'PEA (en miles) N'}[region]
    return {
        ('ocupada','Formosa'):'Población ocupada (ciudad de Formosa- en miles) F',
        ('ocupada','NEA'):'Población ocupada (Region NEA- en miles)',
        ('ocupada','N'):'Población ocupada (en miles) N',
        ('desocupada','Formosa'):'Población desocupada (ciudad de Formosa- En miles) F',
        ('desocupada','NEA'):'Población desocupada (Region NEA- En miles)',
        ('desocupada','N'):'Población desocupada (en miles) N',
        ('subocupada','Formosa'):'Población subocupada (ciudad de Formosa - en miles) F',
        ('subocupada','NEA'):'Población subocupada (Region NEA - en miles)',
        ('subocupada','N'):'Población subocupada (en miles) N',
    }[nombre,region]

def region_fila(texto):
    t=normalizar(texto)
    if t=='formosa': return 'Formosa'
    if t in ('nea','noreste','region noreste','region nea'): return 'NEA'
    if t in ('total 31 aglomerados urbanos','total de 31 aglomerados urbanos','total 31 aglomerados','total de aglomerados urbanos'): return 'N'
    return None

def campos_encabezado(celdas,es_tasa):
    campos={}
    for i,v in enumerate(celdas):
        t=normalizar(v)
        if es_tasa:
            # Excluir demandantes y subdivisiones de subocupación.
            if 'demandante' in t or 'sobreocup' in t: continue
            for k in TASAS:
                if re.fullmatch(r'(?:tasa de )?'+normalizar(k),t): campos[i]=k
        else:
            equivalencias={'poblacion economicamente activa':'PEA','pea':'PEA',
                'poblacion ocupada':'ocupada','ocupados':'ocupada',
                'poblacion desocupada':'desocupada','desocupados':'desocupada',
                'poblacion subocupada':'subocupada','subocupados':'subocupada'}
            if t in equivalencias: campos[i]=equivalencias[t]
    return campos

def leer_tabla(tabla,contexto,datos,evidencias,pagina):
    if not tabla: return
    es_tasa='tasa' in normalizar(contexto)
    es_poblacion='poblacion' in normalizar(contexto) or 'poblaciones' in normalizar(contexto)
    if not es_tasa and not es_poblacion: return
    # No convertir una tabla de porcentajes en miles de personas.
    if es_poblacion and not es_tasa:
        n=normalizar(contexto)
        if 'miles' in n: factor=1.0
        elif 'personas' in n or 'habitantes' in n: factor=0.001
        else: return
    else: factor=1.0
    for inicio in range(min(8,len(tabla))):
        for profundidad in (1,2,3,4):
            cabecera=tabla[inicio:inicio+profundidad]
            if len(cabecera)!=profundidad: continue
            ancho=max(len(r) for r in cabecera)
            combinadas=[' '.join(str(r[i] or '') for r in cabecera if i<len(r)).strip() for i in range(ancho)]
            campos=campos_encabezado(combinadas,es_tasa)
            requeridos=set(TASAS) if es_tasa else {'PEA','ocupada','desocupada','subocupada'}
            if set(campos.values())!=requeridos: continue
            for fila in tabla[inicio+profundidad:]:
                etiqueta=' '.join(str(v or '') for i,v in enumerate(fila) if i not in campos).strip()
                region=region_fila(etiqueta)
                if not region: continue
                for c,nombre in campos.items():
                    if c>=len(fila): raise ValueError('Fila EPH incompleta')
                    valor=numero(fila[c])*factor
                    if valor<0 or es_tasa and valor>100:
                        raise ValueError(f'Dato EPH fuera de rango: {region}/{nombre}')
                    clave=(region,nombre,'tasa' if es_tasa else 'poblacion')
                    if clave in datos and abs(datos[clave]-valor)>1e-6:
                        raise ValueError(f'Dos datos oficiales incompatibles para {clave}; no se eligió al azar')
                    datos[clave]=valor
                    evidencias[str(clave)]={'pagina':pagina,'encabezado':combinadas[c],'etiqueta':etiqueta,'factor_unidad':factor}

def numeros_fila(linea):
    """Números impresos en una fila oficial, sin depender de celdas PDF."""
    return re.findall(r'(?<!\w)[\-\u2212\u2013]?\d[\d.]*,?\d*(?!\w)', linea)

def valor_tasa(token):
    return numero(token)

def valor_poblacion(token):
    # En los cuadros EPH la unidad es miles y el punto es separador de miles:
    # 14.616 significa 14.616 miles de personas, no 14,616.
    limpio = str(token).strip().replace('.', '').replace(',', '.')
    return float(limpio)

def extraer_cuadros_oficiales(textos,datos,evidencias):
    """Lee por encabezados y etiquetas los cuadros 3.1 y 3.2 de INDEC."""
    regiones = {
        'total 31 aglomerados urbanos':'N',
        'noreste':'NEA',
        'formosa':'Formosa',
    }
    for pagina,texto in enumerate(textos,1):
        n = normalizar(texto)
        # INDEC cambia la numeración y el título de los cuadros entre
        # publicaciones. Se identifican por sus encabezados y contenido, no
        # por "Cuadro 3.1/3.2" ni por una frase editorial exacta.
        contiene_regiones = all(x in n for x in (
            'total 31 aglomerados urbanos', 'noreste', 'formosa'
        ))
        es_tasas = contiene_regiones and all(x in n for x in (
            'area geografica', 'actividad', 'empleo', 'desocupacion',
            'subocupacion'
        ))
        es_poblacion = contiene_regiones and all(x in n for x in (
            'poblacion', 'economicamente', 'activa', 'ocupada', 'desocupada',
            'subocupada'
        ))
        # Una página de población también puede mencionar tasas en notas o
        # índices. La unidad/encabezado de población tiene prioridad.
        if es_poblacion and ('poblacion en miles' in n or 'en miles' in n):
            es_tasas = False
        if not es_tasas and not es_poblacion:
            continue
        for linea in texto.splitlines():
            clave_linea = normalizar(linea)
            region = None
            etiqueta = None
            for inicio,codigo in regiones.items():
                if clave_linea.startswith(inicio+' '):
                    region,etiqueta = codigo,inicio
                    break
            if not region:
                continue
            tokens = numeros_fila(linea)
            # El nombre nacional contiene "31"; no es un dato de la fila.
            if region == 'N' and tokens and tokens[0] == '31':
                tokens = tokens[1:]
            if es_tasas and len(tokens) == 7:
                valores = [valor_tasa(v) for v in tokens]
                seleccion = {
                    'actividad':valores[0], 'empleo':valores[1],
                    'desocupación':valores[2], 'subocupación':valores[4],
                }
                for nombre,valor in seleccion.items():
                    datos[region,nombre,'tasa']=valor
                    evidencias[str((region,nombre,'tasa'))]={
                        'pagina':pagina,'cuadro':'3.1','etiqueta':etiqueta,
                        'unidad':'porcentaje'
                    }
            elif es_poblacion and len(tokens) == 6:
                valores = [valor_poblacion(v) for v in tokens]
                seleccion = {
                    'PEA':valores[1], 'ocupada':valores[2],
                    'desocupada':valores[3], 'subocupada':valores[5],
                }
                for nombre,valor in seleccion.items():
                    datos[region,nombre,'poblacion']=valor
                    evidencias[str((region,nombre,'poblacion'))]={
                        'pagina':pagina,'cuadro':'3.2','etiqueta':etiqueta,
                        'unidad':'miles de personas'
                    }

def extraer(b,url):
    datos,evidencias={},{}
    with pdfplumber.open(io.BytesIO(b)) as pdf:
        textos=[p.extract_text() or '' for p in pdf.pages]
        y,q=periodo_pdf(textos,'eph')
        extraer_cuadros_oficiales(textos,datos,evidencias)
        for numero_pagina,p in enumerate(pdf.pages,1):
            # Cada cuadro tiene su contexto propio, no las unidades de otra tabla de la página.
            titulos=[m for m in re.finditer(r'Cuadro\s+\d+(?:\.\d+)*\.', textos[numero_pagina-1])]
            if not titulos: continue
            for tabla in p.find_tables():
                x0,top,x1,bottom=tabla.bbox
                contexto=p.crop((0,max(0,top-130),p.width,min(p.height,top+18))).extract_text() or ''
                leer_tabla(tabla.extract(),contexto,datos,evidencias,numero_pagina)
            # Los cuadros INDEC sin bordes se detectan por alineación de palabras.
            for tabla in p.find_tables({'vertical_strategy':'text','horizontal_strategy':'text','min_words_vertical':3}):
                x0,top,x1,bottom=tabla.bbox
                contexto=p.crop((0,max(0,top-130),p.width,min(p.height,top+18))).extract_text() or ''
                leer_tabla(tabla.extract(),contexto,datos,evidencias,numero_pagina)
    requeridos={(r,t,'tasa') for r in REGIONES for t in TASAS}|{(r,p,'poblacion') for r in REGIONES for p in ('PEA','ocupada','desocupada','subocupada')}
    faltan=requeridos-set(datos)
    if faltan:
        raise ValueError('EPH: faltan datos oficiales identificados por encabezado/unidad: '+', '.join('/'.join(k) for k in sorted(faltan))+'. No se escribió la base ni se estimaron microdatos.')
    valores={}
    for r in REGIONES:
        for t in TASAS: valores[titulo_tasa(t,r)]=datos[r,t,'tasa']
        for p in ('PEA','ocupada','desocupada','subocupada'): valores[titulo_poblacion(p,r)]=datos[r,p,'poblacion']
        pea=datos[r,'PEA','poblacion']; ocupada=datos[r,'ocupada','poblacion']; desocupada=datos[r,'desocupada','poblacion']
        if pea<=0 or abs(pea-ocupada-desocupada)>max(2,pea*0.002):
            raise ValueError(f'Poblaciones EPH no concilian PEA=ocupados+desocupados: {r}')
    # Los datos trimestrales pertenecen al mes de cierre del trimestre:
    # T1=marzo, T2=junio, T3=septiembre y T4=diciembre.
    periodo_correcto=f'{y}-{q*3:02d}-01'
    periodo_erroneo_anterior=destino(y,q).isoformat()
    return {'fuente':'eph','anio':y,'trimestre':q,'periodo_estadistico':f'{y}-T{q}',
        'periodo':periodo_correcto,
        'limpiar_periodo_erroneo':(
            periodo_erroneo_anterior
            if periodo_erroneo_anterior != periodo_correcto else None
        ),
        'url_informe':url,'estado_publicacion':'PUBLICADO',
        'valores':valores,'evidencias':evidencias,'validaciones':{'campos_originales':24,'calculos':24,
        'cobertura':'Formosa, región NEA y total 31 aglomerados urbanos','unidad_poblaciones':'miles de personas'}}

if __name__=='__main__':
    raise SystemExit(ejecutar('eph','https://www.indec.gob.ar/Nivel4/Tema/4/31/58',r'mercado_trabajo_eph_',extraer,formulas()))
