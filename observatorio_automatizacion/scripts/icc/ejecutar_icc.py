from __future__ import annotations

import argparse
import html
import re
import sys
from datetime import date
from pathlib import Path

import fitz

RAIZ=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(RAIZ/"scripts"/"complejos"))
from comun import ARCHIVO_BASE, abrir_base, columna_titulo, descargar, emitir, escribir, fila_periodo, guardar_reporte, guardar_si_cambia, normalizar, numero

PAGINA="https://www.utdt.edu/ver_contenido.php?id_contenido=2575&id_item_menu=4982"
MESES={"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,"julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
TITULOS={
"interior":"Indice de Confianza del Consumidor por Regiones: Interior",
"interior_mensual":"Indice de Confianza del Consumidor Interior Variación respecto al mes anterior",
"interior_anual":"Indice de Confianza del Consumidor Interior Variación respecto a mismo mes del año anterior",
"acumulado":"Indice de Confianza del Consumidor Interior Acumulado del año",
"nacional":"Indice de Confianza del Consumidor Nacional",
"nacional_mensual":"Indice de Confianza del Consumidor Nacional Variación respecto al mes anterior N",
"nacional_anual":"Indice de Confianza del Consumidor Nacional Variación respecto a mismo mes del año anterior N"}

def periodo_documento(contenido):
    doc=fitz.open(stream=contenido,filetype="pdf")
    texto="\n".join(p.get_text("text") for p in list(doc)[:3])
    m=re.search(r"(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+(20\d{2})",texto,re.I)
    if not m: return None
    return date(int(m.group(2)),MESES[normalizar(m.group(1))],1)

def descubrir_pdf():
    pagina=descargar(PAGINA).decode("utf-8","replace")
    enlaces=re.findall(r'href=["\']([^"\']*download\.php\?fname=[^"\']+\.pdf)["\']',pagina,re.I)
    if not enlaces: raise ValueError("No se encontraron informes PDF del ICC")
    candidatos=[]
    # La página de UTDT ya presenta los informes del más reciente al más antiguo.
    # El número interno de download.php es un identificador opaco: NO es una fecha
    # ni crece necesariamente de un mes al siguiente. Ordenarlo numéricamente hizo
    # que septiembre de 2026 quedara afuera y se siguiera leyendo agosto.
    # Conservamos el orden publicado y verificamos el período escrito dentro de
    # cada PDF. Se inspeccionan doce enlaces para tolerar elementos destacados.
    unicos=list(dict.fromkeys(enlaces))
    for enlace in unicos[:12]:
        url=html.unescape(enlace)
        if url.startswith("/"): url="https://www.utdt.edu"+url
        elif not url.startswith("http"): url="https://www.utdt.edu/"+url
        try:
            contenido=descargar(url,60)
            if contenido.startswith(b"%PDF"):
                periodo=periodo_documento(contenido)
                if periodo: candidatos.append((periodo,url,contenido))
        except Exception: pass
    if not candidatos: raise FileNotFoundError("No se pudo descargar un informe válido del ICC")
    _,url,contenido=max(candidatos,key=lambda x:x[0])
    return url,contenido

def bloque_icc(texto,seccion):
    pos=texto.find(seccion)
    if pos<0: raise ValueError(f"No se encontró la sección {seccion} del ICC")
    fragmento=texto[pos:pos+4500]
    m=re.search(r"ICCCIF\s*\n\s*([-0-9,.]+)\s*\n\s*([-0-9,.]+)\s*\n\s*([-0-9,.]+)\s*\n\s*([-0-9,.]+)%\s*\n\s*([-0-9,.]+)%",fragmento,re.I)
    if not m: raise ValueError(f"No se pudo interpretar la fila ICCCIF de {seccion}")
    return [numero(x) for x in m.groups()]

def mes_anterior(periodo):
    indice=periodo.year*12+periodo.month-2
    return date(indice//12,indice%12+1,1)

def extraer():
    url,contenido=descubrir_pdf(); doc=fitz.open(stream=contenido,filetype="pdf"); textos=[p.get_text("text") for p in doc]; todo="\n".join(textos)
    m=re.search(r"(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+(20\d{2})",todo,re.I)
    if not m: raise ValueError("No se pudo detectar el período del ICC")
    periodo=date(int(m.group(2)),MESES[normalizar(m.group(1))],1)
    tabla=next((t for t in textos if "NACIONAL" in t and "INTERIOR" in t and "ICCCIF" in t),None)
    if not tabla: raise ValueError("No se encontró la tabla nacional e interior del ICC")
    nacional=bloque_icc(tabla,"NACIONAL"); interior=bloque_icc(tabla,"INTERIOR")
    valores={TITULOS["interior"]:interior[0],TITULOS["interior_mensual"]:interior[3],TITULOS["interior_anual"]:interior[4],TITULOS["nacional"]:nacional[0],TITULOS["nacional_mensual"]:nacional[3],TITULOS["nacional_anual"]:nacional[4]}
    return periodo,url,valores

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--base",type=Path,default=RAIZ/"base"/ARCHIVO_BASE); ap.add_argument("--solo-extraer",action="store_true"); args=ap.parse_args()
    periodo,url,valores=extraer(); resultado={"fuente":"ICC - UTDT","periodo":periodo.isoformat(),"url_informe":url,"valores":valores,"calculos_propios":[TITULOS["acumulado"]]}
    if not args.solo_extraer:
        libro,ws=abrir_base(args.base); fila=fila_periodo(ws,periodo.year,periodo.month); cambios=[]; coincidencias=[]
        for t,v in valores.items(): escribir(ws,fila,t,v,cambios,coincidencias)
        if periodo.month==1: formula=f"={ws.cell(fila, columna_titulo(ws,TITULOS['interior_mensual'])).coordinate}"
        else:
            c_ac=columna_titulo(ws,TITULOS['acumulado']); c_m=columna_titulo(ws,TITULOS['interior_mensual'])
            anterior=mes_anterior(periodo)
            fila_anterior=fila_periodo(ws,anterior.year,anterior.month)
            formula=f"=((1+{ws.cell(fila_anterior,c_ac).coordinate}/100)*(1+{ws.cell(fila,c_m).coordinate}/100)-1)*100"
        escribir(ws,fila,TITULOS["acumulado"],formula,cambios,coincidencias)
        # El acumulado es un cálculo propio de la base, no un dato original de
        # UTDT. Se conserva en el Excel local como respaldo, pero no se incluye
        # en "valores": así n8n envía al Sheet solamente los seis datos fuente
        # y nunca transmite referencias frágiles como AAC93 o AAA94.
        resultado["calculos_locales"]={TITULOS["acumulado"]:formula}
        respaldo=guardar_si_cambia(libro,args.base,cambios,"icc")
        resultado["actualizacion"]={"archivo_excel":str(args.base.resolve()),"fila":fila,"celdas_actualizadas":len(cambios),"cambios":cambios,"coincidencias":coincidencias,"respaldo":respaldo}
    resultado["reporte"]=str(guardar_reporte(RAIZ,"icc",periodo,resultado).resolve()); emitir(resultado); return 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except Exception as e: print(f"ERROR={e}"); print("ESTADO=ERROR"); raise SystemExit(1)
