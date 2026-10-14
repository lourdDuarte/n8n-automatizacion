
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

import xlrd

RAIZ=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(RAIZ/"scripts"/"complejos"))
from comun import ARCHIVO_BASE, abrir_base, descargar, emitir, escribir, fila_periodo, guardar_reporte, guardar_si_cambia

MESES={"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,"julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
TITULOS=[
"IS Sector privado registrado Var% i.m.","IS Sector privado registrado Número índice octubre 2016=100","IS Sector público Var% i.m.","IS Sector público Número índice octubre 2016=100","IS total registrado Var% i.m.","IS total registrado Número índice octubre 2016=100","IS sector privado no registrado Var% i.m.","IS sector privado no registrado Número índice octubre 2016=100","TOTAL IS Var% i.m.","TOTAL IS Número índice octubre 2016=100",
"IS Sector privado registrado Var% i.a.","IS Sector privado registrado con respecto a diciembre del año anterior (acum)","IS Sector público Var% i.a.","IS sector público con respecto diciembre del año anterior (acum)","IS total registrado Var% i.a.","IS total registrado con respecto a diciembre del año anterior (acum)","IS sector privado no registrado Var% i.a.","IS sector privado no registrado con respecto a diciembre del año anterior (acum)","TOTAL IS Var% i.a.","TOTAL IS con respecto a diciembre del año anterior (acum)",
"IS sector público nacional Var% i.m.","IS sector público nacional Var% i.a.","IS sector público nacional con respecto a diciembre del año anterior","IS sector público provincial Var% i.m.","IS sector público provincial Var% i.a.","IS sector público provincial con respecto a diciembre del año anterior"]

def descubrir():
    """Planilla vigente del índice de salarios aunque el INDEC cambie el nombre: prueba los
    nombres fechados de los últimos meses y los archivos de la página de salarios; usa la
    que trae el período más reciente en los tres cuadros."""
    sys.path.insert(0,str(RAIZ/"scripts"/"comun"))
    from fuentes_web import PAGINAS_INDEC, candidatos, elegir_mas_reciente
    hoy=datetime.now(); fechados=[]
    for retroceso in range(0,4):
        indice=hoy.year*12+hoy.month-1-retroceso; anio=indice//12; mes=indice%12+1
        fechados.append(f"https://www.indec.gob.ar/ftp/cuadros/sociedad/variaciones_salarios_{mes:02d}_{anio%100:02d}.xls")
    def procesar(contenido,url):
        libro=xlrd.open_workbook(file_contents=contenido)
        comunes=set(leer_hoja(libro.sheet_by_name("Cuadro 1"),2,12))&set(leer_hoja(libro.sheet_by_name("Cuadro 3"),2,8))
        return max(comunes),(url,contenido)
    urls=candidatos(fechados,PAGINAS_INDEC["salarios"],incluir=("salario",),extensiones=(".xls",))
    (url,contenido),_,_=elegir_mas_reciente(urls,procesar)
    return url,contenido

def leer_hoja(hoja,inicio,fin):
    anio=None; salida={}
    for r in range(hoja.nrows):
        a=hoja.cell_value(r,0); m=str(hoja.cell_value(r,1)).strip().casefold()
        if isinstance(a,(int,float)) and 1900<a<2100: anio=int(a)
        if anio and m in MESES:
            valores=[]
            for c in range(inicio,fin):
                v=hoja.cell_value(r,c)
                if not isinstance(v,(int,float)): valores.append(0.0)
                else: valores.append(round(float(v),2))
            salida[date(anio,MESES[m],1)]=valores
    return salida

def extraer():
    url,contenido=descubrir(); libro=xlrd.open_workbook(file_contents=contenido)
    c1=leer_hoja(libro.sheet_by_name("Cuadro 1"),2,12); c2=leer_hoja(libro.sheet_by_name("Cuadro 2"),2,12); c3=leer_hoja(libro.sheet_by_name("Cuadro 3"),2,8)
    comunes=sorted(set(c1)&set(c2)&set(c3))
    if not comunes: raise ValueError("No hay períodos comunes en los tres cuadros de salarios")
    return url,{p:c1[p]+c2[p]+c3[p] for p in comunes}

def vacia(ws,fila):
    from comun import columna_titulo
    return all(ws.cell(fila,columna_titulo(ws,t)).value in (None,"",0,0.0,"0.00") for t in TITULOS)

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--base",type=Path,default=RAIZ/"base"/ARCHIVO_BASE); ap.add_argument("--solo-extraer",action="store_true"); args=ap.parse_args()
    url,series=extraer(); ultimo=max(series); resultado={"fuente":"Índice de salarios","periodo":ultimo.isoformat(),"url_planilla":url,"valores":dict(zip(TITULOS,series[ultimo]))}
    if not args.solo_extraer:
        libro,ws=abrir_base(args.base); cambios=[]; coincidencias=[]; filas=[]
        # Completa meses faltantes desde 2026 y siempre controla el último publicado.
        candidatos=[]
        for p in sorted(series):
            if p.year<2026: continue
            try: f=fila_periodo(ws,p.year,p.month)
            except ValueError: continue
            if vacia(ws,f) or p==ultimo: candidatos.append((p,f))
        for p,fila in candidatos:
            antes=len(cambios)
            for t,v in zip(TITULOS,series[p]): escribir(ws,fila,t,v,cambios,coincidencias)
            filas.append({"periodo":p.isoformat(),"fila":fila,"celdas_actualizadas":len(cambios)-antes})
        respaldo=guardar_si_cambia(libro,args.base,cambios,"salarios")
        resultado["actualizacion"]={"archivo_excel":str(args.base.resolve()),"filas":filas,"celdas_actualizadas":len(cambios),"cambios":cambios,"coincidencias":coincidencias,"respaldo":respaldo}
    resultado["reporte"]=str(guardar_reporte(RAIZ,"salarios",ultimo,resultado).resolve()); emitir(resultado); return 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except Exception as e: print(f"ERROR={e}"); print("ESTADO=ERROR"); raise SystemExit(1)
