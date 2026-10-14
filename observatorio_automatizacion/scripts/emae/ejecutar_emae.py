from __future__ import annotations

import argparse
import io
import sys
from datetime import date
from pathlib import Path

import xlrd

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "scripts" / "complejos"))
from comun import ARCHIVO_BASE, abrir_base, descargar, emitir, escribir, fila_periodo, guardar_reporte, guardar_si_cambia

URL_MENSUAL = "https://www.indec.gob.ar/ftp/cuadros/economia/sh_emae_mensual_base2004.xls"
URL_ACTIVIDAD = "https://www.indec.gob.ar/ftp/cuadros/economia/sh_emae_actividad_base2004.xls"
MESES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,"julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}

TITULOS_GENERAL = [
    "EMAE Indice Serie Original 2004=100",
    "EMAE Indice Serie Desestacionalizada 2004=100",
    "EMAE Indice Serie Tendencia-Ciclo 2004=100",
    "EMAE Respecto al mes anterior (desestacionalizado) N",
    "EMAE Respecto a igual mes del año anterior (serie original) N",
    "EMAE Acumulado del año respecto al acumulado de igual periodo del año anterior N",
    "EMAE Respecto al mes anterior (tendencia-ciclo) N",
]
TITULOS_ACTIVIDAD = [
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Agricultura, ganadería, caza y silvicultura",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Pesca",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Explotación de minas y canteras",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Industria manufacturera",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Electricidad, gas y agua",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Construcción",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Comercio mayorista, minorista y reparaciones",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Hoteles y restaurantes",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Transporte y comunicaciones",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Intermediación financiera",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Actividades inmobiliarias, empresariales y de alquiler",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Administración pública y defensa; planes de seguridad social de afiliación obligatoria",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Enseñanza",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Servicios sociales y de salud",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior. Otras actividades de servicios comunitarios, sociales y personales",
    "EMAE por sector de actividad económica. Var% respecto a igual período del año anterior.Impuestos netos de subsidios",
]

# Columnas agregadas en el Observatorio: números índice por sector (hoja "Tabla Letras").
# Solo existen en el Google Sheet (no en el Excel local), por eso van aparte.
SECTORES_INDICE = [
    "Agricultura, ganadería, caza y silvicultura", "Pesca.", "Explotación de minas y canteras",
    "Industria manufacturera", "Electricidad, gas y agua", "Construcción",
    "Comercio mayorista, minorista y reparaciones", "Hoteles y restaurantes",
    "Transporte y comunicaciones", "Intermediación financiera",
    "Actividades inmobiliarias, empresariales y de alquiler",
    "Administración pública y defensa; planes de seguridad social de afiliación obligatoria",
    "Enseñanza", "Servicios sociales y de salud",
    "Otras actividades de servicios comunitarios, sociales y personales", "Impuestos netos de subsidios",
]
TITULOS_INDICE = ["EMAE por sector de actividad económica. Números índice, base 2004=100. " + s for s in SECTORES_INDICE]

def filas_periodos(hoja):
    anio = None
    salida = []
    for r in range(hoja.nrows):
        a, m = hoja.cell_value(r,0), str(hoja.cell_value(r,1)).strip().casefold()
        if isinstance(a,(int,float)) and 1900 < a < 2100: anio=int(a)
        if anio and m in MESES: salida.append((date(anio,MESES[m],1),r))
    return salida

def _planilla(preferida, incluir, hoja):
    """Busca la planilla del EMAE aunque el INDEC cambie el nombre: prueba la conocida y
    las de la página del EMAE que tengan la hoja esperada; usa la del período más reciente."""
    sys.path.insert(0, str(RAIZ / "scripts" / "comun"))
    from fuentes_web import PAGINAS_INDEC, candidatos, elegir_mas_reciente
    def procesar(contenido, url):
        libro = xlrd.open_workbook(file_contents=contenido)
        return max(p for p, _ in filas_periodos(libro.sheet_by_name(hoja))), libro
    urls = candidatos([preferida], PAGINAS_INDEC["emae"], incluir=incluir)
    libro, url, _ = elegir_mas_reciente(urls, procesar)
    return libro, url

def extraer():
    lm, _ = _planilla(URL_MENSUAL, ("emae", "mensual"), "Tabla"); hm=lm.sheet_by_name("Tabla")
    la, _ = _planilla(URL_ACTIVIDAD, ("emae", "actividad"), "Tabla Var Letras"); ha=la.sheet_by_name("Tabla Var Letras")
    fm=filas_periodos(hm); fa=dict(filas_periodos(ha)); periodo,r=max(fm,key=lambda x:x[0])
    if periodo not in fa: raise ValueError("Las dos planillas EMAE no terminan en el mismo período")
    indice_actual=float(hm.cell_value(r,2))
    actual_anio=[float(hm.cell_value(rr,2)) for p,rr in fm if p.year==periodo.year and p.month<=periodo.month]
    anterior_anio=[float(hm.cell_value(rr,2)) for p,rr in fm if p.year==periodo.year-1 and p.month<=periodo.month]
    acumulada=(sum(actual_anio)/sum(anterior_anio)-1)*100
    general=[indice_actual,float(hm.cell_value(r,4)),float(hm.cell_value(r,6)),float(hm.cell_value(r,5)),float(hm.cell_value(r,3)),acumulada,float(hm.cell_value(r,7))]
    actividades=[float(ha.cell_value(fa[periodo],c)) for c in range(2,18)]
    hi=la.sheet_by_name("Tabla Letras"); fi=dict(filas_periodos(hi))
    if periodo not in fi: raise ValueError("La hoja de números índice por sector no llega al mismo período")
    indices=[float(hi.cell_value(fi[periodo],c)) for c in range(2,18)]
    # Control: la variación interanual que publica INDEC debe salir de estos índices.
    anterior=date(periodo.year-1,periodo.month,1)
    if anterior in fi:
        for c,var in zip(range(2,18),actividades):
            base=float(hi.cell_value(fi[anterior],c))
            if base and abs((indices[c-2]/base-1)*100-var)>0.6:
                raise ValueError(f"Los índices por sector no coinciden con sus variaciones (columna {c})")
    return periodo,[round(v,1) for v in general],[round(v,1) for v in actividades],[round(v,1) for v in indices]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--base",type=Path,default=RAIZ/"base"/ARCHIVO_BASE); ap.add_argument("--solo-extraer",action="store_true"); args=ap.parse_args()
    periodo,general,actividades,indices=extraer(); resultado={"fuente":"EMAE","periodo":periodo.isoformat(),"valores":dict(zip(TITULOS_GENERAL+TITULOS_ACTIVIDAD,general+actividades)),"valores_solo_sheet":dict(zip(TITULOS_INDICE,indices))}
    if not args.solo_extraer:
        libro,ws=abrir_base(args.base); fila=fila_periodo(ws,periodo.year,periodo.month); cambios=[]; coincidencias=[]
        for t,v in resultado["valores"].items(): escribir(ws,fila,t,v,cambios,coincidencias)
        respaldo=guardar_si_cambia(libro,args.base,cambios,"emae")
        resultado["actualizacion"]={"archivo_excel":str(args.base.resolve()),"fila":fila,"celdas_actualizadas":len(cambios),"cambios":cambios,"coincidencias":coincidencias,"respaldo":respaldo}
    resultado["reporte"]=str(guardar_reporte(RAIZ,"emae",periodo,resultado).resolve()); emitir(resultado); return 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except Exception as e: print(f"ERROR={e}"); print("ESTADO=ERROR"); raise SystemExit(1)
