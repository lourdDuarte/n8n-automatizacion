from __future__ import annotations

import argparse
import io
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

RAIZ=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(RAIZ/"scripts"/"complejos"))
from comun import ARCHIVO_BASE, abrir_base, columna_titulo, descargar, emitir, escribir, fila_periodo, guardar_reporte, guardar_si_cambia, normalizar, numero

URL="https://www.indec.gob.ar/ftp/cuadros/economia/serie_supermercados.xlsx"
T={
"const_f":"Ventas totales de supermercados (a precios constantes y en millones de pesos) F",
"const_n":"Ventas totales de supermercados (a precios constantes y en millones de pesos) N",
"const_m_f":"Var % de las ventas de supermercados a precios constantes con respecto al mes anterior F",
"const_m_n":"Var % de las ventas de supermercados con respecto al mes anterior N",
"const_a_f":"Var % de las ventas de supermercados a precios constantes con respecto al mismo mes del año anterior F",
"const_a_n":"Var % de las ventas de supermercados a precios constantes con respecto al mismo mes del año anterior N",
"corr_f":"Ventas totales de supermercados (en millones de pesos) F","corr_n":"Ventas totales de supermercados (en millones de pesos) N",
"acum_f":"Acumulado Ventas totales de supermercados (en millones de pesos) F",
"corr_m_f":"Variación porcentual de las ventas de supermercados con respecto al mes anterior F","corr_m_n":"Variación porcentual de las ventas de supermercados con respecto al mes anterior N",
"corr_a_f":"Variación porcentual de las ventas de supermercados con respecto al mismo mes del año anterior F","corr_a_n":"Variación porcentual de las ventas de supermercados con respecto al mismo mes del año anterior N",
"op_f":"Cantidad de operaciones realizadas en supermercados F","op_m_f":"Var% intermensual de la cantidad de operaciones realizadas en supermercados F","op_a_f":"Var% interanual de la cantidad de operaciones realizadas en supermercados F",
"op_n":"Cantidad de operaciones realizadas en supermercados N","op_m_n":"Var% intermensual de la cantidad de operaciones realizadas en supermercados N","op_a_n":"Var% interanual de la cantidad de operaciones realizadas en supermercados N",
"ipc":"Indice de precios al consumidor- Nivel general N"}
RUBROS=["Bebidas","Almacén","Panadería","Lácteos","Carnes","Verdulería y frutería","Alimentos preparados y rotisería","Limpieza y perfumería","Indumentaria, calzado y textiles para el hogar","Electrónicos y art. para el hogar","Otros artículos"]

def fila_fecha(ws,periodo):
    for r in range(1,ws.max_row+1):
        v=ws.cell(r,1).value
        if isinstance(v,(datetime,date)) and v.year==periodo.year and v.month==periodo.month:return r
    raise ValueError(f"No existe {periodo:%Y-%m} en {ws.title}")
def ultimo(ws):
    fechas=[v for (v,) in ws.iter_rows(min_col=1,max_col=1,values_only=True) if isinstance(v,(datetime,date))]
    return date(max(fechas).year,max(fechas).month,1)
def inicio_grupo(ws,nombre):
    for c in range(1,ws.max_column+1):
        if normalizar(ws.cell(3,c).value)==normalizar(nombre):return c
    raise ValueError(f"No se encontró el grupo {nombre} en {ws.title}")
def rubro_titulo(rubro,zona,constante=False):
    if constante:
        pref="Vtas en supermercados a precios constantes"
        variantes={"Lácteos":f"{pref}- Art. seleccionados: Lácteos (en miles de $) F","Almacén":f"{pref} - Art. seleccionados: Almacén (en miles de $) F","Bebidas":f"{pref} - Art. seleccionados: Bebidas (en miles de $) F","Panadería":f"{pref} - Art. seleccionados: Panadería (en miles de $) F","Carnes":f"{pref} - Art. seleccionados: Carnes (en miles de $) F","Verdulería y frutería":f"{pref}- Art. seleccionados: Verdulería y frutería (en miles de $) F","Alimentos preparados y rotisería":f"{pref}- Art. seleccionados: Alimentos preparados y rotisería (en miles de $) F","Limpieza y perfumería":f"{pref} - Art. seleccionados: Limpieza y perfumería (en miles de $) F","Indumentaria, calzado y textiles para el hogar":f"{pref} - Art. seleccionados: Indumentaria, calzado y textiles para el hogar (en miles de $) F","Electrónicos y art. para el hogar":f"{pref}- Art. seleccionados: Electrónicos y art. para el hogar (en miles de $) F","Otros artículos":f"{pref}- Art. seleccionados: Otros artículos (en miles de $) F"}
        return variantes[rubro]
    return f"Ventas en supermercados - Art. seleccionados: {rubro} (en miles de $) {zona}"
def numero_publicado(valor):
    # INDEC usa "s" cuando el dato está protegido por secreto estadístico.
    # No debe convertirse en cero porque cero sería un dato distinto.
    if normalizar(valor) in {"s","...","…","///","-",""}: return None
    return numero(valor)
def desplazar_meses(periodo,cantidad):
    indice=periodo.year*12+periodo.month-1+cantidad
    return date(indice//12,indice%12+1,1)
def formula_variacion(ws,fila,periodo,titulo_base,meses_atras):
    comparacion=desplazar_meses(periodo,-meses_atras)
    fila_comparacion=fila_periodo(ws,comparacion.year,comparacion.month)
    c=columna_titulo(ws,titulo_base)
    return f"={ws.cell(fila,c).coordinate}/{ws.cell(fila_comparacion,c).coordinate}*100-100"

def extraer():
    libro=load_workbook(io.BytesIO(descargar(URL)),data_only=True,read_only=False)
    c3,c5,c6=libro["Cuadro 3."],libro["Cuadro 5."],libro["Cuadro 6."]
    periodo=ultimo(c3)
    if ultimo(c5)!=periodo or ultimo(c6)!=periodo: raise ValueError("Los cuadros de supermercados no terminan en el mismo período")
    r3,r5,r6=fila_fecha(c3,periodo),fila_fecha(c5,periodo),fila_fecha(c6,periodo)
    f5=inicio_grupo(c5,"Formosa"); f6=inicio_grupo(c6,"Formosa")
    valores={T["corr_n"]:numero(c3.cell(r3,2).value),T["const_n"]:numero(c3.cell(r3,3).value),T["corr_f"]:numero(c5.cell(r5,f5).value)/1000,T["op_n"]:numero(c6.cell(r6,7).value),T["op_f"]:numero(c6.cell(r6,f6+5).value)}
    for i,rubro in enumerate(RUBROS,1):
        # En Cuadro 5 la columna 2 es el total nacional y las once columnas
        # siguientes son los rubros (Bebidas comienza en la columna 3).
        nacional=numero_publicado(c5.cell(r5,2+i).value)
        formosa=numero_publicado(c5.cell(r5,f5+i).value)
        if nacional is not None: valores[rubro_titulo(rubro,"N")]=nacional
        # INDEC identifica algunos datos provinciales con secreto estadístico.
        # La base histórica representa esos casos con cero.
        valores[rubro_titulo(rubro,"F")]=formosa if formosa is not None else 0.0
    return periodo,valores

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--base",type=Path,default=RAIZ/"base"/ARCHIVO_BASE); ap.add_argument("--solo-extraer",action="store_true"); args=ap.parse_args()
    periodo,valores=extraer(); resultado={"fuente":"INDEC - Supermercados","periodo":periodo.isoformat(),"url_planilla":URL,"valores":valores,"valores_originales":dict(valores)}
    if not args.solo_extraer:
        libro,ws=abrir_base(args.base); fila=fila_periodo(ws,periodo.year,periodo.month); cambios=[]; coincidencias=[]
        for t,v in valores.items(): escribir(ws,fila,t,round(v,3),cambios,coincidencias)
        ipc=numero(ws.cell(fila,columna_titulo(ws,T["ipc"])).value)
        if ipc<=0: raise ValueError(f"Falta el IPC nacional del período {periodo:%Y-%m}; no se pueden calcular precios constantes de Formosa")
        # Estos doce valores deben formar parte de resultado["valores"] para
        # que ejecutar_complejos.py los envíe al Google Sheet del Observatorio.
        valores[T["const_f"]]=round(valores[T["corr_f"]]/ipc*100,1)
        escribir(ws,fila,T["const_f"],valores[T["const_f"]],cambios,coincidencias)
        for rubro in RUBROS:
            titulo_corriente=rubro_titulo(rubro,"F")
            corriente=valores[titulo_corriente]
            # corriente ya está expresado en miles de pesos y el encabezado
            # constante también está expresado en miles: no se divide por 1000.
            titulo_constante=rubro_titulo(rubro,"F",True)
            valores[titulo_constante]=round(corriente/ipc*100,2)
            escribir(ws,fila,titulo_constante,valores[titulo_constante],cambios,coincidencias)
        formulas={T["const_m_f"]:formula_variacion(ws,fila,periodo,T["const_f"],1),T["const_m_n"]:formula_variacion(ws,fila,periodo,T["const_n"],1),T["const_a_f"]:formula_variacion(ws,fila,periodo,T["const_f"],12),T["const_a_n"]:formula_variacion(ws,fila,periodo,T["const_n"],12),T["corr_m_f"]:formula_variacion(ws,fila,periodo,T["corr_f"],1),T["corr_m_n"]:formula_variacion(ws,fila,periodo,T["corr_n"],1),T["corr_a_f"]:formula_variacion(ws,fila,periodo,T["corr_f"],12),T["corr_a_n"]:formula_variacion(ws,fila,periodo,T["corr_n"],12),T["op_m_f"]:formula_variacion(ws,fila,periodo,T["op_f"],1),T["op_a_f"]:formula_variacion(ws,fila,periodo,T["op_f"],12),T["op_m_n"]:formula_variacion(ws,fila,periodo,T["op_n"],1),T["op_a_n"]:formula_variacion(ws,fila,periodo,T["op_n"],12)}
        c_corr=columna_titulo(ws,T["corr_f"]); c_ac=columna_titulo(ws,T["acum_f"])
        if periodo.month>1:
            anterior=desplazar_meses(periodo,-1)
            fila_anterior=fila_periodo(ws,anterior.year,anterior.month)
            formulas[T["acum_f"]]=f"={ws.cell(fila_anterior,c_ac).coordinate}+{ws.cell(fila,c_corr).coordinate}"
        else:
            formulas[T["acum_f"]]=f"={ws.cell(fila,c_corr).coordinate}"
        for t,v in formulas.items(): escribir(ws,fila,t,v,cambios,coincidencias)
        respaldo=guardar_si_cambia(libro,args.base,cambios,"supermercados")
        resultado["actualizacion"]={"archivo_excel":str(args.base.resolve()),"fila":fila,"ipc_utilizado":ipc,"celdas_actualizadas":len(cambios),"cambios":cambios,"coincidencias":coincidencias,"respaldo":respaldo}
    resultado["reporte"]=str(guardar_reporte(RAIZ,"supermercados",periodo,resultado).resolve()); emitir(resultado); return 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except Exception as e: print(f"ERROR={e}"); print("ESTADO=ERROR"); raise SystemExit(1)
