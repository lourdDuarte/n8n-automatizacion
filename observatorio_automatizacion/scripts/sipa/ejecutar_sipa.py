"""Actualiza SIPA usando exclusivamente el último período oficial publicado."""
from __future__ import annotations
import argparse, io, json, math, re, sys, urllib.parse, urllib.request
from datetime import date, datetime
from pathlib import Path
from openpyxl import load_workbook

SCRIPTS_DIR=Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path: sys.path.insert(0,str(SCRIPTS_DIR))
from comun.columnas_por_titulo import resolver_columnas
PAGINA_SIPA="https://www.argentina.gob.ar/trabajo/estadisticas"; HOJA_BASE="BASE DE DATOS"; ARCHIVO_BASE="BASE DE DATOS INDICADORES PROPIOS.xlsx"
SECTORES=("Agricultura, ganaderÍa, \ncaza y silvicultura","Pesca","Explotación de \nminas y canteras","Industrias manufactureras","Suministro de electricidad, gas y agua","Construcción","Comercio y reparaciones","Hoteles y restaurantes","Transporte, almacenamiento y comunicación","Intermediación financiera","Actividades inmobiliarias, \nempresariales y de alquiler","Enseñanza","Servicios sociales \ny de salud","Servicios comunitarios, sociales y personales")
TITULOS=["Cant. Personas con trabajo registrado. Total país. Serie Original (en miles)","Var% i.m. Personas con trabajo registrado. Total país. Serie Original","Var% i.a. Personas con trabajo registrado. Total país. Serie Original","Var. absoluta i.m. Personas con trabajo registrado. Total país. Serie Original","Var. absoluta i.a. Personas con trabajo registrado. Total país. Serie Original","Cant. Personas con trabajo registrado. Total país. Serie desestacionalizada (en miles)","Var% i.m. Personas con trabajo registrado. Total país. Serie desestacionalizada","Var% i.a. Personas con trabajo registrado. Total país. Serie desestacionalizada","Var. absoluta i.m. Personas con trabajo registrado. Total país. Serie desestacionalizada","Var. absoluta i.a. Personas con trabajo registrado. Total país. Serie desestacionalizada","Cant. asalariados registrados del sector privado. Sin estacionalidad. En miles. N","Var% i.m. asalariados registrados del sector privado N","Var. absoluta i.m. asalariados registrados del sector privado N","Var % i.a. asalariados registrados del sector privado N","Var. absoluta i.a. asalariados registrados del sector privado (en miles) N"]
for indice,s in enumerate(SECTORES):
 base=f"asalariados registrados del sector privado (en miles), según rama de actividad S/E. {s} N"
 if indice == 0:
  variacion=f"asalariados registrados del sector privado, según rama de actividad S/E. {s} N"
 else:
  variacion=base
 TITULOS += [f"Cant. {base}",f"Var% i.m. {variacion}",f"Var% i.a. {variacion}"]
TITULOS += ["Remuneración MEDIANA de los trabajadores registrados del sector privado. Por todo concepto. A valores corrientes (Con estacionalidad) N","Remuneración PROMEDIO de los trabajadores registrados del sector privado. Por todo concepto. A valores corrientes (Con estacionalidad) N","Var. % i.m. remuneración PROMEDIO de los trabajadores registrados del sector privado. C/E","Var. % i.a. remuneración PROMEDIO de los trabajadores registrados del sector privado C/E N","Remuneración MEDIANA de los trabajadores registrados del sector privado Por todo concepto A valores corrientes (Sin estacionalidad) N","Remuneración PROMEDIO de los trabajadores registrados del sector privado Por todo concepto A valores corrientes S/E N","Var. % i.m. remuneración promedio de los trabajadores registrados del sector privado S/E N","Var. % i.a. remuneración promedio de los trabajadores registrados del sector privado S/E N","Cant. asalariados registrados del sector privado. Sin estacionalidad F","Var % i.m. asalariados registrados del sector privado S/E F","Var. absoluta i.m. asalariados registrados del sector privado S/E F","Var% i.a. asalariados registrados del sector privado S/E F","Var. absoluta i.a. asalariados registrados del sector privado S/E F"]
DESTINOS={f"S{i+1:02d}":t for i,t in enumerate(TITULOS)}
# Algunas versiones de la base agregan u omiten "(en miles)" en los títulos
# de las variaciones sectoriales. Ambas formas identifican la misma columna.
for indice,s in enumerate(SECTORES):
 base=f"asalariados registrados del sector privado (en miles), según rama de actividad S/E. {s} N"
 sin_unidad=f"asalariados registrados del sector privado, según rama de actividad S/E. {s} N"
 for desplazamiento,prefijo in ((1,"Var% i.m. "),(2,"Var% i.a. ")):
  clave=f"S{16+indice*3+desplazamiento:02d}"
  DESTINOS[clave]=(prefijo+base,prefijo+sin_unidad)

def raiz_proyecto(): return Path(__file__).resolve().parents[2]
def pedir(url,timeout=90):
 req=urllib.request.Request(url,headers={"User-Agent":"Mozilla/5.0 (automatizacion-indicadores/2.0)"})
 with urllib.request.urlopen(req,timeout=timeout) as r:return r.read()
def descubrir_planilla():
 html=pedir(PAGINA_SIPA).decode("utf-8",errors="replace"); links=re.findall(r'''href\s*=\s*["']([^"']*trabajoregistrado_[^"']*_estadisticas\.xlsx[^"']*)''',html,re.I)
 if not links: raise FileNotFoundError("No se encontró el anexo estadístico SIPA")
 cand=[]
 for e in links:
  u=urllib.parse.urljoin(PAGINA_SIPA,e.replace("&amp;","&")); m=re.search(r"trabajoregistrado_(\d{2})(\d{2})_estadisticas",u,re.I); cand.append((tuple(map(int,m.groups())) if m else (0,0),u))
 u=max(cand)[1]; b=pedir(u)
 if not b.startswith(b"PK"): raise ValueError("La descarga SIPA no es XLSX")
 return u,b
def periodo_celda(v):
 if isinstance(v,(date,datetime)): return date(v.year,v.month,1)
 if isinstance(v,str):
  m=re.fullmatch(r"([a-z]+)-(\d{2})\*?",v.strip().casefold()); meses={"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,"jul":7,"ago":8,"sep":9,"sept":9,"oct":10,"nov":11,"dic":12}
  if m and m[1] in meses:return date(2000+int(m[2]),meses[m[1]],1)
def serie(ws,col):
 out={}
 for row in ws.iter_rows(min_col=1,max_col=col,values_only=True):
  p=periodo_celda(row[0]); v=row[col-1] if len(row)>=col else None
  if p and isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(float(v)):out[p]=float(v)
 return out
def anterior(p,n):
 x=p.year*12+p.month-1-n; return date(x//12,x%12+1,1)
def pct(s,p,n):
 a=s.get(anterior(p,n)); return None if a in (None,0) else round((s[p]/a-1)*100,10)
def dif(s,p,n):
 a=s.get(anterior(p,n)); return None if a is None else round((s[p]-a)*1000,10)
def a_unidades(miles):
 """La hoja A.5.2 publica Formosa en miles (21,802); la columna F de la base va en personas (21802).
 Se pasa a personas sin redondear a miles, para que no quede 22."""
 v=round(float(miles)*1000,6)
 return int(round(v)) if abs(v-round(v))<1e-6 else v
def extraer(contenido,url):
 wb=load_workbook(io.BytesIO(contenido),read_only=True,data_only=True); req=("T.1","A.1","A.2.2","A.4","A.5.2"); falt=[x for x in req if x not in wb.sheetnames]
 if falt: raise ValueError(f"Faltan hojas SIPA: {falt}")
 t_orig=serie(wb["T.1"],2); t_sa=serie(wb["T.1"],3); privado=serie(wb["A.1"],3); sectores=[serie(wb["A.2.2"],c) for c in range(2,16)]; salarios=[serie(wb["A.4"],c) for c in range(2,6)]; formosa=serie(wb["A.5.2"],10); todas=[t_orig,t_sa,privado,*sectores,*salarios,formosa]
 comunes=set.intersection(*(set(s) for s in todas))
 if not comunes: raise ValueError("Las series SIPA no tienen un período oficial común")
 p=max(comunes); vals=[]
 for s in (t_orig,t_sa): vals += [s[p],pct(s,p,1),pct(s,p,12),dif(s,p,1),dif(s,p,12)]
 vals += [privado[p],pct(privado,p,1),dif(privado,p,1),pct(privado,p,12),dif(privado,p,12)]
 for s in sectores: vals += [s[p],pct(s,p,1),pct(s,p,12)]
 med_ce,med_se,prom_ce,prom_se=salarios; vals += [med_ce[p],prom_ce[p],pct(prom_ce,p,1),pct(prom_ce,p,12),med_se[p],prom_se[p],pct(prom_se,p,1),pct(prom_se,p,12)]; vals += [a_unidades(formosa[p]),pct(formosa,p,1),dif(formosa,p,1),pct(formosa,p,12),dif(formosa,p,12)]
 if len(vals)!=70: raise AssertionError(f"Se esperaban 70 valores y se obtuvieron {len(vals)}")
 valores={f"S{i+1:02d}":v for i,v in enumerate(vals) if v is not None}; faltantes=[TITULOS[i] for i,v in enumerate(vals) if v is None]
 return {"fuente":"SIPA - Trabajo registrado","periodo":p.isoformat(),"anio":p.year,"mes_numero":p.month,"url_planilla":url,"consultado_en":datetime.now().isoformat(timespec="seconds"),"valores":valores,"validaciones":{"columnas_esperadas":70,"columnas_disponibles":len(valores),"faltantes":faltantes,"ultimo_periodo_oficial":p.isoformat(),"no_completa_periodos_futuros":True}}
def localizar_fila(ws,anio,mes):
 for f in range(2,ws.max_row+1):
  v=ws.cell(f,1).value
  if isinstance(v,(date,datetime)) and v.year==anio and v.month==mes:return f
 raise ValueError(f"No existe el período {anio}-{mes:02d} en columna A")
def resolver_destinos_sipa(ws):
 """Resuelve títulos existentes y completa solo encabezados SIPA vacíos."""
 columnas={}
 claves_sectoriales=set()
 for indice in range(14):
  base=16+indice*3
  claves_sectoriales.update({f"S{base:02d}",f"S{base+1:02d}",f"S{base+2:02d}"})
 for clave,especificacion in DESTINOS.items():
  if clave in claves_sectoriales: continue
  columnas.update(resolver_columnas(ws,{clave:especificacion}))
 for indice in range(14):
  numero=16+indice*3; clave_base=f"S{numero:02d}"
  columna_base=resolver_columnas(ws,{clave_base:DESTINOS[clave_base]})[clave_base]
  columnas[clave_base]=columna_base
  for desplazamiento in (1,2):
   clave=f"S{numero+desplazamiento:02d}"; columna=columna_base+desplazamiento
   titulo_actual=ws.cell(1,columna).value
   if titulo_actual in (None,""):
    ws.cell(1,columna).value=TITULOS[numero+desplazamiento-1]
   columnas[clave]=columna
 if len(columnas)!=70 or len(set(columnas.values()))!=70:
  raise ValueError("No se pudieron identificar de forma única las 70 columnas SIPA")
 return columnas
def actualizar_base(r,ruta):
 wb=load_workbook(ruta,data_only=False,keep_links=True); ws=wb[HOJA_BASE]; cols=resolver_destinos_sipa(ws); f=localizar_fila(ws,r["anio"],r["mes_numero"]); cambios=[]; iguales=[]
 for k,nuevo in r["valores"].items():
  c=ws.cell(f,cols[k]); antes=c.value
  try: igual=math.isclose(float(str(antes).replace(",","")),float(nuevo),rel_tol=0,abs_tol=1e-7)
  except (TypeError,ValueError): igual=antes==nuevo
  if igual: iguales.append(c.coordinate)
  else: c.value=nuevo; cambios.append({"celda":c.coordinate,"titulo":TITULOS[int(k[1:])-1],"antes":antes,"despues":nuevo})
 # Elimina restos de fórmulas o valores en meses todavía no publicados.
 periodo_oficial=date(r["anio"],r["mes_numero"],1)
 for fila in range(2,ws.max_row+1):
  periodo=periodo_celda(ws.cell(fila,1).value)
  if not periodo or periodo <= periodo_oficial: continue
  for clave,columna in cols.items():
   celda=ws.cell(fila,columna)
   if celda.value not in (None,""):
    antes=celda.value; celda.value=None
    cambios.append({"celda":celda.coordinate,"titulo":TITULOS[int(clave[1:])-1],"antes":antes,"despues":None,"motivo":"período SIPA aún no publicado"})
 if cambios: wb.save(ruta)
 return {"archivo_excel":str(ruta.resolve()),"fila":f,"celdas_actualizadas":len(cambios),"cambios":cambios,"coincidencias":iguales,"periodo_oficial":r["periodo"]}
def main():
 raiz=raiz_proyecto(); ap=argparse.ArgumentParser(); ap.add_argument("--base",type=Path,default=raiz/"base"/ARCHIVO_BASE); ap.add_argument("--solo-extraer",action="store_true"); ap.add_argument("--xlsx",type=Path); a=ap.parse_args()
 try:
  if a.xlsx: contenido=a.xlsx.read_bytes(); url=a.xlsx.resolve().as_uri()
  else: url,contenido=descubrir_planilla()
  r=extraer(contenido,url)
  if not a.solo_extraer:r["actualizacion"]=actualizar_base(r,a.base)
  carpeta=raiz/"scripts"/"reportes"/"sipa"; carpeta.mkdir(parents=True,exist_ok=True); rep=carpeta/f"sipa_{r['anio']}_{r['mes_numero']:02d}.json"; rep.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding="utf-8"); r["reporte"]=str(rep.resolve())
  print(json.dumps(r,ensure_ascii=False,indent=2)); print("ESTADO=OK"); return 0
 except Exception as e: print(f"ERROR={e}"); print("ESTADO=ERROR"); return 1
if __name__=="__main__": raise SystemExit(main())
