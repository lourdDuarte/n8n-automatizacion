"""Extrae el nivel general y las 16 divisiones del IPI (55 valores)."""
from __future__ import annotations
from pathlib import Path
import json, re, sys, unicodedata
import fitz
try:
    import xlrd
except ImportError as error:
    raise SystemExit("ERROR: falta xlrd. Instalar una vez con: python.exe -m pip install xlrd") from error

RAIZ = Path(__file__).resolve().parent.parent
CARPETA_FUENTES = RAIZ / "fuentes" / "ipi"
CARPETA_REPORTES = RAIZ / "reportes" / "ipi"
MESES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
         "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
DIVISIONES = (
    ("alimentos_bebidas", "Alimentos y bebidas"),
    ("productos_tabaco", "Productos de tabaco"),
    ("productos_textiles", "Productos textiles"),
    ("prendas_cuero_calzado", "Prendas de vestir, cuero y calzado"),
    ("madera_papel_edicion", "Madera, papel, edición e impresión"),
    ("refinacion_petroleo", "Refinación del petróleo, coque y combustible nuclear"),
    ("sustancias_quimicos", "Sustancias y productos químicos"),
    ("caucho_plastico", "Productos de caucho y plástico"),
    ("minerales_no_metalicos", "Productos minerales no metálicos"),
    ("metalicas_basicas", "Industrias metálicas básicas"),
    ("productos_metal", "Productos de metal"),
    ("maquinaria_equipo", "Maquinaria y equipo"),
    ("otros_equipos", "Otros equipos, aparatos e instrumentos"),
    ("vehiculos_automotores", "Vehículos automotores, carrocerías, remolques y autopartes"),
    ("otro_transporte", "Otro equipo de transporte"),
    ("muebles_otras", "Muebles y colchones, y otras industrias manufactureras"),
)

def configurar_salida():
    try:
        sys.stdout.reconfigure(encoding="utf-8"); sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError): pass

def normalizar(valor):
    texto=" ".join(str(valor or "").replace("\u00a0"," ").split()).casefold()
    return "".join(c for c in unicodedata.normalize("NFD",texto) if unicodedata.category(c)!="Mn")

def convertir_numero(texto):
    return float(texto.strip().replace("−","-").replace(".","").replace(",","."))

def buscar_pdf():
    archivos=list(CARPETA_FUENTES.glob("ipi_manufacturero_????_??.pdf"))
    if not archivos: raise FileNotFoundError(f"No se encontraron informes en {CARPETA_FUENTES}")
    return max(archivos,key=lambda p:p.stem[-7:])

def leer_paginas(ruta):
    with fitz.open(ruta) as doc:
        if not len(doc): raise ValueError("El PDF no contiene páginas.")
        return [p.get_text("text",sort=True) for p in doc]

def detectar_periodo(paginas):
    patron=re.compile(rf"\b({'|'.join(MESES)})\s+de\s+(20\d{{2}})\b",re.I)
    halladas=patron.findall("\n".join(paginas[:5]))
    if not halladas: raise ValueError("No se pudo detectar el período del informe.")
    mes,anio=halladas[-1]
    return f"{int(anio):04d}-{MESES[mes.lower()]:02d}",mes.capitalize(),int(anio)

def pagina_cuadro_1(paginas):
    for numero,texto in enumerate(paginas,1):
        limpio=texto.casefold()
        if all(x in limpio for x in ("cuadro 1.","ipi manufacturero nivel general","índice serie original","período","números índice")):
            return numero,texto
    raise ValueError("No se encontró el Cuadro 1 del nivel general.")

def extraer_fila_pdf(texto,mes,anio):
    lineas=[" ".join(x.replace("\u00a0"," ").split()) for x in texto.splitlines() if x.strip()]
    numero=r"[+-]?\d+(?:[.,]\d+)?"
    patron=re.compile(rf"^(?:(20\d{{2}})\s*\*?\s+)?({'|'.join(MESES)})\s+"+r"\s+".join([f"({numero})"]*7)+r"\s*$",re.I)
    anio_actual=None
    for linea in lineas:
        m=patron.fullmatch(linea)
        if not m: continue
        if m.group(1): anio_actual=int(m.group(1))
        if anio_actual==anio and m.group(2).casefold()==mes.casefold():
            return [convertir_numero(x) for x in m.groups()[2:]]
    raise ValueError(f"No se encontró la fila completa de {mes} de {anio} en el Cuadro 1.")

def buscar_xls(periodo):
    ruta=CARPETA_FUENTES/f"sh_ipi_manufacturero_{periodo.replace('-','_')}.xls"
    if not ruta.is_file(): raise FileNotFoundError(f"No se encontró la planilla detallada: {ruta}")
    return ruta

def fila_periodo(hoja,anio,mes):
    anio_actual=None
    for fila in range(5,hoja.nrows):
        valor=normalizar(hoja.cell_value(fila,1)).replace("*","")
        if valor:
            try: anio_actual=int(float(valor))
            except ValueError: pass
        if anio_actual==anio and MESES.get(normalizar(hoja.cell_value(fila,2)))==mes: return fila
    raise ValueError(f"No se encontró {anio}-{mes:02d} en {hoja.name!r}.")

def columna_titulo(hoja,titulo):
    halladas=[c for c in range(hoja.ncols) if normalizar(hoja.cell_value(3,c))==normalizar(titulo)]
    if not halladas: raise ValueError(f"Columna {titulo!r} ausente en {hoja.name!r}.")
    # En Cuadros 3 y 4 "Otro equipo de transporte" aparece también como
    # subclase. La primera coincidencia es siempre la división agregada.
    return halladas[0]

def numero_xls(hoja,fila,columna,descripcion):
    valor=hoja.cell_value(fila,columna)
    if isinstance(valor,bool) or not isinstance(valor,(int,float)):
        raise ValueError(f"Dato no numérico para {descripcion}: {valor!r}")
    if not -1000<float(valor)<1000: raise ValueError(f"Dato fuera de rango para {descripcion}: {valor}")
    return round(float(valor),1)

def extraer_divisiones(ruta,anio,mes):
    libro=xlrd.open_workbook(str(ruta),on_demand=True); datos={}
    for hoja_nombre,prefijo,desplazamiento in (("Cuadro 3","ia",0),("Cuadro 4","acum",0),("Cuadro 5","im",1)):
        hoja=libro.sheet_by_name(hoja_nombre); fila=fila_periodo(hoja,anio,mes)
        for clave,titulo in DIVISIONES:
            datos[f"{prefijo}_{clave}"]=numero_xls(hoja,fila,columna_titulo(hoja,titulo)+desplazamiento,f"{hoja_nombre} / {titulo}")
    libro.release_resources()
    if len(datos)!=48: raise RuntimeError(f"Se esperaban 48 valores de divisiones y se obtuvieron {len(datos)}.")
    return datos

def extraer_ipi(ruta_pdf):
    paginas=leer_paginas(ruta_pdf); periodo,mes,anio=detectar_periodo(paginas)
    pagina,texto=pagina_cuadro_1(paginas); n=extraer_fila_pdf(texto,mes,anio)
    datos={"serie_original":n[0],"variacion_interanual":n[1],"variacion_acumulada":n[2],
           "serie_desestacionalizada":n[3],"variacion_mensual_desestacionalizada":n[4],
           "serie_tendencia_ciclo":n[5],"variacion_mensual_tendencia_ciclo":n[6]}
    xls=buscar_xls(periodo); datos.update(extraer_divisiones(xls,anio,MESES[mes.casefold()]))
    if len(datos)!=55: raise RuntimeError(f"Se esperaban 55 valores IPI y se obtuvieron {len(datos)}.")
    return {"archivo_pdf":ruta_pdf.name,"archivo_xls":xls.name,"periodo":periodo,"mes":mes,"anio":anio,
            "pagina_fuente":pagina,"cantidad_datos":len(datos),"datos":datos}

def main():
    configurar_salida(); resultado=extraer_ipi(buscar_pdf()); CARPETA_REPORTES.mkdir(parents=True,exist_ok=True)
    salida=CARPETA_REPORTES/f"ipi_{resultado['periodo'].replace('-','_')}.json"
    salida.write_text(json.dumps(resultado,ensure_ascii=False,indent=2),encoding="utf-8")
    print("="*70); print("EXTRACCIÓN — IPI MANUFACTURERO"); print("="*70)
    print(f"Período: {resultado['periodo']}")
    print(f"Valores validados: {resultado['cantidad_datos']} (7 generales + 48 divisiones)")
    print(f"IPI_JSON={salida.resolve()}"); print("RESULTADO: OK"); return 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except Exception as error:
        configurar_salida(); print(f"ERROR: {type(error).__name__}: {error}",file=sys.stderr); raise SystemExit(1)
