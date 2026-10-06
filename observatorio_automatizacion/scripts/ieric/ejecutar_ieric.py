from __future__ import annotations
import argparse
import io
import json
import re
import sys
import urllib.parse
import urllib.request
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from openpyxl import load_workbook
SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
from comun.columnas_por_titulo import resolver_columnas
PAGINA = 'https://www.ieric.org.ar/estadistica/informes-de-coyuntura/'
HOJA = 'BASE DE DATOS'
MESES = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12}
DESTINOS = {'empresas_constructoras.formosa': 'Empresas constructoras/contratistas F', 'empresas_constructoras.nacion': 'Empresas constructoras/contratistas N', 'puestos_trabajo_registrados.formosa': 'Puestos de trabajo registrados en el sector construccion F', 'puestos_trabajo_registrados.nacion': 'Puestos de trabajo registrados en el sector construccion N', 'salario_promedio_construccion.formosa': 'Salario promedio del sector Construcción (en pesos- con estacionalidad) F', 'salario_promedio_construccion.nacion': 'Salario promedio del sector Construcción (en pesos- con estacionalidad) N'}
DESTINOS.update({'puestos_trabajo_registrados.var_mensual_formosa': 'Var. IM Puestos de trabajo registrados en el sector construccion F', 'puestos_trabajo_registrados.var_mensual_nacion': 'Var. IM Puestos de trabajo registrados en el sector construccion N', 'puestos_trabajo_registrados.var_interanual_formosa': 'Var. IA  Puestos de trabajo registrados en el sector construccion F', 'puestos_trabajo_registrados.var_interanual_nacion': 'Var. IA  Puestos de trabajo registrados en el sector construccion N', 'salario_promedio_construccion.var_mensual_formosa': 'Var. IM salario nominal del sector construccion (C/E) F', 'salario_promedio_construccion.var_mensual_nacion': 'Var. IM salario nominal del sector construccion (C/E) N', 'salario_promedio_construccion.var_interanual_formosa': 'Var. IA salario nominal del sector construccion (C/E) F', 'salario_promedio_construccion.var_interanual_nacion': 'Var. IA salario nominal del sector construccion (C/E) N'})
# Cuadro VIII del informe: variación interanual del salario real (Formosa y Total país).
DESTINOS.update({'salario_real.var_interanual_formosa': 'Var. IA salario real sector construcción (C/E)F', 'salario_real.var_interanual_nacion': 'Var. IA salario real sector construcción (C/E) N'})
# Salario SIN estacionalidad: igual al salario con estacionalidad, salvo junio y diciembre
# (IERIC: "El salario correspondiente a Diciembre y Junio incluye el medio aguinaldo"),
# donde se quita el aguinaldo con el promedio del mes anterior y el siguiente.
SALARIO_CE = {'F': 'Salario promedio del sector Construcción (en pesos- con estacionalidad) F', 'N': 'Salario promedio del sector Construcción (en pesos- con estacionalidad) N'}
SALARIO_SIN = {'F': 'Salario promedio del sector Construcción (en pesos- SIN estacionalidad) F', 'N': 'Salario promedio del sector Construcción (en pesos- SIN estacionalidad) N'}
MESES_AGUINALDO = (6, 12)

def raiz_proyecto() -> Path:
    return Path(__file__).resolve().parents[2]

class LinkParser(HTMLParser):

    def __init__(self) -> None:
        super().__init__()
        self.links: list[tuple[str, str]] = []
        self._href: str | None = None
        self._texto: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() == 'a':
            self._href = dict(attrs).get('href')
            self._texto = []

    def handle_data(self, data: str) -> None:
        if self._href is not None:
            self._texto.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == 'a' and self._href is not None:
            self.links.append((' '.join(''.join(self._texto).split()), self._href))
            self._href = None

def pedir(url: str, timeout: int=60) -> bytes:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (automatizacion-indicadores/1.0)'})
    with urllib.request.urlopen(req, timeout=timeout) as respuesta:
        return respuesta.read()

def descubrir_informe(html: str) -> str:
    parser = LinkParser()
    parser.feed(html)
    candidatos = []
    for texto, href in parser.links:
        absoluto = urllib.parse.urljoin(PAGINA, href)
        if absoluto.lower().endswith('.pdf') and ('informe completo' in texto.lower() or 'informe-coyuntura-construccion' in absoluto.lower()):
            numero = re.search('(?:no-|n[oº°]?[-_ ]?)(\\d{3})', absoluto, re.I)
            candidatos.append((int(numero.group(1)) if numero else 0, absoluto))
    if not candidatos:
        raise ValueError('No se encontró el enlace al Informe completo')
    return max(candidatos)[1]

def texto_pdf(contenido: bytes) -> str:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError('Falta pypdf. Instalalo una sola vez con: python -m pip install pypdf') from exc
    lector = PdfReader(io.BytesIO(contenido))
    return '\n'.join((p.extract_text() or '' for p in lector.pages))

def numero_argentino(valor: str) -> int | float:
    limpio = valor.replace(' ', '').replace('.', '').replace(',', '.')
    numero = float(limpio)
    return int(numero) if numero.is_integer() else numero

def periodo(texto):
    patron = '|'.join(MESES)
    m = re.search(f'\\b({patron})\\s+(?:de\\s+)?(20\\d{{2}})\\b', texto, re.I)
    if not m:
        raise ValueError('No se identifica el período del cuadro')
    nombre = m[1].lower()
    return {'mes': nombre.capitalize(), 'mes_numero': MESES[nombre], 'anio': int(m[2])}

def fila_por_nombre(lineas, nombre):
    coincidencias = [x for x in lineas if re.match(f'^{re.escape(nombre)}\\b', x, re.I)]
    if len(coincidencias) != 1:
        raise ValueError(f'Fila inexistente o repetida: {nombre}')
    return coincidencias[0]

def numeros(linea):
    return [numero_argentino(x) for x in re.findall('-?\\d[\\d.]*?(?:,\\d+)?(?=\\s|%|$)', linea)]

def bloque(lineas, titulo):
    indices = [i for i, x in enumerate(lineas) if titulo.casefold() in x.casefold()]
    if len(indices) != 1:
        raise ValueError(f'Cuadro inexistente o repetido: {titulo}')
    inicio = indices[0]
    fin = next((i for i in range(inicio + 1, len(lineas)) if lineas[i].startswith('Informe de Coyuntura')), len(lineas))
    return lineas[inicio:fin]

def salario_real(lineas, nombre, salario):
    """Fila del Cuadro VIII: salario, IPC i.a., IPC acum., salario real i.a., salario real acum.

    Se identifica por tener 5 números y el mismo salario que el Cuadro VI (control cruzado)."""
    candidatas = []
    for linea in lineas:
        if re.match(f'^{re.escape(nombre)}\\b', linea, re.I):
            valores = numeros(linea)
            if len(valores) == 5 and abs(float(valores[0]) - float(salario)) < 0.06:
                candidatas.append(valores)
    if len(candidatas) != 1:
        raise ValueError(f'Salario real de {nombre}: {len(candidatas)} filas coincidentes en el Cuadro VIII')
    return candidatas[0][3]

def extraer_indicadores(texto):
    texto = texto.replace('\u2212', '-')
    lineas = [re.sub('(?<=\\.)\\s+(?=\\d)', '', ' '.join(x.split())).replace('de trabajo de trabajo', 'de trabajo') for x in texto.splitlines()]
    empresas = bloque(lineas, 'Empresas por Provincia según Tipo de actividad')
    abreviados = {nombre[:3]: numero for nombre, numero in MESES.items()}
    encabezados = []
    for linea in empresas:
        tokens = re.findall('\\b(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[- /](\\d{4}|\\d{2})\\b', linea, re.I)
        if len(tokens) >= 6:
            encabezados.append([(2000 + int(a) if len(a) == 2 else int(a), abreviados[m.lower()]) for m, a in tokens])
    if len(encabezados) != 1:
        raise ValueError('No se identifica un encabezado mensual único para empresas')
    fechas = encabezados[0]
    if fechas[:2] != fechas[2:4] or fechas[:2] != fechas[4:6] or fechas[0] == fechas[1]:
        raise ValueError('Cambió la estructura de columnas de empresas IERIC')
    columna = max(range(2), key=lambda i: fechas[i])
    anio, mes = fechas[columna]
    nombre = next((n for n, v in MESES.items() if v == mes))
    valores_f = numeros(fila_por_nombre(empresas, 'Formosa'))
    valores_n = numeros(fila_por_nombre(empresas, 'Total País'))
    for valores in (valores_f, valores_n):
        if len(valores) < 6 or valores[0] + valores[2] != valores[4] or valores[1] + valores[3] != valores[5]:
            raise ValueError('Las columnas de empresas no coinciden con sus totales')
    laboral = bloque(lineas, 'Puestos de trabajo registrados y Salario promedio por Provincia')
    fecha_laboral = periodo(' '.join(laboral[:3]))
    formosa = fila_por_nombre(laboral, 'Formosa')
    desde_formosa = laboral[laboral.index(formosa) + 1:]
    total = fila_por_nombre(desde_formosa, 'Total')
    trabajo_f, trabajo_n = (numeros(formosa), numeros(total))
    if min(len(trabajo_f), len(trabajo_n)) != 8:
        raise ValueError('Faltan los valores de empleo y salarios')
    real = {**fecha_laboral, 'var_interanual_formosa': salario_real(lineas, 'Formosa', trabajo_f[1]), 'var_interanual_nacion': salario_real(lineas, 'Total', trabajo_n[1])}
    return {'salario_real': real, 'empresas_constructoras': {'mes': nombre.capitalize(), 'mes_numero': mes, 'anio': anio, 'formosa': valores_f[columna], 'nacion': valores_n[columna]}, 'puestos_trabajo_registrados': {**fecha_laboral, 'formosa': trabajo_f[0], 'nacion': trabajo_n[0], 'var_mensual_formosa': trabajo_f[3], 'var_mensual_nacion': trabajo_n[3], 'var_interanual_formosa': trabajo_f[4], 'var_interanual_nacion': trabajo_n[4]}, 'salario_promedio_construccion': {**fecha_laboral, 'formosa': trabajo_f[1], 'nacion': trabajo_n[1], 'var_mensual_formosa': trabajo_f[5], 'var_mensual_nacion': trabajo_n[5], 'var_interanual_formosa': trabajo_f[6], 'var_interanual_nacion': trabajo_n[6]}}

def extraer() -> dict:
    pagina = pedir(PAGINA).decode('utf-8', errors='replace')
    url_pdf = descubrir_informe(pagina)
    contenido = pedir(url_pdf)
    if not contenido.startswith(b'%PDF'):
        raise ValueError('El enlace al informe completo no devolvió un PDF')
    texto = texto_pdf(contenido)
    numero = re.search('No-(\\d{3})', url_pdf, re.I)
    series = extraer_indicadores(texto)
    periodos_fuente = sorted({f"{datos['anio']:04d}-{datos['mes_numero']:02d}-01" for datos in series.values()})
    laboral = series['puestos_trabajo_registrados']
    periodo_destino = f"{laboral['anio']:04d}-{laboral['mes_numero']:02d}-01"
    return {'fuente': 'IERIC', 'version_extractor': 'IERIC_BLOQUE_20260930', 'consultado_en': datetime.now().isoformat(timespec='seconds'), 'informe_numero': int(numero.group(1)) if numero else None, 'url_informe': url_pdf, 'series': series, 'periodo': periodo_destino, 'periodos': periodos_fuente, 'periodos_fuente': periodos_fuente, 'nota': 'Empresas en su mes oficial; empleo y salarios en su mes laboral. Las series conservan sus períodos oficiales de origen.'}

def localizar_fila(ws, anio: int, mes: int) -> int:
    for fila in range(2, ws.max_row + 1):
        valor = ws.cell(fila, 1).value
        if isinstance(valor, datetime) and valor.year == anio and (valor.month == mes):
            return fila
    raise ValueError(f'No existe el período {anio}-{mes:02d} en la columna A')

def _mes_relativo(anio: int, mes: int, delta: int) -> tuple[int, int]:
    total = anio * 12 + (mes - 1) + delta
    return total // 12, total % 12 + 1


def aplicar_sin_estacionalidad(ws, anio: int, mes: int, cargas: dict, cambios: list, detalle: list) -> None:
    """Completa las columnas SIN estacionalidad del mes laboral cargado.

    * Mes común: igual al salario con estacionalidad.
    * Junio o diciembre (con medio aguinaldo): queda pendiente hasta que se publique el mes siguiente.
    * Julio o enero: además completa junio/diciembre con la fórmula que quita el aguinaldo,
      (mes anterior + mes siguiente) / 2.
    """
    from openpyxl.utils import get_column_letter
    columnas = resolver_columnas(ws, {**{f'ce_{k}': v for k, v in SALARIO_CE.items()}, **{f'sin_{k}': v for k, v in SALARIO_SIN.items()}})

    def poner(fila, col, valor, periodo, reemplazar=False):
        celda = ws.cell(fila, col)
        titulo = str(ws.cell(1, col).value)
        fecha = f'{periodo[0]:04d}-{periodo[1]:02d}-01'
        carga = cargas.setdefault(fecha, {'hoja': HOJA, 'row_number': fila, 'periodo': fecha, 'valores': {}})
        carga['valores'][titulo] = valor
        if reemplazar:
            carga['permitir_reemplazos'] = True
        registro = {'celda': celda.coordinate, 'titulo': titulo, 'antes': celda.value, 'despues': valor}
        detalle.append(registro)
        if celda.value != valor:
            celda.value = valor
            cambios.append(registro)

    fila = localizar_fila(ws, anio, mes)
    if mes not in MESES_AGUINALDO:
        for k in ('F', 'N'):
            poner(fila, columnas[f'sin_{k}'], ws.cell(fila, columnas[f'ce_{k}']).value, (anio, mes))
    anterior = _mes_relativo(anio, mes, -1)
    if anterior[1] in MESES_AGUINALDO:
        fila_ant = localizar_fila(ws, *anterior)
        fila_ant2 = localizar_fila(ws, *_mes_relativo(anio, mes, -2))
        for k in ('F', 'N'):
            letra_ce = get_column_letter(columnas[f'ce_{k}'])
            formula = f'=({letra_ce}{fila_ant2}+{letra_ce}{fila})/2'
            # Si junio/diciembre tenía pegado el salario con aguinaldo, se reemplaza por la fórmula.
            poner(fila_ant, columnas[f'sin_{k}'], formula, anterior, reemplazar=True)


def actualizar_base(resultado: dict, ruta_base: Path) -> dict:
    if not ruta_base.is_file():
        raise FileNotFoundError(f'No existe la base: {ruta_base}')
    libro = load_workbook(ruta_base, data_only=False)
    ws = libro[HOJA]
    columnas = resolver_columnas(ws, DESTINOS)
    laboral = resultado['series']['puestos_trabajo_registrados']
    filas = {}
    periodo_destino = resultado['periodo']
    cambios, coincidencias, detalle = ([], [], [])
    cargas = {}
    for clave, titulo in DESTINOS.items():
        serie, campo = clave.split('.', 1)
        valor = resultado['series'][serie][campo]
        datos_serie = resultado['series'][serie]
        fila = localizar_fila(ws, datos_serie['anio'], datos_serie['mes_numero'])
        filas[serie] = fila
        fecha = f"{datos_serie['anio']:04d}-{datos_serie['mes_numero']:02d}-01"
        carga = cargas.setdefault(fecha, {'hoja': HOJA, 'row_number': fila, 'periodo': fecha, 'valores': {}})
        celda = ws.cell(fila, columnas[clave])
        titulo_real = str(ws.cell(1, columnas[clave]).value)
        anterior = celda.value
        if isinstance(anterior, str) and anterior.startswith('='):
            raise ValueError(f'Dato original con fórmula: {celda.coordinate}; no se reemplazó')
        carga['valores'][titulo_real] = valor
        registro = {'celda': celda.coordinate, 'titulo': titulo_real, 'antes': anterior, 'despues': valor}
        detalle.append(registro)
        if anterior == valor:
            coincidencias.append(celda.coordinate)
        else:
            celda.value = valor
            cambios.append(registro)
    aplicar_sin_estacionalidad(ws, laboral['anio'], laboral['mes_numero'], cargas, cambios, detalle)
    respaldo = None
    if cambios:
        import os
        import tempfile
        descriptor, nombre = tempfile.mkstemp(prefix='.ieric_', suffix='.xlsx', dir=ruta_base.parent)
        os.close(descriptor)
        try:
            libro.save(nombre)
            libro.close()
            os.replace(nombre, ruta_base)
        finally:
            if os.path.exists(nombre):
                os.unlink(nombre)
    return {'archivo_excel': str(ruta_base.resolve()), 'filas': filas, 'periodo': periodo_destino, 'celdas_actualizadas': len(cambios), 'hubo_cambios': bool(cambios), 'cambios': cambios, 'detalle_completo': detalle, 'coincidencias': coincidencias, 'respaldo': str(respaldo) if respaldo else None, 'cargas_google': list(cargas.values())}

def aplicar_correccion(ruta_json: Path, ruta_base: Path) -> dict:
    """Corrección única y controlada (no se ejecuta en las corridas diarias).

    Aplica los valores verificados contra los informes oficiales que figuran en el
    archivo JSON, y las fórmulas SIN estacionalidad de los meses con aguinaldo indicados.
    Puede reemplazar datos ya cargados (por eso es una fuente manual aparte).
    """
    import shutil
    datos = json.loads(ruta_json.read_text(encoding='utf-8'))
    libro = load_workbook(ruta_base, data_only=False)
    ws = libro[HOJA]
    titulos = sorted({item['titulo'] for item in datos.get('valores', [])})
    columnas = resolver_columnas(ws, {t: t for t in titulos})
    cargas, cambios, detalle, omitidas = {}, [], [], []
    for item in datos.get('valores', []):
        anio, mes = (int(x) for x in item['periodo'].split('-')[:2])
        fila = localizar_fila(ws, anio, mes)
        celda = ws.cell(fila, columnas[item['titulo']])
        titulo_real = str(ws.cell(1, columnas[item['titulo']]).value)
        if isinstance(celda.value, str) and celda.value.startswith('='):
            omitidas.append(f'{celda.coordinate} (fórmula en el Excel local)')
            continue
        fecha = f'{anio:04d}-{mes:02d}-01'
        carga = cargas.setdefault(fecha, {'hoja': HOJA, 'row_number': fila, 'periodo': fecha, 'valores': {}, 'permitir_reemplazos': True})
        carga['valores'][titulo_real] = item['valor']
        registro = {'celda': celda.coordinate, 'titulo': titulo_real, 'antes': celda.value, 'despues': item['valor'], 'informe': item.get('informe')}
        detalle.append(registro)
        if celda.value != item['valor']:
            celda.value = item['valor']
            cambios.append(registro)
    for periodo in datos.get('sin_estacionalidad_aguinaldo', []):
        anio, mes = (int(x) for x in periodo.split('-')[:2])
        if mes not in MESES_AGUINALDO:
            raise ValueError(f'{periodo} no es un mes con aguinaldo')
        aplicar_sin_estacionalidad(ws, *_mes_relativo(anio, mes, 1), cargas, cambios, detalle)
    respaldo = None
    if cambios:
        carpeta = ruta_base.parent / 'respaldos'
        carpeta.mkdir(parents=True, exist_ok=True)
        respaldo = carpeta / f'antes_correccion_ieric_{datetime.now():%Y%m%d_%H%M%S}.xlsx'
        shutil.copy2(ruta_base, respaldo)
        libro.save(ruta_base)
    libro.close()
    periodos = sorted(cargas)
    return {'fuente': 'IERIC', 'modo': 'correccion_unica', 'archivo_correccion': str(ruta_json),
            'periodo': periodos[-1] if periodos else None, 'periodos': periodos,
            'actualizacion': {'archivo_excel': str(ruta_base.resolve()), 'celdas_actualizadas': len(cambios),
                              'hubo_cambios': bool(cambios), 'cambios': cambios, 'detalle_completo': detalle,
                              'omitidas': omitidas, 'respaldo': str(respaldo) if respaldo else None,
                              'cargas_google': [cargas[k] for k in periodos]}}


def prueba_parser() -> None:
    muestra = '<a href="/wp-content/uploads/2026/06/IERIC-Informe-Coyuntura-Construccion-No-248.pdf">Informe completo</a><a href="/wp-content/uploads/2026/07/IERIC-Informe-Coyuntura-Construccion-No-249.pdf">Informe completo</a>'
    assert descubrir_informe(muestra).endswith('No-249.pdf')
    print('PRUEBA_PARSER=OK')

def main() -> int:
    parser = argparse.ArgumentParser(description='Detecta el último informe IERIC y actualiza la base.')
    parser.add_argument('--probar-parser', action='store_true')
    parser.add_argument('--solo-extraer', action='store_true')
    parser.add_argument('--aplicar-correccion', type=Path, help='Archivo JSON de corrección única (uso manual).')
    parser.add_argument('--base', type=Path, default=raiz_proyecto() / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx')
    args = parser.parse_args()
    if args.probar_parser:
        prueba_parser()
        return 0
    try:
        if args.aplicar_correccion:
            ruta = args.aplicar_correccion
            if not ruta.is_absolute():
                ruta = Path(__file__).resolve().parent / ruta.name
            resultado = aplicar_correccion(ruta, args.base)
            print(json.dumps(resultado, ensure_ascii=False, indent=2, default=str))
            print('ESTADO=OK')
            return 0
        resultado = extraer()
        if not args.solo_extraer:
            resultado['actualizacion'] = actualizar_base(resultado, args.base)
        print(json.dumps(resultado, ensure_ascii=False, indent=2))
        print('ESTADO=OK')
        return 0
    except Exception as error:
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        return 1
if __name__ == '__main__':
    sys.exit(main())
