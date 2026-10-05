"""SRT - Boletín mensual por jurisdicción (Formosa).

Detecta los meses publicados por la SRT que todavía no están en la base,
descarga cada boletín PDF, toma el "Cuadro 1.1" (personas trabajadoras y
parte empleadora por sector de actividad) y carga:

* SRT Evolución Cantidad de Trabajadores  (Total personas trabajadoras)
* SRT Evolución Cantidad de Empleadores   (Total parte empleadora)
* las 21 columnas por sector (personas trabajadoras)

Las variaciones intermensuales, diferencias e interanuales se completan con
la misma fórmula del mes anterior. Antes de escribir se controla que la suma
de los sectores coincida con el total publicado.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
from bloque_mensual import cargar_meses, texto_periodo, ultimo_periodo_cargado  # noqa: E402
from columnas_por_titulo import normalizar_titulo  # noqa: E402

BASE = RAIZ / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx'
SITIO = 'https://www.srt.gob.ar/estadisticas/'
BUSCADOR = SITIO + 'buscadores/buscador_cyf_boletin_provincia.php'
PROVINCIA = '18_for|Formosa'
CARPETA = RAIZ / 'fuentes' / 'srt'
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
         'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
TOTAL_TRABAJADORES = 'SRT Evolución Cantidad de Trabajadores'
TOTAL_EMPLEADORES = 'SRT Evolución Cantidad de Empleadores '
# Sector del boletín -> encabezado de la base.
SECTORES = {
    'Agricultura, ganadería, caza, silvicultura y pesca': 'SRT Agricultura, ganadería, caza, silvicultura y pesca F',
    'Explotación de minas y canteras': 'SRT Explotación de minas y canteras F',
    'Industria manufacturera': 'SRT Industria manufacturera F',
    'Suministro de electricidad, gas, vapor y aire acondicionado': 'SRT Suministro de electricidad, gas, vapor y aire acondicionado F',
    'Suministro de agua, cloacas, gestión de residuos y recuperación de materiales y saneamiento público':
        'SRT Suministro de agua, cloacas, gestión de residuos y recuperación de materiales y saneamiento público F',
    'Construcción': 'SRT Construcción',
    'Comercio al por mayor y al por menor; reparación de vehículos automotores y motocicletas':
        'SRT Comercio al por mayor y al por menor; reparación de vehículos automotores y motocicletas F',
    'Servicio de transporte y almacenamiento': 'SRT Servicio de transporte y almacenamiento F',
    'Servicios de alojamiento y servicios de comida': 'SRT Servicios de alojamiento y servicios de comida F',
    'Información y comunicaciones': 'SRT Información y comunicaciones F',
    'Intermediación financiera y servicios de seguros': 'SRT Intermediación financiera y servicios de seguros F',
    'Servicios inmobiliarios': 'SRT Servicios inmobiliarios F',
    'Servicios profesionales, científicos y técnicos': 'SRT Servicios profesionales, científicos y técnicos F',
    'Actividades administrativas y servicios de apoyo': 'SRT Actividades administrativas y servicios de apoyo F',
    'Administración pública, defensa y seguridad social obligatoria': 'SRT Administración pública, defensa y seguridad social obligatoria F',
    'Enseñanza': 'SRT Enseñanza',
    'Salud humana y servicios sociales': 'SRT Salud humana y servicios sociales F',
    'Servicios artísticos, culturales, deportivos y de esparcimiento': 'SRT Servicios artísticos, culturales, deportivos y de esparcimiento F',
    'Servicios de asociaciones y servicios personales': 'SRT Servicios de asociaciones y servicios personales F',
    'Servicios de organizaciones y órganos extraterritoriales': 'SRT Servicios de organizaciones y órganos extraterritoriales F',
}
CALCULADAS = [
    'SRT Var % intermensual asalariados registrados del sector privado F',
    'SRT Dif intermensual asalariados registrados del sector privado F',
    'SRT Var % interanual asalariados registrados del sector privado F',
    'SRT Variacion % intermensual empleadores registrados del sector privado F',
    'SRT Diferencia intermensual empleadores registrados del sector privado F',
    'SRT Variacion % interanual empleadores F',
]
NUM = r'(\d{1,3}(?:\.\d{3})*|\d+)'
PCT = r'(?:-?\d+,\d+%|-)'


def _plano(t: str) -> str:
    t = unicodedata.normalize('NFKD', t)
    t = ''.join(c for c in t if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', t).strip().casefold()


def _http(url: str, datos: dict | None = None, timeout: int = 60) -> bytes:
    cuerpo = urllib.parse.urlencode(datos).encode() if datos else None
    req = urllib.request.Request(url, data=cuerpo, headers={'User-Agent': 'Mozilla/5.0 (observatorio-formosa)'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def url_boletin(anio: int, mes: int) -> str | None:
    """Pregunta al buscador de la SRT si el boletín del mes está publicado."""
    html = _http(BUSCADOR, {'provincia': PROVINCIA, 'anio': str(anio), 'mes': f'{mes:02d}|{MESES[mes - 1]}'}).decode('utf-8', 'replace')
    m = re.search(r'href="([^"]+\.pdf)"', html)
    return urllib.parse.urljoin(SITIO, m.group(1)) if m else None


TOTAL = re.compile(rf'Total\s+{NUM}\s+{PCT}\s+{NUM}\s+{PCT}\s+\$')


def texto_pdf(ruta: Path) -> tuple[str, str]:
    """Devuelve (texto completo, texto del cuadro por sector desde la fila Total).

    En los boletines viejos el título "Cuadro 1.1" es una imagen: por eso el
    cuadro se ubica por su fila Total (personas, %, empleadores, %, $ masa).
    """
    import fitz
    with fitz.open(ruta) as pdf:
        paginas = [re.sub(r'\s+', ' ', p.get_text()) for p in pdf]
    for pagina in paginas:
        m = TOTAL.search(pagina)
        if m and 'agricultura' in _plano(pagina[m.start():m.start() + 300]):
            return ' '.join(paginas), pagina[m.start():]
    raise ValueError(f'{ruta.name}: no se encontró el cuadro de personas trabajadoras por sector')


def extraer_cuadro(textos, anio: int, mes: int) -> dict:
    completo, texto = textos if isinstance(textos, tuple) else (textos, textos)
    plano = _plano(texto)
    esperado = _plano(f'{MESES[mes - 1]} de {anio}')
    if esperado.replace(' ', '') not in _plano(completo).replace(' ', ''):
        raise ValueError(f'El boletín no corresponde a {MESES[mes - 1]} de {anio}')
    total = re.search(rf'Total\s+{NUM}\s+{PCT}\s+{NUM}\s+{PCT}', texto)
    if not total:
        raise ValueError('No se encontró la fila Total del Cuadro 1.1')
    a_int = lambda s: int(s.replace('.', ''))
    trabajadores, empleadores = a_int(total.group(1)), a_int(total.group(2))
    sectores, suma_t, suma_e = {}, 0, 0
    for nombre in list(SECTORES) + ['Sin dato']:
        # Tolera espacios partidos dentro de palabras ("as egurada").
        patron = r'\s*'.join(re.escape(c) for c in _plano(nombre).replace(' ', ''))
        m = re.search(rf'{patron}\s+{NUM}\s+{PCT}\s+{NUM}\s+{PCT}', plano)
        if not m:
            if nombre == 'Sin dato':
                continue
            raise ValueError(f'No se encontró el sector "{nombre}" en el Cuadro 1.1')
        t, e = a_int(m.group(1)), a_int(m.group(2))
        suma_t += t
        suma_e += e
        if nombre != 'Sin dato':
            sectores[nombre] = t
    if suma_t != trabajadores:
        raise ValueError(f'Control: la suma de sectores ({suma_t}) no coincide con el total de personas trabajadoras ({trabajadores})')
    if suma_e != empleadores:
        raise ValueError(f'Control: la suma de sectores ({suma_e}) no coincide con el total de parte empleadora ({empleadores})')
    return {'trabajadores': trabajadores, 'empleadores': empleadores, 'sectores': sectores}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--base', type=Path, default=BASE)
    parser.add_argument('--meses-atras', type=int, default=2, help='Meses ya cargados que se reenvían al Sheet para verificar.')
    parser.add_argument('--desde', help='Carga histórica puntual (AAAA-MM): recorre desde ese mes y reemplaza lo que haya. '
                                        'Solo se usa con la fuente manual srt_historico.')
    args = parser.parse_args()
    historico = None
    if args.desde:
        m = re.fullmatch(r'(20\d{2})-(\d{2})', args.desde)
        if not m:
            raise ValueError('--desde debe tener el formato AAAA-MM')
        historico = (int(m.group(1)), int(m.group(2)))
    ultimo = ultimo_periodo_cargado(args.base, TOTAL_TRABAJADORES)
    if ultimo is None:
        raise ValueError('La base no tiene ningún mes de SRT cargado')
    hoy = date.today()
    # Meses a revisar: los últimos ya cargados (verificación) y los siguientes hasta hoy.
    candidatos = []
    a, m = ultimo
    for _ in range(args.meses_atras):
        a, m = (a - 1, 12) if m == 1 else (a, m - 1)
    a, m = (a, m + 1) if m < 12 else (a + 1, 1)
    if historico:
        a, m = historico
    while (a, m) <= (hoy.year, hoy.month):
        candidatos.append((a, m))
        a, m = (a, m + 1) if m < 12 else (a + 1, 1)
    CARPETA.mkdir(parents=True, exist_ok=True)
    meses, detalle, nuevos, errores_meses = {}, [], [], []
    for anio, mes in candidatos:
        url = url_boletin(anio, mes)
        if not url:
            if (anio, mes) > ultimo:
                break  # todavía no publicado
            continue
        destino = CARPETA / f'srt_formosa_{anio}_{mes:02d}.pdf'
        try:
            destino.write_bytes(_http(url, timeout=120))
            datos = extraer_cuadro(texto_pdf(destino), anio, mes)
        except Exception as error:
            if not historico:
                raise
            # En la carga histórica un boletín ilegible no frena el resto: se informa.
            errores_meses.append(f'{MESES[mes - 1]} {anio}: {error}')
            continue
        valores = {TOTAL_TRABAJADORES: datos['trabajadores'], TOTAL_EMPLEADORES: datos['empleadores']}
        valores.update({SECTORES[n]: v for n, v in datos['sectores'].items()})
        meses[(anio, mes)] = valores
        detalle.append({'periodo': f'{anio}-{mes:02d}', 'url': url, **datos})
        if (anio, mes) > ultimo:
            nuevos.append((anio, mes))
    if not meses:
        raise ValueError('No se pudo obtener ningún boletín de la SRT')
    calculadas = CALCULADAS + [t for t in _calculadas_por_sector(args.base)]
    actualizacion = cargar_meses(args.base, meses, calculadas, sobrescribir=bool(historico),
                                 etiqueta_respaldo='srt_historico' if historico else 'srt')
    if historico:
        for carga in actualizacion['cargas_google']:
            carga['permitir_reemplazos'] = True
    ultimo_publicado = max(meses)
    resultado = {
        'fuente': 'SRT - Boletín por jurisdicción (Formosa)',
        'periodo': f'{ultimo_publicado[0]}-{ultimo_publicado[1]:02d}',
        'periodos': [f'{a}-{m:02d}' for a, m in sorted(meses)],
        'meses_nuevos': [texto_periodo(p) for p in nuevos],
        'carga_historica_desde': args.desde,
        'meses_con_error': errores_meses,
        'detalle': detalle,
        'actualizacion': actualizacion,
    }
    print('RESULTADO_SRT=' + json.dumps(resultado, ensure_ascii=False, default=str))
    print('ESTADO=OK')
    return 0


def _calculadas_por_sector(base: Path) -> list[str]:
    """Columnas 'Var% Intermensual/Interanual' de cada sector, por título."""
    from openpyxl import load_workbook
    libro = load_workbook(base, read_only=True)
    try:
        titulos = [c.value for c in next(libro['BASE DE DATOS'].iter_rows(min_row=1, max_row=1))]
    finally:
        libro.close()
    return [str(t) for t in titulos if t and str(t).startswith('SRT Var% Inter')]


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as error:  # noqa: BLE001
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        sys.exit(1)
