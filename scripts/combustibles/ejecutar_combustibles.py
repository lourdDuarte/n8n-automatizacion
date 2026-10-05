"""Precio de combustibles - CECHA (YPF/OPESSA, Capital Federal).

CECHA publica la tabla "EVOLUCIÓN DEL PRECIO DE LOS COMBUSTIBLES" como una
imagen que reemplaza cada mes. Este script:

1. descarga la página y la imagen vigente;
2. lee la tabla con OCR (RapidOCR si está instalado; si no, el OCR que trae
   Windows);
3. toma SUPER y GASOIL de cada mes;
4. controla la lectura antes de escribir:
   - los meses que ya están en la base deben coincidir exactamente con la
     imagen (si no coinciden, el OCR leyó mal y no se escribe nada);
   - cada mes nuevo debe tener los 4 precios, SUPER < PREMIUM,
     GASOIL < EURO y una variación razonable respecto del mes anterior;
5. carga "Precio nafta Super" y "Precio Gasoil" y completa sus variaciones
   con la misma fórmula del mes anterior.
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
from bloque_mensual import cargar_meses, texto_periodo, ultimo_periodo_cargado  # noqa: E402

BASE = RAIZ / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx'
URL = 'https://cecha.org.ar/site/index.php/esncuastas-y-consultas/'
CARPETA = RAIZ / 'fuentes' / 'combustibles'
SUPER, GASOIL = 'Precio nafta Super', 'Precio Gasoil'
CALCULADAS = ['VAR %i.m. Precio nafta Super', 'VAR %i.a. Precio nafta Super',
              'VAR %i.m. Precio gasoil', 'VAR %i.a. Precio gasoil']
MESES = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6, 'jul': 7,
         'ago': 8, 'sep': 9, 'set': 9, 'oct': 10, 'nov': 11, 'dic': 12}
ETIQUETA = re.compile(r'^(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[-–.\s]?(\d{2})$', re.I)
PRECIO = re.compile(r'\d{1,3}(?:[.,]\d{3})*[.,]\d{2}')


def _http(url: str, timeout: int = 60) -> bytes:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (observatorio-formosa)'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def url_imagen(html: str) -> str:
    imagenes = re.findall(r'<img[^>]+src="([^"]+/wp-content/uploads/[^"]+\.(?:jpe?g|png|webp))"', html, re.I)
    imagenes = [u for u in imagenes if 'logo' not in u.lower()]
    if not imagenes:
        raise ValueError('La página de CECHA no tiene la imagen de la tabla de precios')
    return urllib.parse.urljoin(URL, imagenes[-1])


# ------------------------------------------------------------------ OCR
def _ocr_rapid(ruta: Path) -> list[tuple[float, float, str]]:
    from rapidocr_onnxruntime import RapidOCR
    resultado, _ = RapidOCR()(str(ruta))
    return [((b[0][0] + b[2][0]) / 2, (b[0][1] + b[2][1]) / 2, t) for b, t, _c in (resultado or [])]


_PS = r'''
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, $type) { $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op)); $t.Wait(); $t.Result }
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync('__RUTA__')) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
foreach ($line in $result.Lines) { foreach ($w in $line.Words) { '{0}|{1}|{2}' -f ($w.BoundingRect.X + $w.BoundingRect.Width / 2), ($w.BoundingRect.Y + $w.BoundingRect.Height / 2), $w.Text } }
'''


def _ocr_windows(ruta: Path) -> list[tuple[float, float, str]]:
    script = _PS.replace('__RUTA__', str(ruta.resolve()).replace("'", "''"))
    with tempfile.NamedTemporaryFile('w', suffix='.ps1', delete=False, encoding='utf-8-sig') as f:
        f.write(script)
    salida = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', f.name],
                            capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=120)
    if salida.returncode:
        raise RuntimeError('OCR de Windows: ' + ((salida.stderr or '') + (salida.stdout or ''))[-400:])
    if not (salida.stdout or '').strip():
        raise RuntimeError('OCR de Windows no devolvió texto. ' + (salida.stderr or '')[-300:])
    palabras = []
    for linea in salida.stdout.splitlines():
        partes = linea.split('|', 2)
        if len(partes) == 3:
            palabras.append((float(partes[0].replace(',', '.')), float(partes[1].replace(',', '.')), partes[2]))
    return palabras


def leer_imagen(ruta: Path) -> tuple[list, str]:
    """Agranda la imagen (mejora el OCR) y prueba los motores disponibles."""
    try:
        from PIL import Image
        im = Image.open(ruta)
        im = im.resize((im.width * 2, im.height * 2))
        ruta = ruta.with_name(ruta.stem + '_x2.png')
        im.save(ruta)
    except Exception:
        pass
    errores = []
    for nombre, motor in (('RapidOCR', _ocr_rapid), ('Windows OCR', _ocr_windows)):
        try:
            palabras = motor(ruta)
            if palabras:
                return palabras, nombre
        except Exception as error:  # noqa: BLE001
            errores.append(f'{nombre}: {error}')
    raise RuntimeError('No hay OCR disponible para leer la imagen de CECHA. ' + ' | '.join(errores)
                       + ' Solución: pip install rapidocr-onnxruntime pillow')


def tabla(palabras: list) -> dict[tuple, list[float]]:
    """Agrupa por fila: etiqueta mmm-aa y los 4 precios a su derecha."""
    altura = sorted(abs(a[1] - b[1]) for a, b in zip(palabras, palabras[1:]) if abs(a[1] - b[1]) > 0)
    tolerancia = 8.0
    meses = {}
    for x, y, texto in palabras:
        m = ETIQUETA.match(texto.strip().replace(' ', ''))
        if not m:
            continue
        clave = (2000 + int(m.group(2)), MESES[m.group(1).lower()])
        fila = sorted((p for p in palabras if abs(p[1] - y) <= tolerancia and p[0] > x), key=lambda p: p[0])
        # Corta en la etiqueta de la tabla vecina (otro año) si estuviera en la misma línea.
        precios = []
        for px, _py, t in fila:
            if ETIQUETA.match(t.strip().replace(' ', '')):
                break
            precios += [int(re.sub(r'\D', '', v)) / 100 for v in PRECIO.findall(t)]
        meses[clave] = precios
    return meses


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--base', type=Path, default=BASE)
    parser.add_argument('--imagen', type=Path, help='Imagen local (pruebas).')
    args = parser.parse_args()
    CARPETA.mkdir(parents=True, exist_ok=True)
    if args.imagen:
        ruta, url = args.imagen, str(args.imagen)
    else:
        url = url_imagen(_http(URL).decode('utf-8', 'replace'))
        ruta = CARPETA / Path(urllib.parse.urlparse(url).path).name
        ruta.write_bytes(_http(url))
    palabras, motor = leer_imagen(ruta)
    textos = ' '.join(t for _x, _y, t in palabras).upper()
    if not all(k in textos for k in ('SUPER', 'PREMIUM', 'GASOIL', 'EURO')):
        raise ValueError('La imagen de CECHA no tiene las columnas SUPER, PREMIUM, GASOIL y EURO')
    leidos = {k: v for k, v in tabla(palabras).items() if len(v) == 4}
    if not leidos:
        raise ValueError('No se pudo leer ninguna fila completa de la tabla de CECHA')
    ultimo = ultimo_periodo_cargado(args.base, SUPER)
    # Control 1: los meses ya cargados deben coincidir con la imagen.
    from openpyxl import load_workbook
    from bloque_mensual import columna, filas_por_periodo, _indice_columnas
    libro = load_workbook(args.base, read_only=False, data_only=True)
    ws = libro['BASE DE DATOS']
    indice, filas = _indice_columnas(ws), filas_por_periodo(ws)
    c_super, c_gasoil = columna(indice, SUPER), columna(indice, GASOIL)
    comprobados = 0
    for clave in sorted(k for k in leidos if ultimo and k <= ultimo)[-3:]:
        guardado = (ws.cell(filas[clave], c_super).value, ws.cell(filas[clave], c_gasoil).value)
        leido = (leidos[clave][0], leidos[clave][2])
        if any(g in (None, 0) for g in guardado):
            continue
        if abs(guardado[0] - leido[0]) > 0.005 or abs(guardado[1] - leido[1]) > 0.005:
            raise ValueError(f'Control de lectura: {texto_periodo(clave)} en la base es Super {guardado[0]} / Gasoil {guardado[1]}, '
                             f'pero la imagen se leyó {leido[0]} / {leido[1]}. No se escribió nada.')
        comprobados += 1
    libro.close()
    if ultimo and comprobados == 0:
        raise ValueError('Control de lectura: no se pudo comparar ningún mes ya cargado con la imagen. No se escribió nada.')
    # Control 2: meses nuevos coherentes.
    nuevos = {}
    anterior = leidos.get(ultimo) if ultimo else None
    for clave in sorted(k for k in leidos if not ultimo or k > ultimo):
        sup, prem, gas, euro = leidos[clave]
        if not (sup < prem and gas < euro):
            raise ValueError(f'Control: {texto_periodo(clave)} tiene precios incoherentes {leidos[clave]}')
        if anterior and not (0.9 <= sup / anterior[0] <= 1.3 and 0.9 <= gas / anterior[2] <= 1.3):
            raise ValueError(f'Control: variación fuera de rango en {texto_periodo(clave)} ({anterior} -> {leidos[clave]})')
        nuevos[clave] = {SUPER: sup, GASOIL: gas}
        anterior = leidos[clave]
    # Reenvía también el último mes ya cargado para verificar el Sheet.
    a_cargar = dict(nuevos)
    if ultimo in leidos:
        a_cargar[ultimo] = {SUPER: leidos[ultimo][0], GASOIL: leidos[ultimo][2]}
    actualizacion = cargar_meses(args.base, a_cargar, CALCULADAS, etiqueta_respaldo='combustibles')
    ultimo_publicado = max(leidos)
    resultado = {
        'fuente': 'Precio de combustibles - CECHA',
        'periodo': f'{ultimo_publicado[0]}-{ultimo_publicado[1]:02d}',
        'periodos': [f'{a}-{m:02d}' for a, m in sorted(a_cargar)],
        'meses_nuevos': [texto_periodo(k) for k in sorted(nuevos)],
        'imagen': url,
        'ocr': motor,
        'meses_comprobados_contra_base': comprobados,
        'actualizacion': actualizacion,
    }
    print('RESULTADO_COMBUSTIBLES=' + json.dumps(resultado, ensure_ascii=False, default=str))
    print('ESTADO=OK')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as error:  # noqa: BLE001
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        sys.exit(1)
