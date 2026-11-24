"""Producción agrícola de Formosa por cultivo y campaña (Estimaciones Agrícolas, MAGyP).

Fuente oficial: https://datosestimaciones.magyp.gob.ar (Dirección de Estimaciones
Agrícolas). Se descarga el dataset completo por departamento, se suma Formosa y el
total país por cultivo y campaña, y se completa el bloque "Producción" de la pestaña
productiva del Sheet (Cultivo, Campaña, superficies, producción y % de participación).

* Cada fila se ubica por (cultivo, campaña); nunca por posición.
* Las campañas nuevas se agregan al final del bloque.
* Las estimaciones de la campaña en curso se corrigen cuando el organismo las revisa.
* Se relee lo escrito y se verifica.
"""
from __future__ import annotations

import re
import sys
import unicodedata
import urllib.parse
import urllib.request
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
from sheet_directo import API, ClienteSheets, clave, es_formula, iguales, letra  # noqa: E402
from sheet_productivos import _leer, _titulo_hoja  # noqa: E402

URL = 'https://datosestimaciones.magyp.gob.ar/reportes.php?reporte=Estimaciones'
DESDE = '2020/21'
TITULOS = ['Cultivo', 'Campaña', 'Sup. Sembrada Formosa (Ha)', 'Sup. Sembrada Nación (Ha)',
           'Sup. Cosechada Formosa (Ha)', 'Sup. Cosechada Nación (Ha)', 'Producción Formosa (Tn)',
           'Producción Nación (Tn)', '% Participación Formosa']
LIMITE_CORRECCIONES = 60


def _norm(t) -> str:
    t = unicodedata.normalize('NFD', str(t or ''))
    return ' '.join(''.join(c for c in t if unicodedata.category(c) != 'Mn').casefold().split())


def descargar(timeout: int = 240) -> str:
    datos = urllib.parse.urlencode({'Dataset': 'Dataset'}).encode()
    req = urllib.request.Request(URL, data=datos, headers={'User-Agent': 'Mozilla/5.0 (Observatorio Formosa/1.0)'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        contenido = r.read()
    texto = contenido.decode('latin-1')
    if 'Provincia' not in texto[:400] or 'Campa' not in texto[:400]:
        raise ValueError('La descarga de Estimaciones Agrícolas no es el dataset esperado')
    return texto


def agregar(texto: str) -> dict:
    """{(cultivo, campaña): [sembrada F, sembrada N, cosechada F, cosechada N, producción F, producción N]}"""
    total = {}
    for linea in texto.splitlines():
        c = [x.strip().strip('"') for x in linea.split(';')]
        if len(c) < 12 or not c[0].isdigit():
            continue
        cultivo, campana = c[5], c[7]
        if not re.fullmatch(r'\d{4}/\d{2}', campana) or campana < DESDE:
            continue
        try:
            nums = [float(x or 0) for x in c[8:11]]
        except ValueError:
            continue
        a = total.setdefault((cultivo, campana), [0.0] * 6)
        a[1] += nums[0]; a[3] += nums[1]; a[5] += nums[2]
        if _norm(c[1]) == 'formosa':
            a[0] += nums[0]; a[2] += nums[1]; a[4] += nums[2]
    return {k: v for k, v in total.items() if v[0] > 0}


def actualizar_sheet(datos: dict, cliente: ClienteSheets | None = None) -> dict:
    cliente = cliente or ClienteSheets()
    titulo = _titulo_hoja(cliente)
    filas = _leer(cliente, titulo, 'A1:BZ400', 'FORMULA')
    fila_enc = next((n for n, f in enumerate(filas[:15], start=1)
                     if {clave('Cultivo'), clave('Campaña')} <= {clave(v) for v in f}), None)
    if not fila_enc:
        raise RuntimeError('No se encontró el encabezado Cultivo / Campaña en la pestaña productiva')
    enc = filas[fila_enc - 1]
    # La tabla es el bloque donde 'Cultivo' va seguido de 'Campaña'. A la derecha puede haber
    # otros cuadros con una columna 'Cultivo' (p. ej. rindes por departamento): no se tocan.
    inicios = [i for i in range(len(enc) - 1)
               if clave(enc[i]) == clave('Cultivo') and clave(enc[i + 1]) == clave('Campaña')]
    if len(inicios) != 1:
        raise RuntimeError(f'Bloque Cultivo/Campaña: {len(inicios)} coincidencias en el Sheet')
    desde, hasta = inicios[0], inicios[0] + len(TITULOS) + 3
    cols = []
    for t in TITULOS:
        encontradas = [i + 1 for i, v in enumerate(enc) if desde <= i < hasta and clave(v) == clave(t)]
        if len(encontradas) != 1:
            raise RuntimeError(f'Columna {t!r}: {len(encontradas)} coincidencias en el Sheet')
        cols.append(encontradas[0])
    c_cult, c_camp = cols[0], cols[1]

    def celda(n, col):
        f = filas[n - 1] if n - 1 < len(filas) else []
        return f[col - 1] if col - 1 < len(f) else None

    existentes, ultima = {}, fila_enc
    for n in range(fila_enc + 1, len(filas) + 1):
        cu, ca = celda(n, c_cult), celda(n, c_camp)
        if clave(cu) == clave('Cultivo') and clave(ca) == clave('Campaña'):
            break  # otro cuadro armado debajo (con su propio encabezado): no es parte de la tabla
        if cu not in (None, ''):
            ultima = n
            existentes.setdefault((_norm(cu), str(ca).strip()), []).append(n)

    # % de participación: se copia la fórmula de la última fila que la tenga (como hace el equipo),
    # adaptada a la fila nueva; si no hay ninguna, se escribe el porcentaje.
    c_pct = cols[-1]
    fila_modelo = next((n for n in range(len(filas), fila_enc, -1) if es_formula(celda(n, c_pct))), None)

    def formula_pct(n):
        if not fila_modelo:
            return None
        return re.sub(rf'(?<![A-Z0-9])(\$?[A-Z]{{1,3}}\$?){fila_modelo}(?!\d)', rf'\g<1>{n}', celda(fila_modelo, c_pct))

    info = {'periodos': [], 'escritas': 0, 'ya_estaban': 0, 'esperadas': 0, 'conservadas': [],
            'no_encontradas': [], 'errores': [], 'cambios': [], 'nuevas_filas': 0}
    texto_crudo, numerico, cambios = [], [], []
    siguiente = ultima + 1
    for (cultivo, campana) in sorted(datos, key=lambda k: (_norm(k[0]), k[1])):
        v = datos[(cultivo, campana)]
        pct = v[4] / v[5] * 100 if v[5] else 0.0
        valores = [round(x) for x in v] + [pct]
        info['esperadas'] += len(TITULOS)
        destino = existentes.get((_norm(cultivo), campana), [])
        if len(destino) > 1:
            info['errores'].append(f'{cultivo} {campana}: {len(destino)} filas repetidas en el Sheet')
            continue
        if destino:
            n = destino[0]
            info['ya_estaban'] += 2
        else:
            n = siguiente
            if any(celda(n, col) not in (None, '') for col in cols):
                info['errores'].append(f'fila {n} no está vacía para agregar {cultivo} {campana}')
                continue
            siguiente += 1
            info['nuevas_filas'] += 1
            texto_crudo += [{'range': f"'{titulo}'!{letra(c_cult)}{n}", 'values': [[cultivo]]},
                            {'range': f"'{titulo}'!{letra(c_camp)}{n}", 'values': [[campana]]}]
            cambios += [{'celda': f'{letra(c_cult)}{n}', 'antes': None, 'despues': cultivo},
                        {'celda': f'{letra(c_camp)}{n}', 'antes': None, 'despues': campana}]
        for col, valor in zip(cols[2:], valores):
            previo = celda(n, col) if destino else None
            if es_formula(previo):
                continue  # el Sheet lo calcula con su fórmula: no se pisa
            if col == cols[-1] and formula_pct(n):
                nuevo = formula_pct(n)
                numerico.append({'range': f"'{titulo}'!{letra(col)}{n}", 'values': [[nuevo]]})
                cambios.append({'celda': f'{letra(col)}{n}', 'antes': previo, 'despues': nuevo, 'titulo': f'{cultivo} {campana}',
                                'formula': True})
                continue
            if col == cols[-1]:
                # % con formato de porcentaje como el resto de la columna (0,0398 -> 3,98%).
                fraccion = valor / 100
                if isinstance(previo, (int, float)) and abs(previo - fraccion) < 0.00006:
                    info['ya_estaban'] += 1
                    continue
                nuevo = f'{valor:.2f}%'
            else:
                if previo not in (None, '') and iguales(previo, valor):
                    info['ya_estaban'] += 1
                    continue
                nuevo = valor
            numerico.append({'range': f"'{titulo}'!{letra(col)}{n}", 'values': [[nuevo]]})
            cambios.append({'celda': f'{letra(col)}{n}', 'antes': previo, 'despues': nuevo,
                            'titulo': f'{cultivo} {campana}'})
    correcciones = [c for c in cambios if c['antes'] not in (None, '') and not c.get('formula')]
    if len(correcciones) > LIMITE_CORRECCIONES:
        raise RuntimeError(f'se iban a corregir {len(correcciones)} datos ya cargados; no se escribió nada por seguridad')
    for datos_escribir, modo in ((texto_crudo, 'RAW'), (numerico, 'USER_ENTERED')):
        if datos_escribir:
            r = cliente.sesion.post(f'{API}/values:batchUpdate', json={'valueInputOption': modo, 'data': datos_escribir}, timeout=90)
            r.raise_for_status()
    if cambios:
        leidas = _leer(cliente, titulo, 'A1:BZ400', 'UNFORMATTED_VALUE')
        for c in cambios:
            if c.get('formula'):
                continue
            m = re.match(r'([A-Z]+)(\d+)', c['celda'])
            col = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m.group(1))))
            fila = leidas[int(m.group(2)) - 1] if int(m.group(2)) - 1 < len(leidas) else []
            obtenido = fila[col - 1] if col - 1 < len(fila) else None
            esperado = c['despues']
            if isinstance(esperado, str) and esperado.endswith('%'):
                esperado = float(esperado[:-1]) / 100
            if not (iguales(obtenido, esperado) or (isinstance(obtenido, (int, float)) and isinstance(esperado, float) and abs(obtenido - esperado) < 0.00006)):
                raise RuntimeError(f"verificación fallida en {c['celda']}: se leyó {obtenido!r}")
    info['escritas'] = len(cambios)
    info['ya_estaban'] = max(0, info['esperadas'] - len(cambios) - len(info['errores']) * len(TITULOS))
    info['cambios'] = cambios[:300]
    info['periodos'] = sorted({k[1] for k in datos})[-2:]
    return info


def ejecutar(cliente: ClienteSheets | None = None) -> dict:
    return actualizar_sheet(agregar(descargar()), cliente)


if __name__ == '__main__':
    import json
    try:
        print(json.dumps(ejecutar(), ensure_ascii=False, indent=2, default=str))
        print('ESTADO=OK')
    except Exception as error:
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        raise SystemExit(1)
