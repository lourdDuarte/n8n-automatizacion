"""Datos para completar la descripción de Transferencias automáticas (no escribe en el Sheet).

Fuente: Dirección Nacional de Asuntos Provinciales (Ministerio de Economía), "Información consolidada"
de transferencias automáticas (columna Total), e IPC Nacional del INDEC. Para el último mes, si el INDEC
todavía no publicó el IPC, se usa la misma inflación estimada que la revisión (INDEC + REM del BCRA).

Calcula, para Formosa, Chaco, Corrientes, Misiones y el total Provincias + CABA:
  - variación real interanual del mes;
  - variación real del acumulado del año (cada mes llevado a precios constantes con el IPC);
  - variación real del mes contra el mismo mes de dos y tres años antes (p. ej. octubre 2023).
Uso: python comparar_transferencias.py [AAAA-MM]   (por defecto, el último mes publicado)
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ejecutar_transferencias_fiscales import (MESES_HOJA, PAGINA_DNAP, _consolidada, _indices_ipc_nacion,  # noqa: E402
                                              descargar, inflacion_interanual_estimada, normalizar)

PROVINCIAS = ['formosa', 'chaco', 'corrientes', 'misiones']


class _LibroXls:
    """Los años viejos de la DNAP están en .xls: misma interfaz que openpyxl (sheetnames, iter_rows)."""

    class _Hoja:
        def __init__(self, h):
            self.h = h

        def iter_rows(self, values_only=True):
            for r in range(self.h.nrows):
                yield tuple(self.h.row_values(r))

    def __init__(self, contenido):
        import xlrd
        self.libro = xlrd.open_workbook(file_contents=contenido)
        self.sheetnames = self.libro.sheet_names()

    def __getitem__(self, nombre):
        return self._Hoja(self.libro.sheet_by_name(nombre))


def _consolidada_cualquiera(anio):
    try:
        return _consolidada(anio)[0]
    except RuntimeError as error:
        if '.xls' not in str(error):
            raise
    import re
    from urllib.parse import urljoin
    pagina = descargar(PAGINA_DNAP).decode('utf-8', errors='replace')
    enlaces = [urljoin('https://www.argentina.gob.ar/', h.replace('blank:#', ''))
               for h in re.findall(r'href=["\']([^"\']+)["\']', pagina, flags=re.I)]
    url = next(u for u in enlaces if re.search(rf'informacion_consolidada_?{anio}[^/]*\.xls$', u, flags=re.I))
    return _LibroXls(descargar(url))


def _mes(libro, mes):
    """{provincia: millones de $} de la columna Total del mes, más 'provincias_caba'."""
    nombre = next((h for h in libro.sheetnames if normalizar(h).startswith(MESES_HOJA[mes - 1][:4])), None)
    if not nombre:
        return None
    filas = list(libro[nombre].iter_rows(values_only=True))
    col = None
    for fila in filas[:15]:
        for k, v in enumerate(fila):
            if isinstance(v, str) and normalizar(v) == 'total' and k > 10:
                col = k
    if col is None:
        return None
    out = {}
    for fila in filas:
        if fila and isinstance(fila[0], str) and col < len(fila) and isinstance(fila[col], (int, float)):
            out[normalizar(fila[0])] = float(fila[col])
    if not out.get('formosa') or not out.get('provincias'):
        return None
    res = {p: out.get(p) for p in PROVINCIAS}
    res['provincias_caba'] = out['provincias'] + (out.get('c.a.b.a') or 0.0)
    return res


def main() -> int:
    libros = {}

    def libro(anio):
        if anio not in libros:
            libros[anio] = _consolidada_cualquiera(anio)
        return libros[anio]

    if len(sys.argv) > 1:
        anio, mes = int(sys.argv[1][:4]), int(sys.argv[1][5:7])
    else:
        from datetime import datetime
        anio = datetime.now().year
        mes = max(m for m in range(1, 13) if _mes(libro(anio), m))
    ipc = _indices_ipc_nacion()
    clave = f'{anio}{mes:02d}'
    estimado = None
    if clave not in ipc:  # el INDEC todavía no publicó el mes: inflación estimada (igual que la revisión)
        tasa, supuesto, origen = inflacion_interanual_estimada(f'{anio}-{mes:02d}-01')
        ipc[clave] = ipc[f'{anio - 1}{mes:02d}'] * (1 + tasa / 100)
        estimado = {'interanual': round(tasa, 2), 'supuesto_mes': round(supuesto, 2), 'origen': origen}

    datos = {a: {m: _mes(libro(a), m) for m in range(1, mes + 1)} for a in (anio, anio - 1)}
    for a in (anio - 2, anio - 3):
        try:
            datos[a] = {mes: _mes(libro(a), mes)}
        except Exception as error:  # noqa: BLE001
            datos[a] = {mes: None}
            print(f'AVISO: {a}: {error}')
    salida = {'periodo': f'{anio}-{mes:02d}', 'ipc_estimado': estimado, 'fuente': 'DNAP (Ministerio de Economía) e IPC INDEC',
              'unidades': 'millones de $; variaciones en %', 'por_provincia': {}}
    for p in PROVINCIAS + ['provincias_caba']:
        x = datos[anio][mes][p] if datos[anio][mes] else None
        y = datos[anio - 1][mes][p] if datos[anio - 1][mes] else None
        r = {'mes': x, 'mes_anio_anterior': y}
        i, i1 = ipc[clave], ipc[f'{anio - 1}{mes:02d}']
        if x and y:
            r['real_interanual'] = round((x / i) / (y / i1) * 100 - 100, 1)
            r['nominal_interanual'] = round(x / y * 100 - 100, 1)
        meses_ok = all(datos[a][m] and datos[a][m][p] for a in (anio, anio - 1) for m in range(1, mes + 1))
        if meses_ok:
            act = sum(datos[anio][m][p] for m in range(1, mes + 1))
            ant = sum(datos[anio - 1][m][p] for m in range(1, mes + 1))
            act_r = sum(datos[anio][m][p] / ipc[f'{anio}{m:02d}'] for m in range(1, mes + 1))
            ant_r = sum(datos[anio - 1][m][p] / ipc[f'{anio - 1}{m:02d}'] for m in range(1, mes + 1))
            r.update(acumulado=act, acumulado_anio_anterior=ant, nominal_acumulado=round(act / ant * 100 - 100, 1),
                     real_acumulado=round(act_r / ant_r * 100 - 100, 1))
        for a in (anio - 2, anio - 3):
            z = datos[a][mes][p] if datos[a].get(mes) else None
            k = f'{a}{mes:02d}'
            if x and z and k in ipc:
                r[f'real_vs_{MESES_HOJA[mes - 1]}_{a}'] = round((x / i) / (z / ipc[k]) * 100 - 100, 1)
                r[f'{MESES_HOJA[mes - 1]}_{a}'] = z
        salida['por_provincia'][p] = r
    print('COMPARACION=' + json.dumps(salida, ensure_ascii=False))
    print('NOTA_TELEGRAM=Comparación de transferencias calculada (no se escribió nada en el Sheet).')
    print('ESTADO=OK')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
