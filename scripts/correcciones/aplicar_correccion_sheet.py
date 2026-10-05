"""Corrección única y verificada de celdas del Google Sheet (uso manual).

El archivo JSON puede traer:
* "valores":  [{periodo, titulo, valor}]  -> datos a escribir (pueden reemplazar lo cargado);
* "formulas": [{periodo, titulo, formula, esperado}] -> fórmulas en coordenadas del Sheet,
  con {fila} = fila del período y {col:Título} = columna de ese encabezado;
* "controles": [{periodo, titulo, esperado}] -> valores que deben quedar después.

Fila por período y columna por encabezado; se relee y verifica. No corre en la
revisión diaria: se ejecuta una vez desde n8n ("Elegir indicador puntual").
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))
from sheet_directo import ClienteSheets, HOJA, Mapa, escribir_actualizaciones, letra, periodo_de  # noqa: E402


def _ubicar(mapa: Mapa, periodo: str, titulo: str) -> tuple[int, int]:
    fila, n = mapa.fila(periodo_de(periodo + '-01'))
    if not fila:
        raise ValueError(f'período {periodo}: {n} filas en el Sheet')
    col, n = mapa.columna(titulo)
    if not col:
        raise ValueError(f'encabezado {titulo!r}: {n} coincidencias en el Sheet')
    return fila, col


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--archivo', required=True)
    args = ap.parse_args()
    ruta = Path(args.archivo)
    if not ruta.is_absolute():
        ruta = Path(__file__).resolve().parent / ruta.name
    try:
        datos = json.loads(ruta.read_text(encoding='utf-8'))
        fuente = datos.get('fuente', 'correccion')
        cambios, errores = [], []
        # 1) Valores
        cargas = {}
        for item in datos.get('valores', []):
            carga = cargas.setdefault(item['periodo'], {'fuente': fuente, 'hoja': HOJA, 'periodo': item['periodo'] + '-01',
                                                        'valores': {}, 'permitir_reemplazos': True})
            carga['valores'][item['titulo']] = item['valor']
        if cargas:
            res = escribir_actualizaciones(list(cargas.values()), RAIZ / 'base' / 'BASE DE DATOS INDICADORES PROPIOS.xlsx')
            info = res.get('por_fuente', {}).get(fuente, {})
            cambios += info.get('cambios', [])
            errores += info.get('errores', [])
            errores += [f'{c}: no se pisó' for c in info.get('conservadas', [])]
        # 2) Fórmulas (coordenadas del Sheet)
        cliente = ClienteSheets()
        mapa = Mapa((cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0], [f[0] if f else None for f in cliente.leer('A:A')])
        for item in datos.get('formulas', []):
            fila, col = _ubicar(mapa, item['periodo'], item['titulo'])

            def col_de(m):
                c, n = mapa.columna(m.group(1))
                if not c:
                    raise ValueError(f'encabezado {m.group(1)!r}: {n} coincidencias')
                return letra(c)
            formula = re.sub(r'\{col:([^}]+)\}', col_de, item['formula']).replace('{fila}', str(fila)).replace('{fila_ant}', str(fila - 1))
            celda = f'{letra(col)}{fila}'
            antes = (cliente.leer(celda, render='FORMULA') or [[None]])[0][0]
            if str(antes).replace(' ', '') != formula.replace(' ', ''):
                cliente.escribir([{'range': f"'{HOJA}'!{celda}", 'values': [[formula]]}])
                cambios.append({'celda': celda, 'antes': antes, 'despues': formula})
            if 'esperado' in item:
                obtenido = (cliente.leer(celda) or [[None]])[0][0]
                if not isinstance(obtenido, (int, float)) or abs(obtenido - item['esperado']) > 0.5:
                    errores.append(f'{celda}: se esperaba {item["esperado"]} y quedó {obtenido!r}')
        # 3) Controles
        for item in datos.get('controles', []):
            fila, col = _ubicar(mapa, item['periodo'], item['titulo'])
            obtenido = (cliente.leer(f'{letra(col)}{fila}') or [[None]])[0][0]
            if not isinstance(obtenido, (int, float)) or abs(obtenido - item['esperado']) > 0.5:
                errores.append(f'{letra(col)}{fila}: se esperaba {item["esperado"]} y quedó {obtenido!r}')
        periodos = sorted({x['periodo'] for x in datos.get('valores', []) + datos.get('formulas', [])})
        salida = {'fuente': fuente, 'periodo': periodos[-1] if periodos else None, 'periodos': periodos,
                  'descripcion': datos.get('descripcion'),
                  'actualizacion': {'celdas_actualizadas': len(cambios), 'cambios': cambios, 'errores': errores}}
        print(json.dumps(salida, ensure_ascii=False, indent=2, default=str))
        if errores:
            print('ERROR=' + '; '.join(errores))
            print('ESTADO=ERROR')
            return 1
        print('ESTADO=OK')
        return 0
    except Exception as error:
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
