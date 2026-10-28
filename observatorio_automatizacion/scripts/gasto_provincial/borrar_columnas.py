"""Borra del Sheet (hoja BASE DE DATOS) las columnas "Gasto provincial MECON - ...".

Pedido de la responsable (19/10/2026): por ahora ese bloque no se usa. Solo se borran las
columnas cuyo encabezado empieza exactamente con "Gasto provincial MECON - " (como máximo 18);
no se toca ninguna otra columna. El indicador queda desactivado en la revisión diaria
(ejecutar_general.py), así que no se vuelve a crear solo.

Uso
    python borrar_columnas.py            # borra
    python borrar_columnas.py --revisar  # solo muestra qué borraría
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / 'scripts' / 'comun'))

PREFIJO = 'Gasto provincial MECON - '
MAXIMO = 18


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--revisar', action='store_true', help='No borra, solo informa.')
    args = ap.parse_args()

    from sheet_directo import API, HOJA, ClienteSheets, letra
    cliente = ClienteSheets()
    encabezados = (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0]
    columnas = [i for i, t in enumerate(encabezados) if str(t or '').startswith(PREFIJO)]  # índices desde 0
    if not columnas:
        print('No hay columnas "Gasto provincial MECON" en el Sheet: no hay nada para borrar.')
        print('ESTADO=OK')
        return 0
    if len(columnas) > MAXIMO:
        raise RuntimeError(f'Se encontraron {len(columnas)} columnas con ese nombre (máximo esperado {MAXIMO}); no se borra nada.')
    detalle = [f'{letra(i + 1)}: {encabezados[i]}' for i in columnas]
    if args.revisar:
        print(json.dumps({'se_borrarian': detalle}, ensure_ascii=False, indent=1))
        print('ESTADO=OK')
        return 0

    r = cliente.sesion.get(API, params={'fields': 'sheets(properties(sheetId,title))'}, timeout=60)
    r.raise_for_status()
    hoja_id = next(s['properties']['sheetId'] for s in r.json()['sheets'] if s['properties']['title'] == HOJA)
    # De derecha a izquierda, para que los índices no se corran al borrar.
    pedidos = [{'deleteDimension': {'range': {'sheetId': hoja_id, 'dimension': 'COLUMNS',
                                              'startIndex': i, 'endIndex': i + 1}}}
               for i in sorted(columnas, reverse=True)]
    cliente.sesion.post(f'{API}:batchUpdate', json={'requests': pedidos}, timeout=90).raise_for_status()

    quedan = [t for t in (cliente.leer('1:1', render='FORMATTED_VALUE') or [[]])[0] if str(t or '').startswith(PREFIJO)]
    if quedan:
        raise RuntimeError(f'Después de borrar quedaron {len(quedan)} columnas "Gasto provincial MECON".')
    marca = RAIZ / 'scripts' / 'reportes' / 'gasto_provincial' / 'historico_cargado.txt'
    if marca.is_file():
        marca.unlink()
    print(json.dumps({'columnas_borradas': len(columnas), 'detalle': detalle}, ensure_ascii=False, indent=1))
    print(f'NOTA_TELEGRAM=🗑️ Se borraron del Sheet las {len(columnas)} columnas "Gasto provincial MECON" (pedido de la responsable). El indicador queda desactivado.')
    print('ESTADO=OK')
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as error:  # noqa: BLE001
        print(f'ERROR={error}')
        print('ESTADO=ERROR')
        raise SystemExit(1)
