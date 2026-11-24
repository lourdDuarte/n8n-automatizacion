"""Prueba del Catálogo API-INDEC (NO escribe en el Sheet ni en la base).

1) IPC: trae el último período publicado y lo compara con lo que cargó la fuente
   'ipc' desde el Excel (scripts/reportes/ipc/).
2) Población de Formosa por departamento (2022-2035) para el trabajo de QGIS.
Guarda todo en reportes/indec/ (sin el token).
"""
from __future__ import annotations

import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from cliente_indec import RAIZ, consultar, consultar_todo  # noqa: E402

SALIDA = RAIZ / 'reportes' / 'indec'


def main() -> int:
    SALIDA.mkdir(parents=True, exist_ok=True)
    resumen = {'consultado_en': datetime.now().isoformat(timespec='seconds')}
    try:
        ultimo = consultar('ipc_nacional', {'select': 'periodo', 'order': 'periodo.desc', 'limit': 1})
        periodo = ultimo[0]['periodo']
        filas = consultar_todo('ipc_nacional', {'periodo': f'eq.{periodo}'},
                               'clasificador,nivel,codigo,region,tipo_de_indicador')
        resumen['ipc'] = {
            'periodo': periodo, 'filas': len(filas),
            'regiones': sorted(Counter(f.get('region') for f in filas).items()),
            'clasificadores': sorted(Counter(f.get('clasificador') for f in filas).items()),
            'tipos': sorted(Counter((f.get('tipo_de_indicador'), f.get('tipo_de_indicador_descripcion')) for f in filas).items()),
            'niveles': sorted(Counter((f.get('clasificador'), f.get('nivel')) for f in filas).items()),
        }
        (SALIDA / f'ipc_{periodo}.json').write_text(json.dumps(filas, ensure_ascii=False, indent=1), encoding='utf-8')
        # Lo que cargó la fuente 'ipc' desde el Excel, para comparar.
        reportes = sorted((RAIZ / 'scripts' / 'reportes' / 'ipc').glob('*.json'), key=lambda r: r.stat().st_mtime)
        if reportes:
            excel = json.loads(reportes[-1].read_text(encoding='utf-8'))
            resumen['ipc']['excel'] = {'archivo': reportes[-1].name, 'periodo': excel.get('periodo'),
                                       'valores': excel.get('valores')}
        print('IPC_OK', periodo, len(filas))
    except Exception as error:  # noqa: BLE001
        resumen['ipc'] = {'error': str(error)}
        print('IPC_ERROR', error)
    try:
        pob = consultar_todo('estimaciones_poblacion_departamento_2022_2035', {'jurisdiccion_cod': 'eq.34'},
                             'departamento_cod,sexo_cod,fecha')
        (SALIDA / 'poblacion_formosa_departamentos_2022_2035.json').write_text(
            json.dumps(pob, ensure_ascii=False, indent=1), encoding='utf-8')
        resumen['poblacion'] = {'filas': len(pob), 'muestra': pob[:3],
                                'departamentos': sorted({(f['departamento_cod'], f['departamento_nom']) for f in pob})}
        print('POBLACION_OK', len(pob))
    except Exception as error:  # noqa: BLE001
        resumen['poblacion'] = {'error': str(error)}
        print('POBLACION_ERROR', error)
    (SALIDA / 'prueba_indec.json').write_text(json.dumps(resumen, ensure_ascii=False, indent=1, default=str), encoding='utf-8')
    ok = 'error' not in resumen['ipc'] and 'error' not in resumen['poblacion']

    # El ejecutor general de n8n necesita recibir un resultado JSON por stdout.
    print(json.dumps({
        'estado': 'OK' if ok else 'ERROR',
        'fuente': 'indec_prueba',
        'periodo': resumen.get('ipc', {}).get('periodo'),
        'detalle': resumen,
    }, ensure_ascii=False, default=str))

    print('ESTADO=' + ('OK' if ok else 'ERROR'))
    return 0 if ok else 1


if __name__ == '__main__':
    raise SystemExit(main())
