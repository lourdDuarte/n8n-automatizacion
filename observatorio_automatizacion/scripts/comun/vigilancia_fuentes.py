"""Alerta cuando un indicador deja de traer datos nuevos aunque el organismo ya publicó.

Dos controles, para todos los indicadores de las tres revisiones (generales,
complejos y productivos):

1. INDEC: cada informe de prensa tiene un PDF con nombre propio (ipc_10_26…,
   emae_09_26…). Si en la página del tema aparece un informe nuevo pero el
   indicador sigue en el mismo período, se avisa el mismo día: casi siempre es
   un cambio de nombre o de formato del archivo de datos.
2. Todos: si pasan más días de lo normal sin un período nuevo (38 días los
   mensuales, 100 los trimestrales, 190 los semestrales), se avisa para revisar.

El aviso del informe nuevo sale una vez; el de días sin novedad, como máximo una vez por semana. El estado se guarda
en scripts/reportes/vigilancia_fuentes.json.
"""
from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
ESTADO = RAIZ / 'scripts' / 'reportes' / 'vigilancia_fuentes.json'

# Prefijo del PDF del informe de prensa en la página del tema del INDEC.
INFORMES_INDEC = {
    'ipc': 'ipc_', 'canasta': 'canasta_', 'emae': 'emae_', 'supermercados': 'super_',
    'salarios': 'salarios_', 'pobreza': 'eph_pobreza_', 'ipi': 'ipi_manufacturero_',
    'eph': 'mercado_trabajo_eph_', 'pbi': 'pib_',
}
DIAS_SIN_NOVEDAD = {'eph': 100, 'pbi': 100, 'bess': 100, 'internet': 100, 'pobreza': 190}
DIAS_MENSUAL = 38
# Fuentes que no son indicadores con calendario propio, o que ya tienen su recordatorio.
NO_VIGILAR = {'control_base', 'anses', 'estimaciones', 'ieric_correccion', 'correccion_transferencias',
              'limpieza_motos'}


def _informe_indec(fuente: str) -> str | None:
    """Nombre del PDF del último informe de prensa publicado en la página del tema."""
    from fuentes_web import PAGINAS_INDEC, archivos_en_pagina
    prefijo = INFORMES_INDEC[fuente]
    for url in archivos_en_pagina(PAGINAS_INDEC[fuente], ('.pdf',), timeout=15):
        nombre = url.rsplit('/', 1)[-1].lower()
        if re.match(re.escape(prefijo) + r'\d', nombre):
            return nombre
    return None


def _texto_periodo(p: str) -> str:
    meses = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
    m = re.match(r'(20\d{2})-(\d{2})', p)
    return f'{meses[int(m.group(2)) - 1]} {m.group(1)[2:]}' if m else p


def revisar(resultados: list[dict], hoy: date | None = None) -> dict[str, str]:
    """Devuelve {fuente: aviso} para las fuentes que hay que revisar hoy."""
    hoy = hoy or date.today()
    try:
        estado = json.loads(ESTADO.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        estado = {}
    avisos = {}
    for r in resultados:
        fuente = str(r.get('fuente') or '')
        periodo = str(r.get('periodo') or '')[:7]
        if (not fuente or fuente in NO_VIGILAR or fuente.endswith(('_historico', '_prueba'))
                or not re.match(r'20\d{2}-\d{2}$', periodo) or r.get('estado') == 'ERROR'):
            continue
        pendiente = r.get('estado') == 'PENDIENTE'  # su período es el que falta, no el último cargado
        if pendiente and not estado.get(fuente):
            continue
        st = estado.setdefault(fuente, {})
        informe = None
        if fuente in INFORMES_INDEC:
            try:
                informe = _informe_indec(fuente)
            except Exception:  # noqa: BLE001 - sin conexión a la página: no se avisa
                informe = None
        if not pendiente and (not st.get('periodo') or periodo > st['periodo']):
            # Período nuevo (o primera vez): se toma como referencia.
            st.clear()
            st.update(periodo=periodo, desde=hoy.isoformat())
            if informe:
                st['informe'] = informe
            continue
        if informe and not st.get('informe'):
            st['informe'] = informe
        mensajes = []
        if informe and st.get('informe') and informe != st['informe'] and st.get('avisado_informe') != informe:
            st['avisado_informe'] = informe  # una sola vez por informe
            mensajes.append(f'el INDEC ya publicó un informe nuevo ({informe}) pero no se encontró el dato '
                            'nuevo: puede que haya cambiado el nombre o el formato del archivo')
        dias = (hoy - date.fromisoformat(st.get('desde') or hoy.isoformat())).days
        limite = DIAS_SIN_NOVEDAD.get(fuente, DIAS_MENSUAL)
        ultimo = st.get('avisado_dias')
        if dias > limite and (not ultimo or (hoy - date.fromisoformat(ultimo)).days >= 7):
            st['avisado_dias'] = hoy.isoformat()  # se repite como máximo una vez por semana
            mensajes.append(f'hace {dias} días que no hay un período nuevo (último: {_texto_periodo(st["periodo"])}); '
                            'revisar si el organismo publicó con otro nombre o formato')
        if mensajes:
            avisos[fuente] = '⏰ Revisar: ' + '; '.join(mensajes) + '.'
    try:
        ESTADO.parent.mkdir(parents=True, exist_ok=True)
        ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=1), encoding='utf-8')
    except OSError:
        pass
    return avisos
