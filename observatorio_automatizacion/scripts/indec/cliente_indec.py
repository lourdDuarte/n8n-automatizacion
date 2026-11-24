"""Cliente del Catálogo API-INDEC (https://api.indec.gob.ar/v1).

El token (JWT) lo genera la usuaria en su perfil del catálogo y lo guarda en
config/indec_token.txt (o en la variable de entorno INDEC_TOKEN). Este módulo
lo lee solo para mandarlo en el encabezado de cada consulta: nunca lo imprime
ni lo guarda en reportes o logs.
"""
from __future__ import annotations

import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
BASE = 'https://api.indec.gob.ar/v1'
ARCHIVO_TOKEN = RAIZ / 'config' / 'indec_token.txt'


def leer_token() -> str:
    token = os.environ.get('INDEC_TOKEN', '').strip()
    if not token and ARCHIVO_TOKEN.is_file():
        token = ARCHIVO_TOKEN.read_text(encoding='utf-8-sig').strip()
    if token.lower().startswith('bearer '):
        token = token[7:].strip()
    token = token.strip().strip('"').strip("'")
    if not token:
        raise RuntimeError('Falta el token del Catálogo API-INDEC (config/indec_token.txt).')
    return token


def consultar(tabla: str, parametros: dict | None = None, timeout: int = 60, reintentos: int = 3) -> list[dict]:
    """Una página (máx. 1000 filas) con filtros estilo PostgREST."""
    url = f'{BASE}/{tabla}'
    if parametros:
        url += '?' + urllib.parse.urlencode(parametros, safe='.,()*:')
    pedido = urllib.request.Request(url, headers={
        'Authorization': 'Bearer ' + leer_token(),
        'Accept': 'application/json',
        'User-Agent': 'Observatorio Formosa (DAEE)',
    })
    for intento in range(reintentos):
        try:
            with urllib.request.urlopen(pedido, timeout=timeout) as r:
                return json.loads(r.read().decode('utf-8'))
        except urllib.error.HTTPError as error:
            cuerpo = error.read().decode('utf-8', 'replace')[:300]
            if error.code in (401, 403):
                raise RuntimeError(f'INDEC rechazó el token ({error.code}): revisar config/indec_token.txt. {cuerpo}') from None
            if error.code < 500 or intento == reintentos - 1:
                raise RuntimeError(f'INDEC API {tabla}: error {error.code}: {cuerpo}') from None
        except (urllib.error.URLError, TimeoutError) as error:
            if intento == reintentos - 1:
                raise RuntimeError(f'INDEC API {tabla}: sin conexión ({error})') from None
        time.sleep(3 * (intento + 1))
    return []


def consultar_todo(tabla: str, parametros: dict, orden: str, maximo: int = 50000) -> list[dict]:
    """Recorre todas las páginas (orden estable por la clave de la tabla)."""
    filas, offset = [], 0
    while offset < maximo:
        pagina = consultar(tabla, {**parametros, 'order': orden, 'limit': 1000, 'offset': offset})
        filas += pagina
        if len(pagina) < 1000:
            break
        offset += 1000
    return filas
