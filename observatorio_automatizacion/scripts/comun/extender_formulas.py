"""Extiende hacia abajo las columnas calculadas del Sheet (variaciones, acumulados, etc.).

Cuando la automatización carga los datos de un mes, las columnas con fórmula que dependen de esos
datos (p. ej. "Var % ... con respecto al mes anterior" = NG93/NG92*100-100) quedaban vacías, porque
la fórmula está solo en el Sheet. Acá se copia la fórmula de la última fila que la tiene, igual que
cuando se arrastra a mano, pero solo si:
  - la celda de destino está vacía (nunca se pisa nada),
  - la fórmula usa, en su misma fila, alguna columna que se acaba de cargar (o otra columna que se
    extendió en esta misma pasada), y
  - las columnas que usa de su misma fila tienen dato en la fila de destino,
  - no apunta a otras pestañas (esas filas no se corresponden).
"""
from __future__ import annotations

import re

REF = re.compile(r"(?<![A-Za-z0-9_'!.$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])")
FILAS_ATRAS = 3  # hasta cuántas filas arriba se busca la última fórmula
VENTANA = 40  # columnas a cada lado del bloque cargado donde también se extienden fórmulas cuyos datos ya están


def _letra(n: int) -> str:
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _numero(letras: str) -> int:
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def _es_formula(v) -> bool:
    return isinstance(v, str) and v.startswith('=')


def _sin_textos(formula: str) -> str:
    return re.sub(r'"[^"]*"', '""', formula)


def referencias(formula: str):
    """[(col, fila, fila_absoluta)] de las celdas de esta misma pestaña."""
    return [(_numero(m.group(2)), int(m.group(4)), bool(m.group(3))) for m in REF.finditer(_sin_textos(formula))]


def desplazar(formula: str, filas: int) -> str:
    partes = re.split(r'("[^"]*")', formula)
    for i in range(0, len(partes), 2):
        partes[i] = REF.sub(lambda m: m.group(0) if m.group(3) else
                            f'{m.group(1)}{m.group(2)}{int(m.group(4)) + filas}', partes[i])
    return ''.join(partes)


def planear(bloque: list[list], fila_inicio: int, fila: int, columnas: set[int]) -> list[tuple[str, str]]:
    """bloque: filas fila_inicio..fila leídas con render FORMULA (desde la columna A).
    Devuelve [(celda, fórmula)] a escribir en la fila `fila` y en las filas vacías intermedias."""
    def celda(r, c):
        f = bloque[r - fila_inicio] if 0 <= r - fila_inicio < len(bloque) else []
        return f[c - 1] if c - 1 < len(f) else ''

    ancho = max((len(f) for f in bloque), default=0)
    fuentes = set(columnas)
    plan, hechas = [], set()
    cambio = True
    while cambio:  # una columna extendida puede alimentar a otra (p. ej. acumulado -> variación)
        cambio = False
        for c in range(1, ancho + 1):
            if c in hechas or c in columnas:
                continue
            if celda(fila, c) not in (None, ''):
                continue
            r0 = next((r for r in range(fila - 1, max(fila_inicio, fila - FILAS_ATRAS) - 1, -1)
                       if _es_formula(celda(r, c))), None)
            if r0 is None:
                continue
            base = celda(r0, c)
            if '!' in base:
                continue
            refs = referencias(base)
            misma = {col for col, f, absoluta in refs if f == r0 and not absoluta and col != c}
            if not misma:
                continue
            # Depende de lo que se cargó, o está en el mismo bloque del indicador y usa columnas que ya
            # tienen dato en la fila (p. ej. supermercados: variaciones de NG, que es una fórmula a otra pestaña).
            if not (misma & fuentes) and not (min(columnas) - VENTANA <= c <= max(columnas) + VENTANA):
                continue
            for r in range(r0 + 1, fila + 1):
                if celda(r, c) not in (None, ''):
                    continue
                if any(celda(r, col) in (None, '') and (r, col) not in hechas_celdas(plan) for col in misma):
                    continue
                plan.append((f'{_letra(c)}{r}', desplazar(base, r - r0)))
            hechas.add(c)
            fuentes.add(c)
            cambio = True
    return plan


def hechas_celdas(plan):
    out = set()
    for a1, _ in plan:
        m = re.match(r'([A-Z]+)(\d+)$', a1)
        out.add((int(m.group(2)), _numero(m.group(1))))
    return out
