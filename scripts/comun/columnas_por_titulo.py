"""Resolución segura de columnas de Excel mediante encabezados.

Las letras de columna pueden cambiar cuando se insertan indicadores. Este
módulo identifica cada destino por el texto de la fila de encabezados y exige
una única coincidencia antes de permitir una escritura.
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Mapping, Sequence

from openpyxl.utils import get_column_letter


def normalizar_titulo(valor: object) -> str:
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    texto = texto.casefold().replace("%", " porcentaje ")
    return re.sub(r"[^a-z0-9]+", " ", texto).strip()


def _variantes(valor: str | Sequence[str]) -> tuple[str, ...]:
    if isinstance(valor, str):
        return (valor,)
    return tuple(valor)


def resolver_columnas(
    hoja,
    especificaciones: Mapping[str, str | Sequence[str]],
    fila_encabezados: int = 1,
) -> dict[str, int]:
    """Devuelve columnas por clave semántica o falla sin modificar el libro."""
    indice: dict[str, list[int]] = {}
    for columna in range(1, hoja.max_column + 1):
        titulo = normalizar_titulo(hoja.cell(fila_encabezados, columna).value)
        if titulo:
            indice.setdefault(titulo, []).append(columna)

    resueltas: dict[str, int] = {}
    errores: list[str] = []
    for clave, titulos in especificaciones.items():
        coincidencias: set[int] = set()
        for titulo in _variantes(titulos):
            coincidencias.update(indice.get(normalizar_titulo(titulo), []))
        if len(coincidencias) == 1:
            resueltas[clave] = coincidencias.pop()
        elif not coincidencias:
            errores.append(f"{clave}: encabezado no encontrado")
        else:
            letras = ", ".join(get_column_letter(c) for c in sorted(coincidencias))
            errores.append(f"{clave}: encabezado ambiguo en {letras}")

    if errores:
        raise ValueError("No se pudieron resolver columnas por título: " + "; ".join(errores))
    if len(set(resueltas.values())) != len(resueltas):
        raise ValueError("Dos indicadores fueron asociados a la misma columna")
    return resueltas


def letra_columna(columna: int) -> str:
    return get_column_letter(columna)
