from __future__ import annotations

import math
from datetime import date


class ErrorControlFuente(ValueError):
    """La fuente terminó, pero su resultado no es seguro para publicar."""


def validar_emision_basica(
    resultado: dict | None,
    *,
    exigir_fuente: bool = False,
) -> dict:
    """Evita que un proceso con código 0 pero sin resultado se tome como OK."""
    if not isinstance(resultado, dict):
        raise ErrorControlFuente(
            "El extractor terminó sin emitir un resultado JSON válido"
        )
    if exigir_fuente and not resultado.get("fuente"):
        raise ErrorControlFuente("El resultado no identifica la fuente")
    tiene_periodo = bool(
        resultado.get("periodo")
        or resultado.get("periodos")
        or (
            resultado.get("anio") is not None
            and resultado.get("mes_numero") is not None
        )
        or resultado.get("series")
    )
    if not tiene_periodo:
        raise ErrorControlFuente("El resultado no informa ningún período")
    return resultado


def validar_productivo(resultado: dict | None, *, fuente: str) -> dict:
    """Contrato estructural para las dos familias productivas."""
    validar_emision_basica(resultado, exigir_fuente=True)
    assert resultado is not None
    if fuente == "comercializacion":
        valores = resultado.get("valores")
        if not isinstance(valores, dict) or len(valores) != 8:
            cantidad = len(valores) if isinstance(valores, dict) else 0
            raise ErrorControlFuente(
                "Posible cambio de formato en comercialización: "
                f"se esperaban 8 cotizaciones y se extrajeron {cantidad}"
            )
        invalidos = [k for k, v in valores.items() if not _valor_valido(v)]
        if invalidos:
            raise ErrorControlFuente(
                "Cotizaciones vacías o inválidas: " + ", ".join(invalidos)
            )
    elif fuente == "pecuaria":
        registros = resultado.get("registros")
        if not isinstance(registros, list) or not registros:
            raise ErrorControlFuente("La extracción pecuaria no entregó registros")
        cantidad = sum(
            len([k for k in registro if k != "periodo"])
            for registro in registros
            if isinstance(registro, dict)
        )
        if cantidad < 10:
            raise ErrorControlFuente(
                "Posible cambio de formato en actividad pecuaria: "
                f"solo se extrajeron {cantidad} indicadores"
            )
        for registro in registros:
            if not isinstance(registro, dict) or not registro.get("periodo"):
                raise ErrorControlFuente("Hay un registro pecuario sin período")
            invalidos = [
                k for k, v in registro.items()
                if k != "periodo" and not _valor_valido(v)
            ]
            if invalidos:
                raise ErrorControlFuente(
                    "Indicadores pecuarios vacíos o inválidos: "
                    + ", ".join(invalidos)
                )
    return resultado


def _periodo_iso(valor) -> str:
    texto = str(valor or "").strip()
    try:
        date.fromisoformat(texto)
    except (TypeError, ValueError) as error:
        raise ErrorControlFuente(f"Período ausente o inválido: {texto!r}") from error
    return texto


def _valor_valido(valor) -> bool:
    if valor is None or isinstance(valor, bool):
        return False
    if isinstance(valor, (int, float)):
        return math.isfinite(float(valor))
    if isinstance(valor, str):
        return bool(valor.strip())
    return False


def validar_resultado(
    resultado: dict | None,
    *,
    fuente: str,
    minimo_valores: int = 0,
    cantidad_exacta: int | None = None,
    titulos_obligatorios: tuple[str, ...] = (),
) -> dict:
    """Valida el contrato mínimo antes de permitir la carga al Sheet."""
    if not isinstance(resultado, dict):
        raise ErrorControlFuente(
            "El extractor terminó sin emitir RESULTADO_FUENTE válido"
        )

    _periodo_iso(resultado.get("periodo"))
    valores = resultado.get("valores")
    if not isinstance(valores, dict):
        # Compatibilidad con el extractor histórico de supermercados.
        valores = resultado.get("valores_originales")
    if not isinstance(valores, dict):
        raise ErrorControlFuente("El resultado no contiene un diccionario de valores")

    invalidos = [titulo for titulo, valor in valores.items() if not _valor_valido(valor)]
    if invalidos:
        muestra = ", ".join(map(str, invalidos[:5]))
        raise ErrorControlFuente(f"Hay valores vacíos o inválidos: {muestra}")

    cantidad = len(valores)
    if cantidad_exacta is not None and cantidad != cantidad_exacta:
        raise ErrorControlFuente(
            f"Posible cambio de formato en {fuente}: se esperaban "
            f"{cantidad_exacta} valores y se extrajeron {cantidad}"
        )
    if cantidad < minimo_valores:
        raise ErrorControlFuente(
            f"Posible cambio de formato en {fuente}: se esperaban al menos "
            f"{minimo_valores} valores y se extrajeron {cantidad}"
        )

    faltantes = [titulo for titulo in titulos_obligatorios if titulo not in valores]
    if faltantes:
        raise ErrorControlFuente(
            "Faltan campos obligatorios: " + ", ".join(faltantes)
        )

    return valores
