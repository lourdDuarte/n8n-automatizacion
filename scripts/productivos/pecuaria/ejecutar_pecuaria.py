"""Ejecutor de Actividad Pecuaria integrado con ejecutar_productivos.py."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

from actualizar_pecuaria import actualizar
from descargar_pecuaria import obtener
from extraer_pecuaria import (combinar_por_periodo, extraer_aves_excel,
                              extraer_faena_bovina, extraer_indicadores_bovinos,
                              extraer_huevos_pdf, extraer_porcinos_excel,
                              extraer_stock_bovino, extraer_stock_porcino,
                              faena_nacion_tablero)


# Punto de arranque tomado del informe oficial enero-julio 2026. A partir de
# aquí el control se mantiene solo con cada nuevo informe oficial descargado.
BASE_ACUMULADA_FORMOSA = {"2026-07": 20381.0}


def faena_formosa_mensual(registro: dict, ruta_estado: Path) -> dict:
    """Convierte el acumulado oficial en valor mensual sin leer el Excel base."""
    periodo = registro["periodo"]
    clave = periodo.strftime("%Y-%m")
    acumulada = float(registro.pop("faena_bovina_formosa_acumulada"))
    estado = {}
    if ruta_estado.exists():
        try:
            estado = json.loads(ruta_estado.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            estado = {}

    if estado.get("periodo") == clave and estado.get("mensual") is not None:
        mensual = float(estado["mensual"])
    else:
        anterior = date(
            periodo.year - (periodo.month == 1),
            12 if periodo.month == 1 else periodo.month - 1,
            1,
        )
        clave_anterior = anterior.strftime("%Y-%m")
        if estado.get("periodo") == clave_anterior:
            acumulada_anterior = float(estado["acumulada"])
        elif clave_anterior in BASE_ACUMULADA_FORMOSA:
            acumulada_anterior = BASE_ACUMULADA_FORMOSA[clave_anterior]
        elif periodo.month == 1:
            acumulada_anterior = 0.0
        else:
            raise ValueError(
                f"No se puede obtener la faena mensual de Formosa {clave}: "
                f"el último informe controlado debe ser {clave_anterior}"
            )
        mensual = acumulada - acumulada_anterior
        if mensual < 0:
            raise ValueError(f"El acumulado oficial de Formosa disminuyó para {clave}")

    registro["faena_bovina_formosa"] = round(mensual, 2)
    ruta_estado.parent.mkdir(parents=True, exist_ok=True)
    ruta_estado.write_text(json.dumps({
        "periodo": clave,
        "acumulada": acumulada,
        "mensual": round(mensual, 2),
        "actualizado_en": datetime.now().astimezone().isoformat(timespec="seconds"),
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    return registro


def raiz_proyecto() -> Path:
    return Path(__file__).resolve().parents[3]


def main() -> int:
    raiz = raiz_proyecto()
    ap = argparse.ArgumentParser(description="Actualiza Actividad Pecuaria")
    ap.add_argument("--base", type=Path, default=raiz / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx")
    ap.add_argument("--solo-extraer", action="store_true")
    args = ap.parse_args()
    try:
        pubs = obtener(raiz / "fuentes" / "productivos" / "pecuaria")
        series = []
        series.append([extraer_stock_bovino(pubs["stock_bovino"].contenido)])
        series.append([extraer_stock_porcino(pubs["stock_porcino"].contenido)])
        faena_bovina = extraer_faena_bovina(pubs["faena_bovina"].contenido, pubs["faena_bovina_excel"].contenido)
        faena_bovina = faena_formosa_mensual(
            faena_bovina,
            raiz / "scripts" / "reportes" / "productivos" / "pecuaria" / "estado_faena_formosa.json",
        )
        # Primero la planilla de indicadores y después el informe mensual de faena:
        # la faena nacional que vale es el "Total de faena bovina" del informe mensual
        # (cabezas enteras), no la estimación de la planilla de indicadores.
        tablero = faena_nacion_tablero()
        if tablero and tablero["periodo"] == faena_bovina["periodo"]:
            faena_bovina["faena_bovina_nacion"] = tablero["faena_bovina_nacion"]
        series.append(extraer_indicadores_bovinos(pubs["indicadores_bovinos"].contenido))
        series.append([faena_bovina])
        series.append(extraer_huevos_pdf(pubs["aves_huevos"].contenido))
        series.append(extraer_aves_excel(pubs["aves_oferta_demanda"].contenido))
        series.append(extraer_porcinos_excel(pubs["indicadores_porcinos"].contenido))
        series = [[max(s, key=lambda x: x["periodo"])] for s in series if s]
        registros = combinar_por_periodo(*series)
        if not registros:
            raise ValueError("Las publicaciones no entregaron registros pecuarios")
        resultado = {
            "fuente": "pecuaria", "nombre": "Actividad pecuaria",
            "periodo": max(x["periodo"] for x in registros).isoformat(),
            "consultado_en": datetime.now().astimezone().isoformat(timespec="seconds"),
            "registros": [{**x, "periodo": x["periodo"].isoformat()} for x in registros],
            "urls": {k: v.url for k, v in pubs.items()},
        }
        if not args.solo_extraer:
            resultado["actualizacion"] = actualizar(args.base, registros)
        carpeta = raiz / "scripts" / "reportes" / "productivos" / "pecuaria"; carpeta.mkdir(parents=True, exist_ok=True)
        reporte = carpeta / f"pecuaria_{datetime.now():%Y%m%d_%H%M%S}.json"
        reporte.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
        resultado["reporte"] = str(reporte)
        print("RESULTADO_PECUARIA=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
        print("ESTADO=OK")
        return 0
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
