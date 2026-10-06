"""Completa fórmulas conocidas por título y recalcula el libro con Excel.

Se ejecuta una sola vez después de todas las fuentes. Nunca modifica las
columnas de datos originales. Para cada grupo toma como plantilla una fórmula
anterior del mismo mes, traduce sus referencias a la fila actual y conserva
las fórmulas que ya existen.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
from comun.columnas_por_titulo import resolver_columnas


RAIZ = Path(__file__).resolve().parents[2]
BASE_PREDETERMINADA = RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA_BASE = "BASE DE DATOS"


@dataclass(frozen=True)
class GrupoFormulas:
    fuente: str
    nombre: str
    activadores: tuple[str, ...]
    formulas: tuple[str, ...]


GRUPOS = (
    GrupoFormulas(
        "atp",
        "ATP - recaudación tributaria",
        (
            "Recaudación tributaria F",
        ),
        (
            "Recaudación tributaria acumulada F",
            "Variación intermensual recaudación tributaria F Nominal",
            "Variación interanual recaudación tributaria F Nominal",
            "Variación intermensual recaudación de IIBB F",
            "Variación interanual recaudación de IIBB F",
            "IIBB/Recaudacion tributaria(%) F",
            "Var% i.m. Recaudación Imp.\nde Sellos F",
            "Var% i.a. Recaudación Imp.\nde Sellos F",
            "Var% i.m. Recaudacion tasas Retrib.\nde Servicios F",
            "Var% i.a. Recaudacion tasas Retrib.\nde Servicios F",
            "Var% i.m. Recaudacion Reg. de Facil.\nde Pago        ",
            "Var% i.a. Recaudacion Reg. de Facil.\nde Pago        ",
            "Var% i.m. Recaudacion Inmob. Rural",
            "Var% i.a.  Recaudacion Inmob. Rural",
            "Var% i.m.  Recaudacion Imp. a los\nJuegos de Azar",
            "Var% i.a.  Recaudacion Imp. a los\nJuegos de Azar",
            "Var% i.m. Recaudacion Transf. Mej.\ne/Tier.\nFisc.",
            "Var% i.a. Recaudacion Transf. Mej.\ne/Tier.\nFisc.",
            "Total recaudacion \"otros impuestos y tasas\"",
        ),
    ),
    GrupoFormulas(
        "dnrpa",
        "DNRPA - automotores y motos",
        (
            "Patentamiento de automóviles F",
            "Patentamiento de automóviles N",
            "Transferencias de automoviles F",
            "Transferencias de automoviles N",
            "Patentamiento de Motos F",
            "Patentamiento de Motos N",
            "Transferencias de Motos F",
            "Transferencias de Motos N",
        ),
        (
            "Patentamiento de automóviles - Acumulado F",
            "Patentamiento de automóviles - Var % intermensual F",
            "Patentamiento de automóviles - Var % interanual F",
            "Transferencias de automoviles acumulada F",
            "Transferencias de automóviles - Var % intermensual F",
            "Transferencias de automóviles - Var % interanual F",
            "Patentamiento de automóviles - Acumulado N",
            "Patentamiento de automóviles - Var % intermensual N",
            "Patentamiento de automóviles - Var % interanual N",
            "Transferencias de automoviles acumulada N",
            "Transferencias de automóviles - Var % intermensual N",
            "Transferencias de automóviles - Var % interanual N",
            "Patentamiento de Motos - Acumulado F",
            "Patentamiento de Motos- Var % intermensual F",
            "Patentamiento de Motos- Var % interanual F",
            "Transferencias de Motos - Acumulado F",
            "Transferencias de Motos- Var % intermensual F",
            "Trasnferencias de Motos- Var % interanual F",
            "Patentamiento de Motos - Acumulado N",
            "Patentamiento de Motos- Var % intermensual N",
            "Patentamiento de Motos- Var % interanual N",
            "Transferencias de Motos - Acumulado N",
            "Transferencias de Motos- Var % intermensual N",
            "Trasnferencias de Motos- Var % interanual N",
        ),
    ),
    GrupoFormulas(
        "ieric",
        "IERIC - empresas constructoras",
        (
            "Empresas constructoras/contratistas F",
            "Empresas constructoras/contratistas N",
        ),
        (
            "Var. IA total de empresas constructoras F",
            "Var. IA total de empresas constructoras N",
        ),
    ),
    GrupoFormulas(
        "iaraf",
        "IARAF - acumulados",
        (
            "Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ F",
            "Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ N",
        ),
        (
            "Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Acumulado anual. En millones de $ F",
            "Transferencias automáticas por coparticipación, leyes compl. y compensac. Acumulado anual. En millones de $ N",
        ),
    ),
    GrupoFormulas(
        "canasta",
        "CBA y CBT",
        (
            "CBA. Valores mensuales por adulto equivalente. Línea de indigencia. (Valor en $)",
            "CBT. Valores mensuales por adulto equivalente. Línea de pobreza (Valor en $)",
        ),
        (
            "Acumulado CBA adulto equivalente año en curso",
            "Acumulado CBT adulto equivalente año en curso",
            "VAR% mensual CBA hogar 2",
            "VAR% anual CBA hogar 2",
            "VAR% mensual CBA hogar 3",
            "VAR% anual CBA hogar 3",
            "VAR% mensual CBT hogar 2",
            "VAR% anual CBT hogar 2",
            "VAR% mensual CBT hogar 3",
            "VAR% anual CBT hogar 3",
        ),
    ),
    GrupoFormulas(
        "sipa",
        "SIPA - trabajo registrado",
        (
            "Cant. Personas con trabajo registrado. Total país. Serie Original (en miles)",
            "Cant. Personas con trabajo registrado. Total país. Serie desestacionalizada (en miles)",
            "Cant. asalariados registrados del sector privado. Sin estacionalidad. En miles. N",
            "Cant. asalariados registrados del sector privado (en miles), según rama de actividad S/E. Agricultura, ganaderÍa, caza y silvicultura N",
        ),
        (
            "Var% i.m. Personas con trabajo registrado. Total país. Serie Original",
            "Var% i.a. Personas con trabajo registrado. Total país. Serie Original",
            "Var. absoluta i.m. Personas con trabajo registrado. Total país. Serie Original",
            "Var. absoluta i.a. Personas con trabajo registrado. Total país. Serie Original",
            "Var% i.m. Personas con trabajo registrado. Total país. Serie desestacionalizada",
            "Var% i.a. Personas con trabajo registrado. Total país. Serie desestacionalizada",
            "Var. absoluta i.m. Personas con trabajo registrado. Total país. Serie desestacionalizada",
            "Var. absoluta i.a. Personas con trabajo registrado. Total país. Serie desestacionalizada ",
            "Var% i.m. asalariados registrados del sector privado N",
            "Var. absoluta i.m. asalariados registrados del sector privado N",
            "Var % i.a. asalariados registrados del sector privado N",
            "Var. absoluta i.a. asalariados registrados del sector privado (en miles) N",
            "Var% i.m.  asalariados registrados del sector privado, según rama de actividad S/E. Agricultura, ganaderÍa, caza y silvicultura N",
            "Var% i.a.  asalariados registrados del sector privado, según rama de actividad S/E. Agricultura, ganaderÍa, caza y silvicultura N",
        ),
    ),
)


def configurar_salida() -> None:
    for flujo in (sys.stdout, sys.stderr):
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def es_dato(valor: object) -> bool:
    return isinstance(valor, (int, float)) and not isinstance(valor, bool) and float(valor) != 0.0


def mes_celda(valor: object) -> int | None:
    if isinstance(valor, (date, datetime)):
        return valor.month
    return None


def plantilla_formula(ws, fila_destino: int, columna: int) -> tuple[int, str]:
    mes = mes_celda(ws.cell(fila_destino, 1).value)
    candidatas = []
    for fila in range(2, fila_destino):
        valor = ws.cell(fila, columna).value
        if isinstance(valor, str) and valor.startswith("="):
            candidatas.append((fila, valor))
    if not candidatas:
        raise ValueError(
            f"No existe una fórmula anterior para {ws.cell(1, columna).value!r}."
        )
    mismo_mes = [item for item in candidatas if mes_celda(ws.cell(item[0], 1).value) == mes]
    return (mismo_mes or candidatas)[-1]


def completar_formulas(ruta: Path, fuentes: set[str] | None = None) -> dict:
    libro = load_workbook(ruta, data_only=False, keep_links=True)
    if HOJA_BASE not in libro.sheetnames:
        raise ValueError(f"No existe la hoja {HOJA_BASE!r}.")
    ws = libro[HOJA_BASE]
    cambios = []
    grupos_procesados = []

    for grupo in GRUPOS:
        # SIPA ya entrega sus 70 valores y cálculos desde el anexo oficial.
        # No copiar fórmulas hacia filas futuras de la base.
        if grupo.fuente == "sipa":
            continue
        if fuentes is not None and grupo.fuente not in fuentes:
            continue
        pedidos = {
            **{f"a{n}": titulo for n, titulo in enumerate(grupo.activadores)},
            **{f"f{n}": titulo for n, titulo in enumerate(grupo.formulas)},
        }
        columnas = resolver_columnas(ws, pedidos)
        activadores = [columnas[f"a{n}"] for n in range(len(grupo.activadores))]
        filas_activas = [
            fila
            for fila in range(2, ws.max_row + 1)
            if all(es_dato(ws.cell(fila, columna).value) for columna in activadores)
        ]
        if not filas_activas:
            continue
        fila = max(filas_activas)
        grupos_procesados.append({"grupo": grupo.nombre, "fila": fila})

        for n, titulo in enumerate(grupo.formulas):
            columna = columnas[f"f{n}"]
            destino = ws.cell(fila, columna)
            if isinstance(destino.value, str) and destino.value.startswith("="):
                continue
            fila_origen, formula = plantilla_formula(ws, fila, columna)
            origen = ws.cell(fila_origen, columna)
            traducida = Translator(formula, origin=origen.coordinate).translate_formula(destino.coordinate)
            if grupo.fuente == "atp" and "Transf. Mej." in titulo:
                traducida = f"=IFERROR({traducida[1:]},0)"
            anterior = destino.value
            destino.value = traducida
            cambios.append({
                "grupo": grupo.nombre,
                "celda": destino.coordinate,
                "titulo": titulo,
                "antes": anterior,
                "despues": traducida,
                "plantilla": origen.coordinate,
            })

    respaldo = None
    if cambios:
        carpeta = ruta.parent / "respaldos"
        carpeta.mkdir(parents=True, exist_ok=True)
        respaldo = carpeta / f"antes_formulas_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        shutil.copy2(ruta, respaldo)
        libro.calculation.fullCalcOnLoad = True
        libro.calculation.forceFullCalc = True
        libro.calculation.calcMode = "auto"
        libro.save(ruta)

    return {
        "archivo_excel": str(ruta.resolve()),
        "formulas_completadas": len(cambios),
        "cambios": cambios,
        "grupos_procesados": grupos_procesados,
        "respaldo": str(respaldo.resolve()) if respaldo else None,
    }


def recalcular_con_excel(ruta: Path) -> dict:
    if sys.platform != "win32":
        # Servidor Linux: no hay Microsoft Excel. No hace falta recalcular aquí:
        # el libro queda marcado para recalcular al abrirse (fullCalcOnLoad),
        # el envío a Google Sheets compara fórmulas (no valores cacheados) y
        # Google Sheets calcula sus propias fórmulas.
        return {
            "recalculado": False,
            "detalle": "Sin Excel (servidor no Windows): el libro se recalcula al abrirse; Google Sheets calcula sus fórmulas.",
        }
    segura = str(ruta.resolve()).replace("'", "''")
    script = rf"""
$ErrorActionPreference = 'Stop'
$excel = $null
$libro = $null
try {{
    $ruta = '{segura}'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $libro = $excel.Workbooks.Open($ruta, 0, $false)
    if ($libro.ReadOnly) {{ throw 'Excel abrió el archivo en modo solo lectura.' }}
    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()
    $libro.Save()
    Write-Output 'RECALCULO_COM=OK'
}}
finally {{
    if ($libro -ne $null) {{ $libro.Close($false) }}
    if ($excel -ne $null) {{ $excel.Quit() }}
    if ($libro -ne $null) {{ [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($libro) }}
    if ($excel -ne $null) {{ [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""
    proceso = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proceso.returncode != 0:
        detalle = (proceso.stderr or proceso.stdout).strip()
        raise RuntimeError(f"Excel no pudo recalcular y guardar el libro: {detalle}")
    return {"recalculado": True, "detalle": proceso.stdout.strip()}


def main() -> int:
    configurar_salida()
    parser = argparse.ArgumentParser(description="Completa fórmulas y recalcula el Excel base.")
    parser.add_argument("--base", type=Path, default=BASE_PREDETERMINADA)
    parser.add_argument(
        "--fuentes",
        help="Fuentes separadas por comas; limita las fórmulas a esos indicadores.",
    )
    parser.add_argument("--sin-recalcular", action="store_true", help="Solo para pruebas: no abre Excel.")
    args = parser.parse_args()
    try:
        ruta = args.base.expanduser().resolve()
        if not ruta.is_file():
            raise FileNotFoundError(f"No existe la base: {ruta}")
        fuentes = (
            {item.strip().casefold() for item in args.fuentes.split(",") if item.strip()}
            if args.fuentes
            else None
        )
        resultado = completar_formulas(ruta, fuentes)
        resultado["recalculo"] = (
            {"recalculado": False, "detalle": "Omitido por --sin-recalcular"}
            if args.sin_recalcular
            else recalcular_con_excel(ruta)
        )
        print("FINALIZACION_EXCEL=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
        print("ESTADO=OK")
        return 0
    except Exception as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
