"""Actualiza los nueve indicadores CAMMESA y entrega la carga completa a n8n."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import json
import os
import sys
import tempfile

from openpyxl import load_workbook


SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
from comun.columnas_por_titulo import resolver_columnas


PROYECTO = Path(__file__).resolve().parents[2]
ARCHIVO_BASE = PROYECTO / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
REPORTES = PROYECTO / "scripts" / "reportes" / "cammesa"
HOJA = "BASE DE DATOS"

MAPEO = {
    "demanda_comercial": "Demanda (MWh) Categoría Tarifa Comercial F",
    "demanda_industrial_comercial_grande": "Demanda (MWh) Categoría Tarifa Industrial / Comercial Grande F",
    "demanda_residencial": "Demanda (MWh) Categoría Tarifa Residencial F",
    "variacion_mensual_comercial": "Var % mensual demanda (MWh) Categoría Tarifa Comercial F",
    "variacion_anual_comercial": "Var % anual demanda (MWh) Categoría Tarifa Comercial F",
    "variacion_mensual_industrial_comercial_grande": "Var % mensual demanda (MWh) Categoría Tarifa Industrial / Comercial Grande F",
    # Se conservan los espacios dobles que tienen estos dos encabezados
    # en la base original del Observatorio.
    "variacion_anual_industrial_comercial_grande": "Var % anual demanda  (MWh) Categoría Tarifa Industrial / Comercial Grande F",
    "variacion_mensual_residencial": "Var % mensual demanda  (MWh)  Categoría Tarifa Residencial F",
    "variacion_anual_residencial": "Var % anual demanda (MWh) Categoría Tarifa Residencial F",
}

CALCULOS = {
    "comercial": "demanda_comercial",
    "industrial_comercial_grande": "demanda_industrial_comercial_grande",
    "residencial": "demanda_residencial",
}


def configurar_salida() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def control_reciente() -> Path:
    controles = [p for p in REPORTES.glob("cammesa_????_??.json") if p.is_file()]
    if not controles:
        raise FileNotFoundError(f"No hay controles CAMMESA en {REPORTES}.")
    return max(controles, key=lambda p: p.stem[-7:])


def cargar_control(ruta: Path) -> tuple[str, dict]:
    contenido = json.loads(ruta.read_text(encoding="utf-8"))
    periodo = str(contenido.get("periodo", ""))
    datos = contenido.get("datos")
    if not isinstance(datos, dict) or set(datos) != set(MAPEO):
        faltan = sorted(set(MAPEO) - set(datos or {}))
        sobran = sorted(set(datos or {}) - set(MAPEO))
        raise RuntimeError(f"Control CAMMESA inválido. Faltan={faltan}; sobran={sobran}")
    return periodo, datos


def periodo_celda(valor) -> str | None:
    if isinstance(valor, (date, datetime)):
        return f"{valor.year:04d}-{valor.month:02d}"
    texto = str(valor or "").strip().casefold()
    meses = {m: i + 1 for i, m in enumerate(("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"))}
    partes = texto.split()
    if len(partes) == 2 and partes[0][:3] in meses:
        try:
            anio = int(partes[1]); anio = 2000 + anio if anio < 100 else anio
            return f"{anio:04d}-{meses[partes[0][:3]]:02d}"
        except ValueError:
            return None
    return texto if len(texto) >= 7 else None


def localizar_fila(hoja, periodo: str) -> int:
    columna_periodo = resolver_columnas(hoja, {"periodo": "PERIODO"})["periodo"]
    filas = [
        fila for fila in range(2, hoja.max_row + 1)
        if periodo_celda(hoja.cell(fila, columna_periodo).value) == periodo
    ]
    if len(filas) != 1:
        raise RuntimeError(f"Se esperaba una fila para {periodo} y se encontraron {len(filas)}.")
    return filas[0]


def periodo_anterior(periodo: str) -> str:
    anio, mes = map(int, periodo.split("-"))
    return f"{anio - 1:04d}-12" if mes == 1 else f"{anio:04d}-{mes - 1:02d}"


def periodo_interanual(periodo: str) -> str:
    anio, mes = map(int, periodo.split("-"))
    return f"{anio - 1:04d}-{mes:02d}"


def calcular_variaciones_desde_base(hoja, columnas: dict, periodo: str, datos: dict) -> None:
    """Calcula variaciones contra la serie ya publicada por el Observatorio."""
    fila_anterior = localizar_fila(hoja, periodo_anterior(periodo))
    fila_interanual = localizar_fila(hoja, periodo_interanual(periodo))
    for sufijo, campo_nivel in CALCULOS.items():
        columna = columnas[campo_nivel]
        base_mensual = numero(hoja.cell(fila_anterior, columna).value)
        base_interanual = numero(hoja.cell(fila_interanual, columna).value)
        actual = float(datos[campo_nivel])
        if not isinstance(base_mensual, (int, float)) or float(base_mensual) == 0:
            raise RuntimeError(f"Falta la base mensual para calcular {campo_nivel}.")
        if not isinstance(base_interanual, (int, float)) or float(base_interanual) == 0:
            raise RuntimeError(f"Falta la base interanual para calcular {campo_nivel}.")
        datos[f"variacion_mensual_{sufijo}"] = round((actual / float(base_mensual) - 1) * 100, 2)
        datos[f"variacion_anual_{sufijo}"] = round((actual / float(base_interanual) - 1) * 100, 2)


def numero(valor):
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor or "").strip().replace(" ", "")
    if not texto:
        return valor
    try:
        if "," in texto and "." in texto:
            texto = texto.replace(",", "")
        elif "," in texto:
            texto = texto.replace(",", ".")
        return float(texto)
    except ValueError:
        return valor


def iguales(anterior, nuevo) -> bool:
    a, n = numero(anterior), numero(nuevo)
    if isinstance(a, (int, float)) and isinstance(n, (int, float)):
        return abs(float(a) - float(n)) < 0.005
    return a == n


def main() -> int:
    configurar_salida()
    periodo, datos = cargar_control(control_reciente())
    if not ARCHIVO_BASE.is_file():
        raise FileNotFoundError(f"No existe el Excel: {ARCHIVO_BASE}")

    libro = load_workbook(ARCHIVO_BASE, data_only=False, keep_links=True)
    temporal = None
    try:
        hoja = libro[HOJA]
        columnas = resolver_columnas(hoja, MAPEO)
        fila = localizar_fila(hoja, periodo)
        calcular_variaciones_desde_base(hoja, columnas, periodo, datos)
        detalle_completo = []
        cambios = []
        for campo, columna in columnas.items():
            celda = hoja.cell(fila, columna)
            if isinstance(celda.value, str) and celda.value.startswith("="):
                raise RuntimeError(f"No se modificó la base porque {celda.coordinate} contiene una fórmula.")
            nuevo = datos[campo]
            detalle_completo.append({
                "celda": celda.coordinate,
                "titulo": MAPEO[campo],
                "antes": celda.value,
                "despues": nuevo,
            })
            if not iguales(celda.value, nuevo):
                cambios.append((celda, celda.value, nuevo, MAPEO[campo]))

        for celda, _, nuevo, _ in cambios:
            celda.value = nuevo

        if cambios:
            ARCHIVO_BASE.parent.mkdir(parents=True, exist_ok=True)
            descriptor, nombre = tempfile.mkstemp(suffix=".xlsx", dir=ARCHIVO_BASE.parent)
            os.close(descriptor)
            temporal = Path(nombre)
            libro.save(temporal)
    finally:
        libro.close()

    if temporal is not None:
        try:
            os.replace(temporal, ARCHIVO_BASE)
        except PermissionError as error:
            raise PermissionError("Cerrá completamente el Excel local y volvé a ejecutar.") from error
        finally:
            temporal.unlink(missing_ok=True)

    resultado = {
        "fuente": "cammesa",
        "nombre": "CAMMESA - Demanda eléctrica Formosa",
        "periodo": periodo,
        "estado": "OK",
        "actualizacion": {
            "archivo_excel": str(ARCHIVO_BASE),
            "hoja": HOJA,
            "periodo": periodo,
            "fila": fila,
            "celdas_actualizadas": len(cambios),
            "hubo_cambios": bool(cambios),
            "cambios": [
                {"titulo": titulo, "antes": anterior, "despues": nuevo}
                for _, anterior, nuevo, titulo in cambios
            ],
            "detalle_completo": detalle_completo,
        },
    }
    print(f"Período: {periodo}")
    print(f"Celdas actualizadas: {len(cambios)}")
    print("RESULTADO_CAMMESA=" + json.dumps(resultado, ensure_ascii=False, separators=(",", ":")))
    print("RESULTADO: OK")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        configurar_salida()
        print(f"ERROR: {type(error).__name__}: {error}", file=sys.stderr)
        raise SystemExit(1)
