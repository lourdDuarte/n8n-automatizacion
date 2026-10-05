"""Actualiza los 55 encabezados IPI con el último control validado."""

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


# ----------------------------------------------------------------------
# CONFIGURACIÓN
# ----------------------------------------------------------------------

CARPETA_PROYECTO = Path(__file__).resolve().parents[2]

ARCHIVO_BASE = (
    CARPETA_PROYECTO
    / "base"
    / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
)

CARPETA_REPORTES = (
    CARPETA_PROYECTO
    / "scripts"
    / "reportes"
    / "ipi"
)

HOJA_BASE = "BASE DE DATOS"

MAPEO_COLUMNAS = {
    "serie_original": "IPI INDICE SERIE ORIGINAL 2004=100",
    "serie_desestacionalizada": "IPI INDICE SERIE DESESTACIONALIZADA 2004=100",
    "serie_tendencia_ciclo": "IPI INDICE SERIE TENDENCIA - CICLO 2004=100",
    "variacion_acumulada": "IPI SERIE ORIG Acumulado del año respecto al acumulado de igual periodo del año anterior N",
    "variacion_interanual": "IPI Respecto a igual mes del año anterior (Serie original) N",
    "variacion_mensual_desestacionalizada": "IPI Variación porcentual respecto al mes anterior (serie desestacionalizada) N",
    "variacion_mensual_tendencia_ciclo": "IPI Variación porcentual respecto al mes anterior (serie tendencia - ciclo) N",
}

DIVISIONES_BASE = {
    "alimentos_bebidas": "Alimentos y bebidas",
    "productos_tabaco": "Productos de tabaco",
    "productos_textiles": "Productos textiles",
    "prendas_cuero_calzado": "Prendas de vestir, cuero y calzado",
    "madera_papel_edicion": "Madera, papel, edición e impresión",
    "refinacion_petroleo": "Refinación del petróleo, coque y combustible nuclear",
    "sustancias_quimicos": "Sustancias y productos químicos",
    "caucho_plastico": "Productos de caucho y plástico",
    "minerales_no_metalicos": "Productos minerales no metálicos",
    "metalicas_basicas": "Industrias metálicas básicas",
    "productos_metal": "Productos de metal",
    "maquinaria_equipo": "Maquinaria y equipo",
    "otros_equipos": "Otros equipos, aparatos e instrumentos",
    "vehiculos_automotores": "Vehículos automotores, carrocerías, remolques y autopartes",
    "otro_transporte": "Otro equipo de transporte",
    "muebles_otras": "Muebles y colchones, y otras industrias manufactureras",
}

for clave, titulo in DIVISIONES_BASE.items():
    MAPEO_COLUMNAS[f"ia_{clave}"] = f"Var% i.a. IPI divisiones. {titulo} (serie orig)"
    MAPEO_COLUMNAS[f"im_{clave}"] = f"Var% i.m. IPI divisiones. {titulo} (serie desest)"
    MAPEO_COLUMNAS[f"acum_{clave}"] = f"Var% i.a. IPI divisiones. {titulo} (acum)"

# Dos títulos de la base histórica no contienen el espacio después del punto.
MAPEO_COLUMNAS["ia_sustancias_quimicos"] = "Var% i.a. IPI divisiones.Sustancias y productos químicos (serie orig)"
MAPEO_COLUMNAS["im_sustancias_quimicos"] = "Var% i.m. IPI divisiones.Sustancias y productos químicos (serie desest)"
MAPEO_COLUMNAS["acum_sustancias_quimicos"] = "Var% i.a. IPI divisiones.Sustancias y productos químicos (acum)"


# ----------------------------------------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------------------------------------

def configurar_salida() -> None:
    """Configura stdout y stderr en UTF-8 para CMD y n8n."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def periodo_del_nombre(archivo: Path) -> tuple[int, int]:
    """Obtiene año y mes desde un nombre ipi_AAAA_MM.json."""
    partes = archivo.stem.split("_")

    if len(partes) != 3 or partes[0].lower() != "ipi":
        raise ValueError(
            f"Nombre de control IPI inválido: {archivo.name}"
        )

    try:
        anio = int(partes[1])
        mes = int(partes[2])
    except ValueError as error:
        raise ValueError(
            f"Nombre de control IPI inválido: {archivo.name}"
        ) from error

    if anio < 2000 or not 1 <= mes <= 12:
        raise ValueError(
            f"Período inválido en el nombre: {archivo.name}"
        )

    return anio, mes


def buscar_ultimo_control() -> Path:
    """Busca el control IPI más reciente en la carpeta correspondiente."""
    if not CARPETA_REPORTES.exists():
        raise FileNotFoundError(
            f"No existe la carpeta de reportes del IPI: {CARPETA_REPORTES}"
        )

    controles = [
        archivo
        for archivo in CARPETA_REPORTES.glob("ipi_????_??.json")
        if archivo.is_file()
    ]

    if not controles:
        raise FileNotFoundError(
            "No se encontró ningún archivo ipi_AAAA_MM.json en: "
            f"{CARPETA_REPORTES}"
        )

    control = max(controles, key=periodo_del_nombre)

    print(f"Control IPI localizado: {control}")

    return control


def cargar_control(ruta_control: Path) -> tuple[str, dict[str, int | float]]:
    """Carga el JSON y valida el período y los 55 datos obligatorios."""
    try:
        contenido = ruta_control.read_text(encoding="utf-8")
        resultado = json.loads(contenido)
    except OSError as error:
        raise OSError(
            f"No se pudo leer el control: {ruta_control}"
        ) from error
    except json.JSONDecodeError as error:
        raise ValueError(
            f"El control no contiene un JSON válido: {ruta_control}"
        ) from error

    if not isinstance(resultado, dict):
        raise ValueError(
            f"El contenido de {ruta_control.name} debe ser un objeto JSON."
        )

    periodo = resultado.get("periodo")
    datos = resultado.get("datos")

    if not isinstance(periodo, str):
        raise ValueError(
            f"El control {ruta_control.name} no contiene un período válido."
        )

    partes = periodo.split("-")

    if len(partes) != 2:
        raise ValueError(
            f"El período debe tener formato AAAA-MM: {periodo!r}"
        )

    try:
        anio = int(partes[0])
        mes = int(partes[1])
    except ValueError as error:
        raise ValueError(
            f"El período debe tener formato AAAA-MM: {periodo!r}"
        ) from error

    if anio < 2000 or not 1 <= mes <= 12:
        raise ValueError(f"Período inválido: {periodo!r}")

    periodo_nombre = periodo_del_nombre(ruta_control)

    if periodo_nombre != (anio, mes):
        raise ValueError(
            "El período interno del JSON no coincide con su nombre: "
            f"{ruta_control.name} contiene {periodo!r}."
        )

    if not isinstance(datos, dict):
        raise ValueError(
            f"El control {ruta_control.name} no contiene el objeto 'datos'."
        )

    datos_validados: dict[str, int | float] = {}
    campos_invalidos = []

    for campo in MAPEO_COLUMNAS:
        valor = datos.get(campo)

        if (
            valor is None
            or isinstance(valor, bool)
            or not isinstance(valor, (int, float))
        ):
            campos_invalidos.append(campo)
        else:
            datos_validados[campo] = valor

    if campos_invalidos:
        raise ValueError(
            "Faltan datos numéricos obligatorios o son inválidos: "
            + ", ".join(campos_invalidos)
        )

    return periodo, datos_validados


def localizar_fila_periodo(hoja, periodo: str) -> int:
    """Busca el período del IPI en la columna A."""
    anio, mes = map(int, periodo.split("-"))

    for fila in range(2, hoja.max_row + 1):
        valor = hoja.cell(row=fila, column=1).value

        if (
            isinstance(valor, (datetime, date))
            and valor.year == anio
            and valor.month == mes
        ):
            return fila

        if isinstance(valor, str):
            texto = valor.strip()

            if texto in {
                f"{anio}-{mes:02d}",
                f"{mes:02d}/{anio}",
                f"{mes}/{anio}",
            }:
                return fila

    raise ValueError(
        f"No se encontró el período {periodo} en la columna A "
        f"de la hoja {HOJA_BASE!r}."
    )


def es_formula(valor) -> bool:
    """Indica si una celda contiene una fórmula."""
    return isinstance(valor, str) and valor.startswith("=")


def crear_archivo_temporal() -> Path:
    """Crea un archivo temporal junto al Excel original."""
    descriptor, nombre = tempfile.mkstemp(
        prefix="ipi_temporal_",
        suffix=".xlsx",
        dir=ARCHIVO_BASE.parent,
    )

    os.close(descriptor)

    return Path(nombre)


# ----------------------------------------------------------------------
# PROCESO PRINCIPAL
# ----------------------------------------------------------------------

def main() -> int:
    configurar_salida()

    print("=" * 72)
    print("ACTUALIZACIÓN DEL EXCEL — IPI MANUFACTURERO")
    print("=" * 72)
    print(f"Proyecto: {CARPETA_PROYECTO}")
    print(f"Excel: {ARCHIVO_BASE}")
    print(f"Reportes: {CARPETA_REPORTES}")
    print("-" * 72)

    if not CARPETA_PROYECTO.exists():
        raise FileNotFoundError(
            f"No existe la carpeta del proyecto: {CARPETA_PROYECTO}"
        )

    if not ARCHIVO_BASE.exists():
        raise FileNotFoundError(
            f"No existe el Excel: {ARCHIVO_BASE}"
        )

    ruta_control = buscar_ultimo_control()
    periodo, datos = cargar_control(ruta_control)

    print(f"Control seleccionado: {ruta_control.name}")
    print(f"Período del control: {periodo}")

    libro = load_workbook(
        ARCHIVO_BASE,
        data_only=False,
        keep_links=True,
    )

    archivo_temporal: Path | None = None
    cambios: list[dict] = []

    try:
        if HOJA_BASE not in libro.sheetnames:
            raise ValueError(
                f"No existe la hoja {HOJA_BASE!r} en el Excel."
            )

        hoja = libro[HOJA_BASE]
        columnas = resolver_columnas(hoja, MAPEO_COLUMNAS)
        fila = localizar_fila_periodo(hoja, periodo)

        print(f"Fila localizada: {fila}")
        print("-" * 72)

        # Primero valida las 55 celdas sin modificar ninguna.
        for columna in columnas.values():
            celda = hoja.cell(fila, columna)

            if es_formula(celda.value):
                raise ValueError(
                    "No se modificó ninguna celda porque "
                    f"{celda.coordinate} contiene una fórmula."
                )

        # Determina cuáles celdas necesitan actualización y prepara, además,
        # la carga completa que se enviará al Observatorio.
        detalle_completo: list[dict] = []
        for campo, columna in columnas.items():
            celda = hoja.cell(fila, columna)
            valor_anterior = celda.value
            valor_nuevo = datos[campo]

            detalle_completo.append(
                {
                    "celda": celda.coordinate,
                    "titulo": MAPEO_COLUMNAS[campo],
                    "antes": valor_anterior,
                    "despues": valor_nuevo,
                }
            )

            if valor_anterior == valor_nuevo:
                print(
                    f"{celda.coordinate}: "
                    f"{valor_anterior!r} -> {valor_nuevo!r} "
                    "[SIN CAMBIOS]"
                )
            else:
                cambios.append(
                    {
                        "celda": celda,
                        "anterior": valor_anterior,
                        "nuevo": valor_nuevo,
                        "titulo": MAPEO_COLUMNAS[campo],
                    }
                )

        if not cambios:
            print("-" * 72)
            print("Celdas actualizadas: 0")
            resultado_n8n = {
                "fuente": "ipi", "nombre": "IPI manufacturero",
                "periodo": periodo, "estado": "OK",
                "actualizacion": {
                    "archivo_excel": str(ARCHIVO_BASE), "hoja": HOJA_BASE,
                    "periodo": periodo, "fila": fila,
                    "celdas_actualizadas": 0, "hubo_cambios": False,
                    "cambios": [], "detalle_omitido_en_telegram": True,
                    # Aunque el Excel local ya coincida, se entregan los 55
                    # valores a n8n para que el Observatorio pueda repararse
                    # o sincronizarse de manera independiente.
                    "detalle_completo": detalle_completo,
                },
            }
            print("RESULTADO_IPI=" + json.dumps(resultado_n8n, ensure_ascii=False, separators=(",", ":")))
            print("RESULTADO: SIN CAMBIOS")
            print("=" * 72)
            return 0

        # Modifica exclusivamente los 55 encabezados IPI resueltos.
        for cambio in cambios:
            celda = cambio["celda"]
            celda.value = cambio["nuevo"]

            print(
                f"{celda.coordinate}: "
                f"{cambio['anterior']!r} -> "
                f"{cambio['nuevo']!r} "
                "[ACTUALIZADA]"
            )

        # Guarda primero en un archivo temporal.
        archivo_temporal = crear_archivo_temporal()
        libro.save(archivo_temporal)

    finally:
        libro.close()

    if archivo_temporal is None or not archivo_temporal.exists():
        raise RuntimeError(
            "No se generó correctamente el archivo temporal del Excel."
        )

    try:
        # Reemplaza el Excel únicamente después de guardarlo correctamente.
        os.replace(archivo_temporal, ARCHIVO_BASE)

    except PermissionError as error:
        raise PermissionError(
            "No se pudo guardar el Excel. "
            "Cerralo completamente y volvé a ejecutar."
        ) from error

    finally:
        if archivo_temporal is not None:
            archivo_temporal.unlink(missing_ok=True)

    print("-" * 72)
    print(f"Celdas actualizadas: {len(cambios)}")
    print(f"EXCEL_IPI={ARCHIVO_BASE}")
    # Telegram recibe un resumen corto; el detalle completo permanece en el JSON.
    resultado_n8n = {
        "fuente": "ipi",
        "nombre": "IPI manufacturero",
        "periodo": periodo,
        "estado": "OK",
        "actualizacion": {
            "archivo_excel": str(ARCHIVO_BASE),
            "hoja": HOJA_BASE,
            "periodo": periodo,
            "fila": fila,
            "celdas_actualizadas": len(cambios),
            "hubo_cambios": bool(cambios),
            "cambios": [],
            "detalle_omitido_en_telegram": True,
            "detalle_completo": detalle_completo,
        },
    }
    print("RESULTADO_IPI=" + json.dumps(resultado_n8n, ensure_ascii=False, separators=(",", ":")))
    print("RESULTADO: OK")
    print("=" * 72)

    return 0


# ----------------------------------------------------------------------
# EJECUCIÓN
# ----------------------------------------------------------------------

if __name__ == "__main__":
    try:
        codigo_salida = main()
    except Exception as error:
        configurar_salida()

        print("=" * 72, file=sys.stderr)
        print(
            "RESULTADO: ERROR — NO SE ACTUALIZÓ EL EXCEL",
            file=sys.stderr,
        )
        print(
            f"ERROR: {type(error).__name__}: {error}",
            file=sys.stderr,
        )
        print("=" * 72, file=sys.stderr)

        codigo_salida = 1

    raise SystemExit(codigo_salida)
