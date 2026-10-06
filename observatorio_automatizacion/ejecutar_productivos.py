"""Ejecutor independiente de indicadores productivos para n8n."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from openpyxl import load_workbook


RAIZ = Path(__file__).resolve().parent
ARCHIVO_EXCEL = RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx"
HOJA_DESTINO = "BASE DE DATOS - Indicadores Pro"

sys.path.insert(0, str(RAIZ / "scripts" / "comun"))
from control_resultados import ErrorControlFuente, validar_productivo


def serializar_json(valor):
    """Convierte tipos de Excel/Python a valores aceptados por JSON y n8n."""
    if isinstance(valor, (datetime, date)):
        return valor.isoformat()
    if isinstance(valor, Path):
        return str(valor)
    raise TypeError(f"Object of type {valor.__class__.__name__} is not JSON serializable")


def capturar_hoja() -> dict:
    libro = load_workbook(ARCHIVO_EXCEL, read_only=True, data_only=False)
    try:
        hoja = libro[HOJA_DESTINO]
        filas = list(hoja.iter_rows(values_only=True))
        if len(filas) < 2:
            return {"encabezados": {}, "valores": {}}
        # En Indicadores Productivos la fila 1 contiene los bloques generales
        # y los encabezados reales de las columnas se encuentran en la fila 2.
        return {
            "encabezados": {c: valor for c, valor in enumerate(filas[1], start=1)},
            "valores": {
                (f, c): valor
                for f, fila in enumerate(filas[2:], start=3)
                for c, valor in enumerate(fila, start=1)
            },
        }
    finally:
        libro.close()


def normalizar_periodo(valor) -> str:
    meses = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
    if isinstance(valor, datetime):
        # En la hoja histórica algunos períodos fueron guardados como fechas
        # donde el año está en el día (por ejemplo 20/01/2025 = ene-20).
        anio = 2000 + valor.day if 20 <= valor.day <= 99 else valor.year
        return f"{meses[valor.month - 1]} {anio % 100:02d}"
    texto = str(valor).strip()
    coincidencia = re.fullmatch(r"(20\d{2})-(\d{2})(?:-\d{2})?", texto)
    if coincidencia:
        anio, mes = map(int, coincidencia.groups())
        return f"{meses[mes - 1]} {anio % 100:02d}"
    return texto


def periodo_iso(valor) -> str:
    """Formato técnico que usa el subworkflow para localizar año y mes."""
    if isinstance(valor, datetime):
        anio = 2000 + valor.day if 20 <= valor.day <= 99 else valor.year
        return f"{anio:04d}-{valor.month:02d}-01"
    texto = str(valor or "").strip().casefold()
    coincidencia = re.fullmatch(r"(20\d{2})-(\d{2})(?:-\d{2})?", texto)
    if coincidencia:
        return f"{int(coincidencia.group(1)):04d}-{int(coincidencia.group(2)):02d}-01"
    meses = {m: i + 1 for i, m in enumerate(
        ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
    )}
    coincidencia = re.fullmatch(r"([a-z]{3})\s+(\d{2})", texto)
    if coincidencia and coincidencia.group(1) in meses:
        return f"20{coincidencia.group(2)}-{meses[coincidencia.group(1)]:02d}-01"
    raise ValueError(f"No se pudo normalizar el período productivo: {valor}")


def normalizar_titulo(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c)).casefold()
    return " ".join(re.sub(r"[^a-z0-9]+", " ", texto).split())


ENCABEZADOS_PECUARIA = {
    "periodo",
    "stock fsa cabezas bovinos",
    "stock nacion cabezas bovinos",
    "faena fsa cabezas bovinos",
    "faena nacion cabezas bovinos",
    "consumo per capita n kg hab ano bovinos",
    "stock fsa cabezas porcinos",
    "stock n cabezas porcinos",
    "faena total n cabezas porcinos",
    "consumo per capita n kg hab ano porcinos",
    "produccion destino industria miles de huevos n aves",
    "produccion carne tn",
    "faena miles cab n aves",
    "consumo per capita n kg hab ano aves",
    "consumo total de proteina de origen animal kg n",
}
ENCABEZADO_PECUARIA_OPCIONAL = "produccion carne tn"


def columnas_pecuarias(encabezados: dict) -> list[int]:
    """Localiza el bloque por títulos; no depende de letras ni posiciones."""
    encontradas = {
        normalizar_titulo(titulo): columna
        for columna, titulo in encabezados.items()
        if normalizar_titulo(titulo) in ENCABEZADOS_PECUARIA
    }
    faltantes = (ENCABEZADOS_PECUARIA - {ENCABEZADO_PECUARIA_OPCIONAL}) - set(encontradas)
    if faltantes:
        raise ValueError(
            "Faltan encabezados pecuarios en la hoja: " + ", ".join(sorted(faltantes))
        )
    return [encontradas[titulo] for titulo in ENCABEZADOS_PECUARIA if titulo in encontradas]


def preparar_actualizaciones_google(antes: dict, despues: dict) -> list[dict]:
    encabezados = despues["encabezados"]
    filas = {}
    for f, c in set(antes["valores"]) | set(despues["valores"]):
        previo = antes["valores"].get((f, c))
        nuevo = despues["valores"].get((f, c))
        if previo == nuevo:
            continue
        titulo = encabezados.get(c)
        if titulo in (None, ""):
            continue
        item = filas.setdefault(
            f,
            {
                "hoja": HOJA_DESTINO,
                "row_number": f,
                "periodo": None,
                "valores": {},
            },
        )
        item["valores"][str(titulo)] = nuevo
    return list(filas.values())


def preparar_filas_actuales_google(resultados: list[dict], estado: dict) -> list[dict]:
    """Entrega la fila mensual completa para una sincronizacion idempotente."""
    encabezados = estado["encabezados"]
    valores = estado["valores"]
    salida = []
    filas_incluidas = set()

    for resultado in resultados:
        if resultado.get("estado") != "OK":
            continue
        candidatas = resultado.get("filas_localizadas") or [resultado.get("fila_localizada")]
        # Si la ejecución anterior ya había escrito el Excel local, puede no
        # haber cambios nuevos. En ese caso se localizan las filas por período
        # para sincronizarlas igualmente con Google Sheets.
        if resultado.get("fuente") == "pecuaria" and not any(isinstance(f, int) for f in candidatas):
            columnas_bloque = columnas_pecuarias(encabezados)
            col_periodo = next(
                (c for c in columnas_bloque if normalizar_titulo(encabezados.get(c)) == "periodo"),
                None,
            )
            # Sin cambios locales, se reenvía únicamente el período más
            # reciente. Los períodos históricos ya cargados no deben volver a
            # pasar por el subworkflow del Observatorio.
            periodo_actual = resultado.get("periodo")
            periodos = {normalizar_periodo(periodo_actual)} if periodo_actual else set()
            candidatas = [
                f for (f, c), valor in valores.items()
                if c == col_periodo and normalizar_periodo(valor) in periodos
            ]
        for fila in candidatas:
            if not isinstance(fila, int) or fila < 3 or fila in filas_incluidas:
                continue
            datos = {}
            columnas = columnas_pecuarias(encabezados) if resultado.get("fuente") == "pecuaria" else range(3, 19)
            for columna in columnas:
                titulo = encabezados.get(columna)
                valor = valores.get((fila, columna))
                if titulo not in (None, "") and valor not in (None, "", 0):
                    datos[str(titulo)] = (
                        normalizar_periodo(valor)
                        if normalizar_titulo(titulo) == "periodo"
                        else valor
                    )
            if datos:
                columna_periodo_fila = next(
                    (c for c in columnas if normalizar_titulo(encabezados.get(c)) == "periodo"),
                    None,
                )
                periodo_fila = resultado.get("periodos_por_fila", {}).get(
                    str(fila), valores.get((fila, columna_periodo_fila), resultado.get("periodo"))
                )
                salida.append({"hoja": HOJA_DESTINO, "row_number": fila,
                               "periodo": periodo_iso(periodo_fila), "valores": datos})
                filas_incluidas.add(fila)

    return salida


TITULO_CARNE_AVIAR_SHEET = "Produccion/ Carne (tn)"


def agregar_carne_aviar(cargas: list[dict]) -> None:
    """La producción de carne aviar (tn) existe en el Sheet pero no en el Excel local.

    Se toma del último informe de pecuaria y se agrega a la carga del mismo período,
    para que esa columna no quede vacía en la plataforma."""
    carpeta = RAIZ / "scripts" / "reportes" / "productivos" / "pecuaria"
    informes = sorted(carpeta.glob("pecuaria_*.json")) if carpeta.is_dir() else []
    if not informes:
        return
    try:
        datos = json.loads(informes[-1].read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return
    por_periodo = {str(r.get("periodo", ""))[:7]: r["produccion_carne_aviar"]
                   for r in datos.get("registros", []) if r.get("produccion_carne_aviar") not in (None, "", 0)}
    for carga in cargas:
        if carga.get("fuente") == "pecuaria" and str(carga.get("periodo", ""))[:7] in por_periodo:
            carga["valores"][TITULO_CARNE_AVIAR_SHEET] = por_periodo[str(carga["periodo"])[:7]]


@dataclass(frozen=True)
class Fuente:
    nombre: str
    ruta_relativa: str

    def ruta(self) -> Path:
        return RAIZ / self.ruta_relativa


FUENTES = {
    "comercializacion": Fuente(
        "Comercialización de granos",
        "scripts/productivos/comercializacion/ejecutar_comercializacion.py",
    ),
    "pecuaria": Fuente(
        "Actividad pecuaria",
        "scripts/productivos/pecuaria/ejecutar_pecuaria.py",
    ),
}


FUENTES_DIRECTAS = {"estimaciones": "Producción agrícola por cultivo (Estimaciones MAGyP)"}


def seleccionar(texto: str) -> list[str]:
    texto = texto.strip().casefold()
    if texto in {"todas", "todo", "all"}:
        return list(FUENTES) + list(FUENTES_DIRECTAS)
    seleccion = []
    for item in texto.split(","):
        clave = item.strip().casefold()
        if clave and clave not in seleccion:
            seleccion.append(clave)
    desconocidas = [x for x in seleccion if x not in FUENTES and x not in FUENTES_DIRECTAS]
    if desconocidas:
        raise ValueError(
            f"Fuentes desconocidas: {', '.join(desconocidas)}. "
            f"Disponibles: {', '.join(FUENTES)}"
        )
    if not seleccion:
        raise ValueError("No se indicó ninguna fuente")
    return seleccion


def ejecutar(clave: str, fuente: Fuente) -> dict:
    ruta = fuente.ruta()
    if not ruta.is_file():
        return {"fuente": clave, "nombre": fuente.nombre, "estado": "ERROR", "detalle": f"No existe {ruta}"}
    entorno = os.environ.copy()
    entorno["PYTHONIOENCODING"] = "utf-8"
    entorno["PYTHONUTF8"] = "1"
    inicio = time.monotonic()
    try:
        proceso = subprocess.run(
            [sys.executable, "-u", str(ruta)],
            cwd=RAIZ,
            env=entorno,
            text=True,
            capture_output=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            timeout=600,
        )
    except subprocess.TimeoutExpired:
        return {
            "fuente": clave,
            "nombre": fuente.nombre,
            "estado": "ERROR",
            "codigo": 124,
            "duracion_segundos": round(time.monotonic() - inicio, 2),
            "celdas_actualizadas": 0,
            "hubo_cambios": False,
            "salida_final": "",
            "error_final": "La descarga o extracción superó el límite de 10 minutos",
        }
    salida = proceso.stdout.strip()
    error = proceso.stderr.strip()
    cambios = None
    periodo = None
    fila_localizada = None
    filas_localizadas = []
    periodos_por_fila = {}
    periodos_disponibles = []
    detalle = None
    for linea in salida.splitlines():
        if linea.startswith(("RESULTADO_COMERCIALIZACION=", "RESULTADO_PECUARIA=")):
            try:
                detalle = json.loads(linea.split("=", 1)[1])
                periodo = detalle.get("periodo")
                actualizacion = detalle.get("actualizacion", {})
                cambios = actualizacion.get("celdas_actualizadas", 0)
                fila_localizada = actualizacion.get("fila_localizada")
                filas_localizadas = actualizacion.get("filas_localizadas", [])
                periodos_por_fila = actualizacion.get("periodos_por_fila", {})
                periodos_disponibles = [
                    registro.get("periodo") for registro in detalle.get("registros", [])
                    if registro.get("periodo")
                ]
            except json.JSONDecodeError:
                pass
    error_control = ""
    if proceso.returncode == 0:
        try:
            validar_productivo(detalle, fuente=clave)
        except ErrorControlFuente as error_calidad:
            error_control = str(error_calidad)
    estado = "OK" if proceso.returncode == 0 and not error_control else "ERROR"
    return {
        "fuente": clave,
        "nombre": fuente.nombre,
        "estado": estado,
        "codigo": proceso.returncode,
        "duracion_segundos": round(time.monotonic() - inicio, 2),
        "periodo": periodo,
        "fila_localizada": fila_localizada,
        "filas_localizadas": filas_localizadas,
        "periodos_por_fila": periodos_por_fila,
        "periodos_disponibles": periodos_disponibles,
        "celdas_actualizadas": cambios if estado == "OK" else 0,
        "hubo_cambios": bool(cambios) if estado == "OK" else False,
        "detalle": error_control,
        "salida_final": salida[-6000:],
        "error_final": (
            error[-3000:]
            + ("\nCONTROL_CALIDAD=" + error_control if error_control else "")
        ).strip(),
    }


def main() -> int:
    for flujo in (sys.stdout, sys.stderr):
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass
    parser = argparse.ArgumentParser(description="Ejecutor de indicadores productivos")
    parser.add_argument("--fuentes", default="todas")
    parser.add_argument("--listar", action="store_true")
    parser.add_argument("--estricto", action="store_true")
    args = parser.parse_args()
    if args.listar:
        listado = {
            clave: {"nombre": fuente.nombre, "ruta": fuente.ruta_relativa, "existe": fuente.ruta().is_file()}
            for clave, fuente in FUENTES.items()
        }
        print(json.dumps(listado, ensure_ascii=False, indent=2))
        print("ESTADO=" + ("OK" if all(x["existe"] for x in listado.values()) else "ERROR"))
        return 0
    try:
        seleccion = seleccionar(args.fuentes)
    except ValueError as error:
        print(f"ERROR={error}")
        print("ESTADO=ERROR")
        return 2
    inicio = datetime.now().astimezone()
    estado_antes = capturar_hoja()
    resultados = [ejecutar(clave, FUENTES[clave]) for clave in seleccion if clave in FUENTES]
    errores = [x["fuente"] for x in resultados if x["estado"] == "ERROR"]
    total_cambios = sum(x.get("celdas_actualizadas") or 0 for x in resultados)
    estado_despues = capturar_hoja()
    actualizaciones_google = preparar_filas_actuales_google(resultados, estado_despues)
    # Escritura directa y verificada en el Sheet (antes la hacía n8n y caía en filas equivocadas).
    for carga in actualizaciones_google:
        carga["fuente"] = "pecuaria" if any(normalizar_titulo(t) == "periodo" for t in carga["valores"]) else "comercializacion"
    agregar_carne_aviar(actualizaciones_google)
    # Son los datos que la extracción oficial de esta revisión acaba de confirmar:
    # pueden corregir lo que haya quedado mal en el Sheet (la fila se ubica por período).
    for carga in actualizaciones_google:
        carga["confirmada_por_fuente"] = True
    try:
        sys.path.insert(0, str(RAIZ / "scripts" / "comun"))
        from sheet_productivos import escribir_productivos
        sheet_directo = escribir_productivos(actualizaciones_google, ARCHIVO_EXCEL)
    except Exception as error:
        sheet_directo = {"estado": "ERROR", "por_fuente": {}, "errores": ["google_sheet"], "detalle": str(error)}
        errores.append("google_sheet")
        resultados.append({"fuente": "google_sheet", "nombre": "Escritura en Google Sheets (productivos)",
                           "estado": "ERROR", "detalle": f"No se pudo escribir en el Sheet: {error}",
                           "celdas_actualizadas": 0})
    # Estimaciones agrícolas (Cultivo / Campaña / superficies / producción): escriben directo
    # en el bloque "Producción" de la pestaña productiva.
    if "estimaciones" in seleccion:
        resultado_est = {"fuente": "estimaciones", "nombre": FUENTES_DIRECTAS["estimaciones"], "estado": "OK"}
        try:
            sys.path.insert(0, str(RAIZ / "scripts" / "productivos" / "estimaciones"))
            from estimaciones import ejecutar as ejecutar_estimaciones
            info = ejecutar_estimaciones()
            sheet_directo.setdefault("por_fuente", {})["estimaciones"] = info
        except Exception as error:
            resultado_est.update(estado="ERROR", detalle=f"Estimaciones agrícolas: {error}")
            errores.append("estimaciones")
        resultados.append(resultado_est)
    total_cambios = 0
    for resultado in resultados:
        datos = sheet_directo.get("por_fuente", {}).get(resultado.get("fuente"))
        if resultado.get("estado") != "OK":
            continue
        if datos is None:
            resultado["celdas_actualizadas"] = 0
            resultado["hubo_cambios"] = False
            continue
        resultado["celdas_actualizadas"] = datos.get("escritas", 0)
        resultado["hubo_cambios"] = bool(datos.get("escritas"))
        if datos.get("periodos"):
            resultado["periodo_sheet"] = datos["periodos"]
        if datos.get("errores"):
            resultado["estado"] = "ERROR"
            resultado["detalle"] = "Sheet: " + "; ".join(datos["errores"][:3])
            errores.append(resultado["fuente"])
        total_cambios += datos.get("escritas", 0)
    actualizaciones_google = []
    resumen = {
        "inicio": inicio.isoformat(timespec="seconds"),
        "fin": datetime.now().astimezone().isoformat(timespec="seconds"),
        "estado": "ERROR" if errores else "OK",
        "errores": errores,
        "notificacion": {
            "hubo_cambios": total_cambios > 0,
            "total_celdas_actualizadas": total_cambios,
            "fuentes_modificadas": [x["fuente"] for x in resultados if x.get("hubo_cambios")],
        },
        "actualizaciones_google": actualizaciones_google,
        "sheet_directo": sheet_directo,
        "resultados": resultados,
    }
    try:
        sys.path.insert(0, str(RAIZ / "scripts" / "comun"))
        from mensaje_telegram import armar_mensaje
        resumen["notificacion"]["mensaje_telegram"] = armar_mensaje(resumen)
    except Exception as error:
        resumen["notificacion"]["mensaje_telegram"] = f"⚠️ No se pudo armar el aviso: {error}"
    carpeta_n8n = RAIZ / "salidas_n8n"
    carpeta_n8n.mkdir(parents=True, exist_ok=True)
    destino_n8n = carpeta_n8n / "resumen_productivos.json"
    temporal_n8n = carpeta_n8n / "resumen_productivos.tmp"
    temporal_n8n.write_text(
        json.dumps(resumen, ensure_ascii=False, separators=(",", ":"), default=serializar_json),
        encoding="utf-8",
    )
    temporal_n8n.replace(destino_n8n)
    print("RESUMEN_PRODUCTIVOS_N8N=" + json.dumps(
        resumen, ensure_ascii=False, separators=(",", ":"), default=serializar_json
    ))
    print(f"ESTADO={resumen['estado']}")
    # El resumen conserva estado ERROR para que n8n lo notifique, pero el nodo
    # continúa hasta leerlo y no bloquea las fuentes productivas correctas.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
