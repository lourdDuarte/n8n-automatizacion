"""Ejecutor de la rama complejos para los nodos existentes de n8n.

Guardar en la raíz de automatizacion_indicadores. IERIC permanece en
scripts/ieric/ejecutar_ieric.py y se ejecuta desde la rama general.
Este ejecutor siempre emite un resumen de la ejecución, incluidos los errores.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
HOJA = "BASE DE DATOS"
FUENTES = {
    "emae": ("EMAE", "ejecutar_emae.py"),
    "salarios": ("Índice de salarios", "ejecutar_salarios.py"),
    "icc": ("Índice de confianza del consumidor", "ejecutar_icc.py"),
    "supermercados": ("Encuesta de supermercados", "ejecutar_supermercados.py"),
    "pobreza": ("Pobreza e indigencia", "ejecutar_pobreza.py"),
}


def configurar_salida():
    for flujo in (sys.stdout, sys.stderr):
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def clave(valor):
    texto = unicodedata.normalize("NFD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = " ".join(texto.split()).casefold()
    return re.sub(r"\s*([.,:;])\s*", r"\1", texto)


def localizar(nombre):
    # Resolver únicamente nombres explícitos. Una coincidencia ambigua falla.
    rutas = sorted(p for p in (RAIZ / "scripts").rglob(nombre) if p.is_file())
    if len(rutas) != 1:
        detalle = ", ".join(str(p.relative_to(RAIZ)) for p in rutas)
        raise ValueError(
            f"Se esperaba un único {nombre} dentro de scripts; encontrados: "
            f"{len(rutas)}" + (f" ({detalle})" if detalle else "")
        )
    return rutas[0]


def seleccionar(texto):
    texto = texto.strip().casefold()
    if texto in ("todas", "todo", "all", "complejo", "complejos"):
        return list(FUENTES)
    seleccion = list(dict.fromkeys(x.strip() for x in texto.split(",") if x.strip()))
    desconocidas = [x for x in seleccion if x not in FUENTES]
    if not seleccion or desconocidas:
        raise ValueError("Fuentes inválidas: " + str(desconocidas or texto)
                         + ". Disponibles: " + ", ".join(FUENTES))
    return seleccion


def extraer_json(salida):
    decoder = json.JSONDecoder()
    posicion = 0
    candidatos = []
    while posicion < len(salida):
        inicio = salida.find("{", posicion)
        if inicio < 0:
            break
        try:
            valor, longitud = decoder.raw_decode(salida[inicio:])
        except json.JSONDecodeError:
            posicion = inicio + 1
            continue
        if isinstance(valor, dict):
            candidatos.append(valor)
        posicion = inicio + longitud
    validos = [x for x in candidatos if x.get("periodo") and any(
        k in x for k in ("valores", "valores_originales", "actualizacion")
    )]
    if not validos:
        raise ValueError("El extractor no emitió un resultado JSON con período y valores")
    return validos[-1]


def preparar_carga(datos, base):
    periodo = str(datos.get("periodo", "")).strip()
    if not re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])(?:-\d{2})?", periodo):
        raise ValueError(f"Período de datos inválido: {periodo!r}")
    validacion = datos.get("validaciones") or {}
    faltantes = validacion.get("faltantes") or validacion.get("columnas_faltantes")
    if faltantes:
        raise ValueError("La fuente informa un bloque incompleto: " + str(faltantes))
    actualizacion = datos.get("actualizacion") or {}
    detalle = actualizacion.get("detalle_completo")
    if isinstance(detalle, list) and detalle:
        valores = {x["titulo"]: x.get("despues") for x in detalle
                   if isinstance(x, dict) and x.get("titulo")}
    else:
        valores = datos.get("valores_originales") or datos.get("valores")
    if not isinstance(valores, dict) or not valores:
        raise ValueError("No se recibieron valores con encabezados para el Observatorio")
    invalidos = [k for k, v in valores.items() if v is None or v == ""]
    if invalidos:
        raise ValueError("Campos sin valor: " + ", ".join(invalidos))

    from openpyxl import load_workbook
    libro = load_workbook(base, read_only=True, data_only=False)
    try:
        ws = libro[HOJA]
        mapa = {}
        for celda in ws[1]:
            if celda.value is not None:
                mapa.setdefault(clave(celda.value), []).append(str(celda.value))
        enviados = {}
        for titulo, valor in valores.items():
            buscado = clave(titulo)
            if buscado == clave("Índice de Confianza del Consumidor Interior"):
                buscado = clave("Indice de Confianza del Consumidor por Regiones: Interior")
            candidatas = mapa.get(buscado, [])
            if len(candidatas) != 1:
                raise ValueError(f"Encabezado no localizado de forma única: {titulo!r}. "
                                 "No se adivina la columna ni se carga el bloque parcialmente")
            if candidatas[0] in enviados:
                raise ValueError(f"Dos campos apuntan a la misma columna: {titulo}")
            enviados[candidatas[0]] = valor
    finally:
        libro.close()
    # Columnas que existen solo en el Google Sheet (p. ej. números índice del EMAE):
    # se envían con su título y el Sheet las ubica por encabezado.
    for titulo, valor in (datos.get("valores_solo_sheet") or {}).items():
        if valor is None or valor == "":
            raise ValueError("Campo sin valor: " + titulo)
        enviados[titulo] = valor
    return {"hoja": HOJA, "periodo": periodo, "valores": enviados}


def ejecutar(fuente, base, carpeta_logs):
    nombre, archivo = FUENTES[fuente]
    inicio = time.monotonic()
    resultado = {"fuente": fuente, "nombre": nombre, "estado": "ERROR",
                 "celdas_actualizadas": 0, "hubo_cambios": False,
                 "periodo": None, "cambios": []}
    salida = error = ""
    try:
        ruta = localizar(archivo)
        comando = [sys.executable, "-u", str(ruta)]
        # Preferir la extracción sin tocar el respaldo cuando el módulo la admite.
        if "--solo-extraer" in ruta.read_text(encoding="utf-8-sig"):
            comando.append("--solo-extraer")
        entorno = os.environ.copy()
        entorno.update(PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        proceso = subprocess.run(comando, cwd=RAIZ, env=entorno, capture_output=True,
                                 text=True, encoding="utf-8", errors="replace",
                                 timeout=600, check=False)
        salida, error = proceso.stdout or "", proceso.stderr or ""
        resultado["codigo"] = proceso.returncode
        if proceso.returncode:
            raise RuntimeError(f"El extractor terminó con código {proceso.returncode}: "
                               + (error or salida)[-2000:])
        datos = extraer_json(salida)
        if str(datos.get("estado", "OK")).upper() not in ("OK", "EXITO", "ÉXITO"):
            raise ValueError("El extractor informa error: " + str(datos.get("detalle") or datos.get("estado")))
        carga = preparar_carga(datos, base)
        resultado.update(estado="OK", periodo=carga["periodo"],
                         detalle="Datos preparados; carga en Sheet pendiente",
                         columnas_preparadas=len(carga["valores"]))
        return resultado, carga
    except subprocess.TimeoutExpired:
        resultado.update(codigo=124, detalle="La descarga o extracción superó 10 minutos")
    except Exception as exc:
        resultado["detalle"] = str(exc)
    finally:
        resultado["duracion_segundos"] = round(time.monotonic() - inicio, 2)
        resultado["salida_final"] = salida[-4000:]
        resultado["error_final"] = error[-4000:] or (resultado.get("detalle", "") if resultado["estado"] == "ERROR" else "")
        log = carpeta_logs / f"{fuente}.log"
        log.write_text(salida + "\n" + error + "\n" + resultado.get("detalle", ""), encoding="utf-8")
        resultado["log"] = str(log)
    return resultado, None


def guardar_resumen(resumen):
    carpeta = RAIZ / "salidas_n8n"
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / "resumen_complejos.json"
    temporal = carpeta / ("resumen_complejos_" + str(os.getpid()) + ".tmp")
    temporal.write_text(json.dumps(resumen, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    os.replace(temporal, destino)
    print("RESUMEN_COMPLEJOS_N8N=" + json.dumps(resumen, ensure_ascii=False, separators=(",", ":")))
    print("ESTADO=" + resumen["estado"])
    print("RESUMEN_ARCHIVO=" + str(destino))


def escribir_en_sheet(resumen, base):
    """Escribe directo en Google Sheets (verificado) y deja el resultado real en cada fuente."""
    sys.path.insert(0, str(RAIZ / "scripts" / "comun"))
    from sheet_directo import escribir_actualizaciones, describir
    try:
        escritura = escribir_actualizaciones(resumen["actualizaciones_google"], base)
    except Exception as exc:
        escritura = {"estado": "ERROR", "por_fuente": {}, "errores": ["sheet"], "detalle": str(exc)}
        resumen["errores"].append("google_sheet")
        resumen["resultados"].append({"fuente": "google_sheet", "nombre": "Escritura en Google Sheets",
                                      "estado": "ERROR", "detalle": f"No se pudo escribir en el Sheet: {exc}",
                                      "celdas_actualizadas": 0, "hubo_cambios": False})
    resumen["sheet_directo"] = escritura
    modificadas, total = [], 0
    for resultado in resumen["resultados"]:
        datos = escritura.get("por_fuente", {}).get(resultado.get("fuente"))
        if resultado.get("estado") != "OK" or datos is None:
            continue
        resultado["detalle"] = describir(datos)
        resultado["celdas_actualizadas"] = datos.get("escritas", 0)
        resultado["hubo_cambios"] = bool(datos.get("escritas"))
        resultado["cambios"] = datos.get("cambios", [])[:40]
        if datos.get("errores"):
            resultado["estado"] = "ERROR"
            resumen["errores"].append(resultado["fuente"])
        if datos.get("escritas"):
            modificadas.append(resultado["fuente"])
            total += datos["escritas"]
    resumen["notificacion"] = {"hubo_cambios": bool(modificadas), "total_celdas_actualizadas": total,
                               "fuentes_modificadas": modificadas}
    # n8n ya no escribe: solo envía el aviso.
    resumen["actualizaciones_google"] = []


def main():
    configurar_salida()
    inicio = datetime.now().astimezone()
    resumen = {"version_ejecutor": "COMPLEJOS_RESUMEN_20260930", "inicio": inicio.isoformat(timespec="seconds"),
               "estado": "ERROR", "solicitadas": [], "ejecutadas": [], "errores": [],
               "omitidas": [], "resultados": [], "actualizaciones_google": [],
               "notificacion": {"hubo_cambios": False, "total_celdas_actualizadas": 0,
                                "fuentes_modificadas": []}}
    try:
        parser = argparse.ArgumentParser(description="Ejecutor de complejos para n8n")
        parser.add_argument("--fuentes", default="todas")
        parser.add_argument("--estricto", action="store_true", help="Compatibilidad: todo fallo se informa siempre")
        parser.add_argument("--listar", action="store_true")
        parser.add_argument("--base", type=Path, default=RAIZ / "base" / "BASE DE DATOS INDICADORES PROPIOS.xlsx")
        args = parser.parse_args()
        if args.listar:
            listado = {}
            for fuente, (nombre, archivo) in FUENTES.items():
                try:
                    ruta = localizar(archivo)
                    listado[fuente] = {"nombre": nombre, "ruta": str(ruta), "existe": True}
                except Exception as exc:
                    listado[fuente] = {"nombre": nombre, "existe": False, "detalle": str(exc)}
            print(json.dumps(listado, ensure_ascii=False, indent=2))
            return 0
        seleccion = seleccionar(args.fuentes)
        resumen["solicitadas"] = seleccion
        logs = RAIZ / "logs" / "complejos" / inicio.strftime("%Y%m%d_%H%M%S_%f")
        logs.mkdir(parents=True, exist_ok=True)
        for fuente in seleccion:
            resultado, carga = ejecutar(fuente, args.base, logs)
            resumen["resultados"].append(resultado)
            resumen["ejecutadas"].append(fuente)
            if resultado["estado"] != "OK":
                resumen["errores"].append(fuente)
            elif carga:
                carga["fuente"] = fuente
                resumen["actualizaciones_google"].append(carga)
        escribir_en_sheet(resumen, args.base)
        resumen["estado"] = "ERROR" if resumen["errores"] else "OK"
    except Exception as exc:
        resumen["errores"].append("ejecutor_complejos")
        resumen["resultados"].append({"fuente": "ejecutor_complejos", "nombre": "Ejecutor de complejos",
                                     "estado": "ERROR", "detalle": str(exc), "error_final": str(exc),
                                     "celdas_actualizadas": 0, "hubo_cambios": False})
    resumen["fin"] = datetime.now().astimezone().isoformat(timespec="seconds")
    try:
        sys.path.insert(0, str(RAIZ / "scripts" / "comun"))
        from mensaje_telegram import armar_mensaje
        resumen["notificacion"]["mensaje_telegram"] = armar_mensaje(resumen)
    except Exception as exc:
        resumen["notificacion"]["mensaje_telegram"] = f"⚠️ No se pudo armar el aviso: {exc}"
    # Word con la descripción para el Observatorio (EMAE, supermercados, pobreza) cuando hay datos nuevos.
    try:
        sys.path.insert(0, str(RAIZ / "scripts" / "reportes" / "descripciones"))
        from generar_descripciones import para_revision
        resumen["descripciones"], linea = para_revision(resumen)
        if linea:
            resumen["notificacion"]["mensaje_telegram"] = (resumen["notificacion"].get("mensaje_telegram") or "") + "\n\n" + linea
    except Exception as exc:
        resumen["descripciones"] = [{"estado": "ERROR", "detalle": str(exc)[:300]}]
    guardar_resumen(resumen)
    # n8n debe leer y notificar el diagnóstico también cuando una fuente falla.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
