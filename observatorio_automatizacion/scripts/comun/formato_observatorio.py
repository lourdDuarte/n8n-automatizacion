from __future__ import annotations

from datetime import date, timedelta
import os
import time
import re
import unicodedata

SHEET_ID = "1cVd_KBbtjMN8vUHlUks1IKwPf0c4Oq0RcYn4B5N5PkA"
HOJA = "BASE DE DATOS"
API = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"


def _fecha(valor):
    if not valor:
        return None

    if "numberValue" in valor:
        fecha = date(1899, 12, 30) + timedelta(
            days=int(valor["numberValue"])
        )
        return fecha.replace(day=1)

    texto = str(valor.get("stringValue", "")).strip()
    coincidencia = re.fullmatch(
        r"(20\d{2})-(0[1-9]|1[0-2])", texto
    )
    if coincidencia:
        return date(
            int(coincidencia[1]), int(coincidencia[2]), 1
        )

    try:
        return date.fromisoformat(texto[:10]).replace(day=1)
    except ValueError:
        pass

    meses = {
        "ene": 1, "feb": 2, "mar": 3, "abr": 4,
        "may": 5, "jun": 6, "jul": 7, "ago": 8,
        "sep": 9, "oct": 10, "nov": 11, "dic": 12,
    }
    coincidencia = re.fullmatch(
        r"([a-záéíóú]{3,10})[ -](\d{2}|\d{4})",
        texto.lower(),
    )
    if coincidencia and coincidencia[1][:3] in meses:
        anio = int(coincidencia[2])
        if anio < 100:
            anio += 2000
        return date(anio, meses[coincidencia[1][:3]], 1)

    return None


def _clave(texto):
    simple = unicodedata.normalize("NFD", str(texto or ""))
    simple = "".join(
        caracter
        for caracter in simple
        if unicodedata.category(caracter) != "Mn"
    )
    simple = " ".join(simple.split()).casefold()

    # Reconoce diferencias de espacios alrededor de puntuación.
    # No modifica los encabezados del Sheet.
    return re.sub(r"\s*([.,:;])\s*", r"\1", simple)


def _formato(celda):
    return celda.get("userEnteredFormat", {}).get("numberFormat")


def _letra(numero):
    texto = ""
    while numero:
        numero, resto = divmod(numero - 1, 26)
        texto = chr(65 + resto) + texto
    return texto


def solicitudes_formato(filas, encabezados, destinos, sheet_id):
    nombres = {}
    for columna, celda in enumerate(encabezados):
        titulo = celda.get("userEnteredValue", {}).get(
            "stringValue", ""
        )
        clave = _clave(titulo)
        if clave:
            nombres.setdefault(clave, []).append(columna)

    cambios = []
    permitidos = {"NUMBER", "PERCENT", "CURRENCY", "SCIENTIFIC"}

    for numero_fila, titulos in sorted(destinos.items()):
        actual = filas.get(numero_fila, [])

        for titulo in sorted(titulos):
            clave = _clave(titulo)

            if clave == _clave(
                "Índice de Confianza del Consumidor Interior"
            ):
                clave = _clave(
                    "Indice de Confianza del Consumidor "
                    "por Regiones: Interior"
                )

            columnas = nombres.get(clave, [])
            if not columnas:
                raise ValueError(
                    f"Encabezado inexistente: {titulo}"
                )
            if len(columnas) > 1:
                letras = ", ".join(
                    _letra(columna + 1) for columna in columnas
                )
                raise ValueError(
                    f"Encabezado duplicado en {letras}: {titulo}"
                )

            columna = columnas[0]
            celda = actual[columna] if columna < len(actual) else {}

            if _formato(celda):
                continue

            modelo = None
            for anterior in range(numero_fila - 1, 1, -1):
                fila = filas.get(anterior, [])
                previa = (
                    fila[columna] if columna < len(fila) else {}
                )
                formato = _formato(previa)
                valor = (
                    previa.get("effectiveValue")
                    or previa.get("userEnteredValue", {})
                )
                if (
                    formato
                    and formato.get("type") in permitidos
                    and "numberValue" in valor
                ):
                    modelo = formato
                    break

            if modelo is None:
                raise ValueError(
                    f"Sin formato numérico histórico para {titulo} "
                    f"en fila {numero_fila}"
                )

            cambios.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": numero_fila - 1,
                        "endRowIndex": numero_fila,
                        "startColumnIndex": columna,
                        "endColumnIndex": columna + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": modelo
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            })

    return cambios


def preparar_formatos(actualizaciones):
    relevantes = [
        item for item in actualizaciones
        if item.get("hoja") == HOJA and item.get("valores")
    ]
    if not relevantes:
        return {"estado": "NO_REQUERIDO", "celdas": 0}

    if not os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"):
        raise RuntimeError(
            "Falta configurar GOOGLE_APPLICATION_CREDENTIALS"
        )

    import google.auth
    from google.auth.transport.requests import Request
    import requests

    credenciales, _ = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    credenciales.refresh(Request())

    sesion = requests.Session()
    sesion.headers["Authorization"] = (
        f"Bearer {credenciales.token}"
    )

    def pedir(url, **opciones):
        # El Sheet es grande: se lee solo lo necesario y con reintentos (antes se leía toda la
        # hoja con formatos y a veces la API tardaba más de 45 s: "Read timed out").
        ultimo = None
        for intento in range(3):
            try:
                respuesta = sesion.get(url, timeout=120, **opciones)
                respuesta.raise_for_status()
                return respuesta.json()
            except Exception as error:  # noqa: BLE001
                ultimo = error
                time.sleep(5 * (intento + 1))
        raise ultimo

    meta = pedir(
        API,
        params={
            "fields": "sheets.properties(sheetId,title,gridProperties)"
        },
    )
    coincidencias = [
        item["properties"]
        for item in meta["sheets"]
        if item["properties"]["title"] == HOJA
    ]
    if len(coincidencias) != 1:
        raise ValueError(f"No encuentro la hoja {HOJA}")

    hoja = coincidencias[0]
    limite = hoja["gridProperties"]["rowCount"]
    ancho = hoja["gridProperties"]["columnCount"]

    def leer_bloques(rangos):
        datos = pedir(
            API,
            params={
                "ranges": rangos,
                "includeGridData": "true",
                "fields": (
                    "sheets(data(startRow,rowData(values("
                    "userEnteredValue,effectiveValue,"
                    "userEnteredFormat.numberFormat))))"
                ),
            },
        )
        filas = {}
        for bloque in datos["sheets"][0].get("data", []):
            inicio = bloque.get("startRow", 0)
            for desplazamiento, fila in enumerate(
                bloque.get("rowData", [])
            ):
                filas[inicio + desplazamiento + 1] = fila.get(
                    "values", []
                )
        return filas

    # 1) Solo la columna A para ubicar los períodos.
    columna_a = leer_bloques([f"'{HOJA}'!A1:A{limite}"])
    periodos = {}
    for numero, celdas in columna_a.items():
        if not celdas:
            continue
        periodo = _fecha(
            celdas[0].get("effectiveValue")
            or celdas[0].get("userEnteredValue")
        )
        if periodo:
            if periodo in periodos:
                raise ValueError(
                    f"Período duplicado en el Sheet: {periodo}"
                )
            periodos[periodo] = numero

    destinos = {}
    for item in relevantes:
        periodo = _fecha({
            "stringValue": str(item.get("periodo", ""))
        })
        if not periodo:
            raise ValueError(
                f"Período inválido: {item.get('periodo')}"
            )
        numero = periodos.get(periodo)
        if numero is None:
            raise ValueError(f"No existe {periodo} en {HOJA}")

        destinos.setdefault(numero, set()).update(
            titulo for titulo in item["valores"]
            if _clave(titulo) != "periodo"
            and not str(titulo).startswith("__")
        )
    if not destinos:
        return {"estado": "NO_REQUERIDO", "celdas": 0}

    # 2) Encabezados + las filas a formatear y los 3 años anteriores (de donde se copia el
    #    formato). Si ahí no hay un formato de referencia, se lee la hoja completa.
    desde = max(2, min(destinos) - 36)
    rangos = [f"'{HOJA}'!A1:{_letra(ancho)}1", f"'{HOJA}'!A{desde}:{_letra(ancho)}{max(destinos)}"]
    filas = leer_bloques(rangos)
    try:
        solicitudes = solicitudes_formato(
            filas, filas.get(1, []), destinos, hoja["sheetId"]
        )
    except ValueError as error:
        if "Sin formato numérico histórico" not in str(error) or desde == 2:
            raise
        filas = leer_bloques([f"'{HOJA}'!A1:{_letra(ancho)}{max(destinos)}"])
        solicitudes = solicitudes_formato(
            filas, filas.get(1, []), destinos, hoja["sheetId"]
        )

    for inicio in range(0, len(solicitudes), 200):
        respuesta = sesion.post(
            f"{API}:batchUpdate",
            json={
                "requests": solicitudes[inicio:inicio + 200]
            },
            timeout=60,
        )
        respuesta.raise_for_status()

    return {
        "estado": "OK",
        "celdas": len(solicitudes),
        "periodos": len(destinos),
    }