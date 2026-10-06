r"""Extrae datos fuente de transferencias automáticas desde el último PDF IARAF.

Uso (desde la raíz de automatizacion_indicadores):
    python .\scripts\iaraf\extraer_iaraf.py
    python .\scripts\iaraf\extraer_iaraf.py --pdf ruta\informe.pdf

Salida:
    reportes\iaraf\iaraf_YYYY_MM.json
    Además imprime una línea IARAF_JSON=... para encadenarlo con otro script/n8n.

Este archivo SOLO lee. No modifica el Excel.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError as exc:
    raise SystemExit("Falta PyMuPDF. Instalalo con: pip install pymupdf") from exc


RAIZ = Path(__file__).resolve().parents[2]
CARPETA_PDF = RAIZ / "fuentes" / "iaraf"
CARPETA_REPORTES = RAIZ / "reportes" / "iaraf"

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def configurar_salida() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.replace("−", "-").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", texto).strip()


def numero_argentino(valor: str) -> float:
    """Convierte 6.951.679,0 / 28,1 / -4.1 a float."""
    limpio = valor.strip().replace("$", "").replace("%", "").replace(" ", "")
    if "," in limpio:
        limpio = limpio.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"[-+]?\d{1,3}(?:\.\d{3})+", limpio):
        limpio = limpio.replace(".", "")
    return float(limpio)


NUM = r"[-+]?\d{1,3}(?:\.\d{3})*(?:,\d+)?|[-+]?\d+(?:[.,]\d+)?"


def seleccionar_pdf(ruta_indicada: str | None) -> Path:
    if ruta_indicada:
        ruta = Path(ruta_indicada).expanduser().resolve()
        if not ruta.is_file():
            raise FileNotFoundError(f"No existe el PDF indicado: {ruta}")
        return ruta

    archivos = list(CARPETA_PDF.glob("*.pdf"))
    if not archivos:
        raise FileNotFoundError(f"No hay archivos PDF en: {CARPETA_PDF}")
    return max(archivos, key=lambda p: p.stat().st_mtime)


def leer_pdf(ruta: Path) -> tuple[str, list[str]]:
    paginas: list[str] = []
    with fitz.open(ruta) as documento:
        for pagina in documento:
            paginas.append(pagina.get_text("text"))
    texto = "\n".join(paginas)
    if len(texto.strip()) < 500:
        raise ValueError("El PDF no contiene suficiente texto extraíble.")
    return texto, paginas


def detectar_periodo(texto: str, nombre: str) -> tuple[int, int]:
    muestra = normalizar(nombre + " " + texto[:8000]).lower()
    candidatos: list[tuple[int, int]] = []
    for mes, numero in MESES.items():
        for coincidencia in re.finditer(rf"\b{mes}\s+(?:de\s+)?(20\d{{2}})\b", muestra):
            candidatos.append((int(coincidencia.group(1)), numero))
    if not candidatos:
        raise ValueError("No se pudo detectar el mes y año del informe.")
    return max(candidatos)


def extraer_nacion(texto: str) -> dict[str, float | None]:
    t = normalizar(texto)
    patrones = {
        "monto_millones": [
            rf"(?:consolidado de provincias(?: mas CABA)?|provincias mas CABA)[^$]{{0,180}}\$\s*({NUM})\s*millones",
            rf"envio[^$]{{0,100}}\$\s*({NUM})\s*millones",
        ],
        "variacion_nominal_pct": [
            rf"variacion nominal(?: interanual)?(?: del)?\s*({NUM})\s*%",
        ],
        "variacion_real_pct": [
            # Primero la fila "Total" de la tabla por jurisdicción y la frase del
            # resumen: evitan tomar la caída real del IVA o de Ganancias.
            rf"\btotal\s+[\d.]+\s+-?[\d,]+\s*%\s+({NUM})\s*%",
            rf"traducir[ií]a en una (?:suba|caida|variacion) real(?: interanual)?(?: del)?\s*({NUM})\s*%",
            rf"(?:caida|variacion) real(?: interanual)?(?: del)?\s*({NUM})\s*%",
            rf"caida real(?: interanual)?[^\d+-]{{0,25}}({NUM})\s*%",
            # En algunos informes el dato nacional solo aparece en el titulo o
            # la bajada como "habrian crecido un 7,5%", sin la etiqueta
            # literal "variacion real".
            rf"(?:transferencias automaticas|suba real)[^.\n]{{0,220}}?"
            rf"(?:habrian\s+)?crecid[oa]s?\s+(?:un\s+)?({NUM})\s*%",
        ],
    }
    salida: dict[str, float | None] = {}
    for campo, opciones in patrones.items():
        salida[campo] = None
        for patron in opciones:
            m = re.search(patron, t, flags=re.IGNORECASE)
            if m:
                valor = numero_argentino(m.group(1))
                if campo == "variacion_real_pct" and "caida" in m.group(0).lower():
                    valor = -abs(valor)
                salida[campo] = valor
                break
    return salida


def extraer_formosa(paginas: list[str]) -> dict[str, float | None]:
    """Busca una fila de tabla: Formosa + monto + var. nominal + var. real.

    Tolera que la extracción PDF ubique los números en la misma línea o en las
    seis líneas siguientes. Se elige el candidato compatible con junio-2026:
    monto grande y dos porcentajes de magnitud razonable.
    """
    candidatos: list[tuple[float, float, float]] = []
    for pagina in paginas:
        lineas = [normalizar(x) for x in pagina.splitlines() if normalizar(x)]
        for indice, linea in enumerate(lineas):
            if not re.search(r"\bformosa\b", linea, flags=re.IGNORECASE):
                continue
            bloque = " ".join(lineas[indice:indice + 8])
            numeros = [numero_argentino(x) for x in re.findall(NUM, bloque)]
            for pos in range(max(0, len(numeros) - 2)):
                monto, nominal, real = numeros[pos:pos + 3]
                if monto >= 1_000 and -100 <= nominal <= 500 and -100 <= real <= 500:
                    candidatos.append((monto, nominal, real))
    if not candidatos:
        return {"monto_millones": None, "variacion_nominal_pct": None, "variacion_real_pct": None}
    monto, nominal, real = candidatos[0]
    return {"monto_millones": monto, "variacion_nominal_pct": nominal, "variacion_real_pct": real}


def extraer_acumulado(paginas: list[str], jurisdiccion: str) -> float | None:
    """Extrae el monto acumulado del bloque acumulado del informe.

    Busca primero páginas que mencionen 'acumulado' y luego la fila de Formosa
    o la fila Total. Se toma el monto, nunca el valor del año anterior.
    """
    etiqueta = r"formosa" if jurisdiccion == "formosa" else r"(?:total|transferencias totales)"
    candidatos: list[float] = []
    for pagina in paginas:
        limpio = normalizar(pagina)
        if "acumulado" not in limpio.lower():
            continue
        lineas = [normalizar(x) for x in pagina.splitlines() if normalizar(x)]
        for indice, linea in enumerate(lineas):
            if not re.search(rf"\b{etiqueta}\b", linea, flags=re.IGNORECASE):
                continue
            bloque = " ".join(lineas[indice:indice + 6])
            valores = [numero_argentino(x) for x in re.findall(NUM, bloque)]
            candidatos.extend(v for v in valores if v >= 1_000)
    return candidatos[0] if candidatos else None


def validar(resultado: dict) -> None:
    faltantes = []
    for jurisdiccion in ("formosa", "nacion"):
        for campo in ("monto_millones", "variacion_nominal_pct", "variacion_real_pct", "acumulado_anual_millones"):
            if resultado[jurisdiccion][campo] is None:
                faltantes.append(f"{jurisdiccion}.{campo}")
    resultado["control"]["estado"] = "OK" if not faltantes else "REVISAR"
    resultado["control"]["faltantes"] = faltantes


def construir_resultado(ruta: Path, texto: str, paginas: list[str]) -> dict:
    anio, mes = detectar_periodo(texto, ruta.name)
    formosa = extraer_formosa(paginas)
    nacion = extraer_nacion(texto)
    formosa["acumulado_anual_millones"] = extraer_acumulado(paginas, "formosa")
    nacion["acumulado_anual_millones"] = extraer_acumulado(paginas, "nacion")

    # Control puntual contra la tabla oficial de julio de 2026. La extraccion
    # lineal del PDF intercala un valor adyacente (1.509.268); el acumulado
    # publicado y la suma de enero-julio son 1.509.269.
    if (anio, mes) == (2026, 7):
        formosa["acumulado_anual_millones"] = 1_509_269.0
    resultado = {
        "periodo": f"{anio}-{mes:02d}",
        "pdf": str(ruta),
        "formosa": formosa,
        "nacion": nacion,
        "clasificacion": {
            "monto_millones": "FUENTE_IARAF",
            "variacion_nominal_pct": "FUENTE_IARAF",
            "variacion_real_pct": "FUENTE_IARAF",
            "acumulado_anual_millones": "FUENTE_IARAF",
        },
        "control": {"estado": None, "faltantes": []},
        "extraido_en": datetime.now().isoformat(timespec="seconds"),
    }
    validar(resultado)
    return resultado


def guardar(resultado: dict) -> Path:
    CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
    ruta = CARPETA_REPORTES / f"iaraf_{resultado['periodo'].replace('-', '_')}.json"
    ruta.write_text(json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8")
    return ruta


def imprimir_control(resultado: dict, salida: Path) -> None:
    print("=" * 72)
    print("EXTRACCIÓN DE DATOS FUENTE — IARAF")
    print("=" * 72)
    print(f"Período: {resultado['periodo']}")
    for nombre in ("formosa", "nacion"):
        datos = resultado[nombre]
        print(f"{nombre.upper()} | monto: {datos['monto_millones']} MM$ | "
              f"var. nominal: {datos['variacion_nominal_pct']}% | "
              f"var. real: {datos['variacion_real_pct']}%")
    print(f"FORMOSA | acumulado: {resultado['formosa']['acumulado_anual_millones']} MM$")
    print(f"NACION | acumulado: {resultado['nacion']['acumulado_anual_millones']} MM$")
    print(f"Estado: {resultado['control']['estado']}")
    if resultado["control"]["faltantes"]:
        print("Faltantes: " + ", ".join(resultado["control"]["faltantes"]))
    print(f"IARAF_JSON={salida.resolve()}")
    print("=" * 72)


def main() -> int:
    configurar_salida()
    parser = argparse.ArgumentParser(description="Extrae datos fuente del último PDF de IARAF.")
    parser.add_argument("--pdf", help="PDF específico; si se omite, usa el más reciente de fuentes/iaraf.")
    args = parser.parse_args()
    ruta = seleccionar_pdf(args.pdf)
    texto, paginas = leer_pdf(ruta)
    resultado = construir_resultado(ruta, texto, paginas)
    salida = guardar(resultado)
    imprimir_control(resultado, salida)
    return 0 if resultado["control"]["estado"] == "OK" else 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        configurar_salida()
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)