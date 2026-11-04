import argparse
import hashlib
import html
import json
import re
import sys
import unicodedata
import urllib.request
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urljoin, urlparse

import fitz

RAIZ = Path(__file__).resolve().parents[2]
PAGINA = "https://www.iaraf.org/index.php"

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11,
    "diciembre": 12,
}

TITULOS = [
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ F",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Nominal F",
    "Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Var% Anual Real F",
    "Transferencias automáticas por coparticipación, leyes complementarias y compensaciones Acumulado anual. En millones de $ F",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. en millones $ N",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Nominal N",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Var% Anual Real N",
    "Transferencias automáticas por coparticipación, leyes compl. y compensac. Acumulado anual. En millones de $ N",
]

def normalizar(texto):
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(
        c for c in texto if unicodedata.category(c) != "Mn"
    )
    return " ".join(texto.split()).lower()

def limpiar_html(texto):
    return html.unescape(re.sub(r"<[^>]+>", " ", texto))

def descargar(url):
    pedido = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(pedido, timeout=45) as respuesta:
        return respuesta.read()

def fecha_publicacion(url, titulo):
    m = re.search(
        r"(\d{2})-(\d{2})-(\d{2}|\d{4})(?:$|[/?])",
        url,
    )
    if m:
        anio = int(m[3])
        if anio < 100:
            anio += 2000
        return datetime(anio, int(m[2]), int(m[1]))

    patron = "|".join(MESES)
    m = re.search(
        rf"\b({patron})\s+(?:de\s+)?(20\d{{2}})\b",
        normalizar(titulo),
    )
    if m:
        return datetime(int(m[2]), MESES[m[1]], 1)

    raise ValueError(
        f"No se identifica la fecha del informe: {titulo}"
    )

def descubrir():
    # Busca publicaciones recientes, incluyendo paginacion.
    candidatos = {}
    for inicio in range(0, 50, 5):
        url = PAGINA if inicio == 0 else f"{PAGINA}?start={inicio}"
        pagina = descargar(url).decode("utf-8", errors="replace")

        bloques = re.findall(
            r"<h2\b[^>]*>(.*?)</h2>(.*?)(?=<h2\b|$)",
            pagina,
            flags=re.I | re.S,
        )

        for titulo_html, cuerpo in bloques:
            titulo = limpiar_html(titulo_html)
            texto = normalizar(titulo)

            if not (
                "transferencias automaticas" in texto
                or "coparticipacion nacional" in texto
            ):
                continue

            for etiqueta, interior in re.findall(
                r"<a\b([^>]*)>(.*?)</a>",
                cuerpo,
                flags=re.I | re.S,
            ):
                if "leer mas" not in normalizar(limpiar_html(interior)):
                    continue
                enlace = re.search(
                    r'href=["\']([^"\']+)["\']',
                    etiqueta,
                    flags=re.I,
                )
                if enlace:
                    destino = urljoin(
                        PAGINA,
                        html.unescape(enlace[1]),
                    )
                    candidatos[destino] = (
                        fecha_publicacion(destino, titulo),
                        titulo,
                    )
                    break

        if candidatos:
            # El inicio presenta las publicaciones mas recientes.
            break

    if not candidatos:
        raise RuntimeError(
            "IARAF no mostro un informe mensual de "
            "transferencias en las paginas revisadas."
        )

    url = max(candidatos, key=lambda x: candidatos[x][0])
    return url, candidatos[url][1]

def enlace_pdf(url_nota):
    pagina = descargar(url_nota).decode("utf-8", errors="replace")
    enlaces = [
        html.unescape(x)
        for x in re.findall(
            r'href=["\']([^"\']+)["\']',
            pagina,
            flags=re.I,
        )
    ]

    for enlace in enlaces:
        absoluto = urljoin(url_nota, enlace)
        if "drive.google.com" in absoluto:
            partes = urlparse(absoluto)
            identificador = parse_qs(partes.query).get("id", [None])[0]
            if not identificador:
                m = re.search(r"/file/d/([^/]+)", partes.path)
                identificador = m[1] if m else None
            if identificador:
                return (
                    "https://drive.google.com/uc?export=download&id="
                    + identificador
                )

        if urlparse(absoluto).path.lower().endswith(".pdf"):
            return absoluto

    raise RuntimeError(
        "El informe mensual existe, pero no se encontro "
        "su enlace de descarga."
    )

def numero(texto):
    texto = texto.strip().replace("−", "-").replace("–", "-")
    if not re.fullmatch(r"-?\d[\d.]*(?:,\d+)?%?", texto):
        raise ValueError(f"Valor no numerico en el cuadro: {texto}")
    return float(
        texto.replace(".", "").replace(",", ".").replace("%", "")
    )

def fila(texto, jurisdiccion):
    lineas = [x.strip() for x in texto.splitlines() if x.strip()]
    posiciones = [
        i for i, linea in enumerate(lineas)
        if normalizar(linea) == normalizar(jurisdiccion)
    ]
    if len(posiciones) != 1:
        raise ValueError(
            f"Fila inexistente o duplicada: {jurisdiccion}"
        )
    inicio = posiciones[0] + 1
    datos = lineas[inicio:inicio + 3]
    if len(datos) != 3:
        raise ValueError(f"Fila incompleta: {jurisdiccion}")
    return [numero(x) for x in datos]

def extraer_pdf(contenido):
    documento = fitz.open(stream=contenido, filetype="pdf")
    paginas = [pagina.get_text() for pagina in documento]
    documento.close()

    def cuadro(numero_cuadro):
        coincidencias = [
            texto for texto in paginas
            if re.search(
                rf"\bCuadro\s+{numero_cuadro}\b",
                texto,
                flags=re.I,
            )
            and "Jurisdicción" in texto
        ]
        if len(coincidencias) != 1:
            raise ValueError(
                f"No se identifica un unico Cuadro {numero_cuadro}. "
                "Posible cambio de estructura del informe."
            )
        return coincidencias[0]

    mensual = cuadro(1)
    anual = cuadro(5)

    patron = "|".join(MESES)
    fechas_mensuales = set(re.findall(
        rf"\b({patron})\s+(?:de\s+)?(20\d{{2}})\b",
        normalizar(mensual),
    ))
    if len(fechas_mensuales) != 1:
        raise ValueError(
            "No se identifica un periodo mensual unico."
        )

    nombre_mes, anio = next(iter(fechas_mensuales))
    periodo = f"{int(anio):04d}-{MESES[nombre_mes]:02d}-01"

    if not re.search(
        rf"\b{nombre_mes}\s+(?:de\s+)?{anio}\b",
        normalizar(anual),
    ):
        raise ValueError(
            "El cuadro mensual y el acumulado no tienen "
            "el mismo periodo."
        )

    formosa = fila(mensual, "Formosa")
    nacion = fila(mensual, "Total")
    acumulado_formosa = fila(anual, "Formosa")[0]
    acumulado_nacion = fila(anual, "Total")[0]

    if not (
        0 < formosa[0] <= nacion[0]
        and acumulado_formosa >= formosa[0]
        and acumulado_nacion >= nacion[0]
    ):
        raise ValueError(
            "Los montos mensuales y acumulados no son coherentes."
        )

    valores = dict(zip(TITULOS, [
        *formosa, acumulado_formosa,
        *nacion, acumulado_nacion,
    ]))

    return periodo, valores

# --------------------------------------------------------------- fuente oficial (DNAP)
PAGINA_DNAP = "https://www.argentina.gob.ar/economia/sechacienda/asuntosprovinciales/ron"
MESES_HOJA = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
              "septiembre", "octubre", "noviembre", "diciembre"]


def _consolidada(anio):
    """Libro 'Información consolidada' del año (Dirección Nacional de Asuntos Provinciales)."""
    import io
    from openpyxl import load_workbook
    pagina = descargar(PAGINA_DNAP).decode("utf-8", errors="replace")
    enlaces = [urljoin("https://www.argentina.gob.ar/", h.replace("blank:#", ""))
               for h in re.findall(r'href=["\']([^"\']+)["\']', pagina, flags=re.I)]
    candidatos = [u for u in enlaces
                  if re.search(rf"informacion_consolidada_?{anio}[^/]*\.xlsx?$", u, flags=re.I)]
    if not candidatos:
        raise RuntimeError(f"La página de la DNAP no muestra la información consolidada {anio}")
    contenido = descargar(candidatos[0])
    if candidatos[0].lower().endswith(".xls"):
        raise RuntimeError(f"La información consolidada {anio} está en .xls (formato viejo)")
    return load_workbook(io.BytesIO(contenido), read_only=True, data_only=True), candidatos[0]


def _mes_dnap(libro, mes):
    """(Formosa, Provincias + CABA) en millones de $ de la columna Total (1)+(2) del mes."""
    nombre = next((h for h in libro.sheetnames if normalizar(h).startswith(MESES_HOJA[mes - 1][:4])), None)
    if not nombre:
        return None
    filas = list(libro[nombre].iter_rows(values_only=True))
    col = None
    for fila in filas[:15]:
        for k, v in enumerate(fila):
            if isinstance(v, str) and normalizar(v) == "total" and k > 10:
                col = k
    if col is None:
        raise RuntimeError(f"Hoja {nombre}: no se encuentra la columna Total (1)+(2)")
    def valor(etiqueta):
        for fila in filas:
            if fila and isinstance(fila[0], str) and normalizar(fila[0]) == etiqueta \
                    and isinstance(fila[col], (int, float)):
                return float(fila[col])
        return None
    formosa, provincias, caba = valor("formosa"), valor("provincias"), valor("c.a.b.a")
    if not formosa or not provincias:
        return None
    return formosa, provincias + (caba or 0.0)


SERIE_IPC = "https://www.indec.gob.ar/ftp/cuadros/economia/serie_ipc_divisiones.csv"
PAGINA_REM = "https://www.bcra.gob.ar/relevamiento-expectativas-mercado-rem/"


def _indices_ipc_nacion():
    """{'AAAAMM': índice IPC nivel general Nacional} desde la serie oficial del INDEC."""
    texto = descargar(SERIE_IPC).decode("latin-1", errors="replace")
    indices = {}
    for linea in texto.splitlines():
        c = [x.strip().strip('"') for x in linea.split(";")]
        if len(c) >= 8 and c[0] == "0" and c[7].lower() == "nacional":
            try:
                indices[c[3]] = float(c[4].replace(",", "."))
            except ValueError:
                pass
    if not indices:
        raise RuntimeError("La serie del IPC del INDEC no trae el nivel general Nacional")
    return indices


def _rem_mensual(anio, mes):
    """Mediana del REM (BCRA) para la inflación mensual de ese mes, si está publicada."""
    import io
    from datetime import date, timedelta
    from openpyxl import load_workbook
    pagina = descargar(PAGINA_REM).decode("utf-8", errors="replace")
    enlaces = [urljoin(PAGINA_REM, h) for h in re.findall(r'href=["\']([^"\']+)["\']', pagina, flags=re.I)]
    tablas = [u for u in enlaces if re.search(r"tablas[^/]*\.xlsx$", u, flags=re.I)]
    if not tablas:
        return None
    libro = load_workbook(io.BytesIO(descargar(tablas[0])), read_only=True, data_only=True)
    hoja = libro[libro.sheetnames[0]]
    fin_mes = (date(anio + (mes == 12), mes % 12 + 1, 1) - timedelta(days=1))
    en_ipc = False
    for fila in hoja.iter_rows(values_only=True):
        textos = " ".join(str(x) for x in fila if isinstance(x, str)).lower()
        if "precios minoristas" in textos:
            en_ipc = "nivel general" in textos   # la primera tabla es el IPC nivel general Nacional
            continue
        if not en_ipc or len(fila) < 4:
            continue
        periodo, referencia, mediana = fila[1], fila[2], fila[3]
        if isinstance(periodo, (int, float)):
            periodo = date(1899, 12, 30) + timedelta(days=int(periodo))
        if hasattr(periodo, "year") and (periodo.year, periodo.month) == (fin_mes.year, fin_mes.month) \
                and "mensual" in str(referencia).lower() and isinstance(mediana, (int, float)):
            return float(mediana), tablas[0].rsplit("/", 1)[-1]
    return None


def inflacion_interanual_estimada(periodo):
    """Mismo cálculo que IARAF: IPC Nacional del INDEC hasta el mes anterior por una
    inflación supuesta para el mes informado, contra el mismo mes del año anterior.
    Supuesto del mes: mediana del REM del BCRA; si no está, el dato del INDEC del mes
    (si ya salió) o, en último caso, la inflación del mes anterior."""
    anio, mes = int(periodo[:4]), int(periodo[5:7])
    previo = f"{anio - (mes == 1)}{(mes - 2) % 12 + 1:02d}"
    hace_un_anio = f"{anio - 1}{mes:02d}"
    indices = _indices_ipc_nacion()
    if previo not in indices or hace_un_anio not in indices:
        raise RuntimeError(f"Faltan índices del IPC ({previo} o {hace_un_anio})")
    supuesto = None
    try:
        supuesto = _rem_mensual(anio, mes)
        if supuesto:
            supuesto = (supuesto[0], f"REM BCRA ({supuesto[1]})")
    except Exception as error:  # noqa: BLE001
        print(f"AVISO: no se pudo leer el REM del BCRA: {error}")
    actual = f"{anio}{mes:02d}"
    if not supuesto and actual in indices:
        supuesto = ((indices[actual] / indices[previo] - 1) * 100, "INDEC (dato del mes)")
    if not supuesto:
        antes = f"{int(previo[:4]) - (int(previo[4:]) == 1)}{(int(previo[4:]) - 2) % 12 + 1:02d}"
        supuesto = ((indices[previo] / indices[antes] - 1) * 100, "INDEC (mes anterior, sin REM)")
    tasa = (indices[previo] * (1 + supuesto[0] / 100) / indices[hace_un_anio] - 1) * 100
    return tasa, supuesto[0], supuesto[1]


def desde_dnap():
    """Mismos datos que publica IARAF, tomados de la fuente oficial (Ministerio de Economía)."""
    hoy = datetime.now()
    libro, url = _consolidada(hoy.year)
    meses = [m for m in range(1, 13) if _mes_dnap(libro, m)]
    if not meses and hoy.month <= 2:
        libro, url = _consolidada(hoy.year - 1)
        meses = [m for m in range(1, 13) if _mes_dnap(libro, m)]
    if not meses:
        raise RuntimeError("La información consolidada de la DNAP no tiene meses con datos")
    anio = int(re.search(r"(20\d{2})", url.rsplit("/", 1)[-1])[1])
    mes = max(meses)
    actual = [_mes_dnap(libro, m) for m in range(1, mes + 1)]
    if any(x is None for x in actual):
        raise RuntimeError("Faltan meses intermedios en la información consolidada de la DNAP")
    previo, _ = _consolidada(anio - 1)
    anterior = _mes_dnap(previo, mes)
    if not anterior:
        raise RuntimeError(f"No se encuentra {MESES_HOJA[mes - 1]} {anio - 1} en la DNAP")
    periodo = f"{anio:04d}-{mes:02d}-01"
    f, n = actual[-1]
    nom_f = round((f / anterior[0] - 1) * 100, 1)
    nom_n = round((n / anterior[1] - 1) * 100, 1)
    valores = {
        TITULOS[0]: round(f), TITULOS[1]: nom_f, TITULOS[3]: round(sum(x[0] for x in actual)),
        TITULOS[4]: round(n), TITULOS[5]: nom_n, TITULOS[7]: round(sum(x[1] for x in actual)),
    }
    try:
        tasa, supuesto, origen = inflacion_interanual_estimada(periodo)
        valores[TITULOS[2]] = round(((1 + nom_f / 100) / (1 + tasa / 100) - 1) * 100, 1)
        valores[TITULOS[6]] = round(((1 + nom_n / 100) / (1 + tasa / 100) - 1) * 100, 1)
        ipc = {"interanual": round(tasa, 2), "supuesto_mes": round(supuesto, 2), "origen_supuesto": origen}
    except Exception as error:  # noqa: BLE001
        print(f"AVISO: no se pudo estimar la inflación del período: {error}")
        ipc = None
    return periodo, valores, url, ipc


def desde_iaraf(args):
    """Informe de IARAF: web, o último PDF guardado si la web no deja consultar."""
    if args.pdf:
        return args.pdf.read_bytes(), None, None, args.pdf.name, True
    try:
        url_nota, titulo = descubrir()
        url_pdf = enlace_pdf(url_nota)
        return descargar(url_pdf), url_nota, url_pdf, titulo, True
    except Exception as error:  # noqa: BLE001
        locales = [*(RAIZ / "fuentes" / "iaraf").glob("*.pdf"),
                   *(RAIZ / "reportes" / "transferencias_fiscales").glob("*.pdf")]
        if not locales:
            raise
        ultimo = max(locales, key=lambda r: r.stat().st_mtime)
        print(f"AVISO: no se pudo consultar IARAF ({error}); se usa {ultimo.name}")
        return ultimo.read_bytes(), None, None, ultimo.name, False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--solo-extraer", action="store_true")
    parser.add_argument("--pdf", type=Path)
    args = parser.parse_args()
    carpeta = RAIZ / "reportes" / "transferencias_fiscales"
    carpeta.mkdir(parents=True, exist_ok=True)

    # Fuente 1: IARAF (trae además su variación real con inflación estimada).
    iaraf = None
    try:
        contenido, url_nota, url_pdf, titulo, en_linea = desde_iaraf(args)
        if not contenido.startswith(b"%PDF"):
            raise RuntimeError("La descarga no devolvio un PDF (pantalla de acceso o cambio de enlace).")
        periodo_i, valores_i = extraer_pdf(contenido)
        iaraf = (periodo_i, valores_i, contenido, url_nota, url_pdf, titulo, en_linea)
    except Exception as error:  # noqa: BLE001
        print(f"AVISO: IARAF no disponible: {error}")

    # Fuente 2: Ministerio de Economía - DNAP (oficial; misma cifra que usa IARAF).
    dnap = None
    try:
        dnap = desde_dnap()
    except Exception as error:  # noqa: BLE001
        print(f"AVISO: DNAP no disponible: {error}")

    if not iaraf and not dnap:
        raise RuntimeError("No se pudo consultar ni IARAF ni la DNAP (Ministerio de Economía).")

    # Si la web de IARAF no dejó consultar pero la fuente 'iaraf' ya guardó el informe de ese mismo mes,
    # se usa ese informe (si no, cada revisión alternaba entre la cifra de IARAF y la estimación DNAP).
    if dnap and (not iaraf or dnap[0] > iaraf[0]):
        local = RAIZ / "fuentes" / "iaraf" / f"iaraf_transferencias_{dnap[0][:4]}_{dnap[0][5:7]}.pdf"
        if local.exists():
            try:
                contenido_l = local.read_bytes()
                periodo_l, valores_l = extraer_pdf(contenido_l)
                if periodo_l == dnap[0]:
                    iaraf = (periodo_l, valores_l, contenido_l, None, None, local.name, False)
            except Exception as error:  # noqa: BLE001
                print(f"AVISO: no se pudo leer {local.name}: {error}")

    notas = []
    usar_dnap = bool(dnap and (not iaraf or dnap[0] > iaraf[0]))
    reemplazar = False
    if iaraf and not usar_dnap:
        # Si ese mes se había cargado con la DNAP (variación real estimada), el informe de
        # IARAF la reemplaza: es la cifra de referencia del Observatorio.
        try:
            previo = json.loads((carpeta / f"transferencias_{iaraf[0][:7]}.json").read_text(encoding="utf-8"))
            reemplazar = str(previo.get("origen_dato", "")).startswith("DNAP")
        except (OSError, ValueError):
            pass
    if usar_dnap:
        periodo, valores, url_dnap, ipc = dnap
        origen = "DNAP - Ministerio de Economía (información consolidada RON)"
        real = (f"variación real estimada como IARAF: inflación i.a. {ipc['interanual']:.1f}% "
                f"(IPC INDEC + {ipc['supuesto_mes']:.2f}% supuesto para el mes, {ipc['origen_supuesto']}); "
                "cuando IARAF publique, se reemplaza por su cifra" if ipc
                else "la variación real queda pendiente (no se pudo estimar la inflación)")
        notas.append(f"Dato tomado de la fuente oficial (Ministerio de Economía - DNAP), {real}.")
        url_nota, url_pdf, titulo, contenido = url_dnap, None, url_dnap.rsplit("/", 1)[-1], None
    else:
        periodo, valores, contenido, url_nota, url_pdf, titulo, en_linea = iaraf
        origen = "IARAF"
        if not en_linea:
            notas.append("La página de IARAF no se pudo consultar (verificación anti-robots); "
                         "se usó el último informe guardado y se controló con la DNAP.")
        pdf = carpeta / f"transferencias_{periodo[:7]}.pdf"
        pdf.write_bytes(contenido)
        # Control cruzado: si la DNAP tiene el mismo mes, los montos tienen que coincidir.
        if dnap and dnap[0] == periodo:
            for t in (TITULOS[0], TITULOS[4]):
                a, b = valores.get(t), dnap[1].get(t)
                if a and b and abs(a / b - 1) > 0.005:
                    notas.append(f"Diferencia IARAF vs DNAP en '{t[-1]}': {a:,.0f} vs {b:,.0f} millones; revisar.")
    if reemplazar:
        notas.append("IARAF publicó el mes que estaba cargado con la DNAP: se reemplazó la variación real estimada por la de IARAF.")
    for nota in notas:
        print("NOTA_TELEGRAM=" + nota)

    resultado = {
        "fuente": "transferencias_fiscales",
        "nombre": "Transferencias automaticas - IARAF",
        "estado": "OK",
        "periodo": periodo,
        "origen_dato": origen,
        "consultado_en": datetime.now().astimezone().isoformat(),
        "url_informe": url_nota,
        "url_pdf": url_pdf,
        "titulo_informe": titulo,
        "sha256_pdf": hashlib.sha256(contenido).hexdigest() if contenido else None,
        "valores": valores,
        "permitir_reemplazos": reemplazar,
        "inflacion_estimada": dnap[3] if usar_dnap else None,
        "columnas_preparadas": len(valores),
        "nota": (
            "Montos de transferencias automáticas (RON: coparticipación, leyes complementarias y "
            "compensaciones) a Formosa y al total de provincias + CABA. Fuente: " + origen + ". "
            "El acumulado corresponde a enero hasta el mes informado."
        ),
    }

    reporte = carpeta / f"transferencias_{periodo[:7]}.json"
    resultado["reporte"] = str(reporte)
    reporte.write_text(
        json.dumps(resultado, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(json.dumps(resultado, ensure_ascii=False, indent=2))
    print("ESTADO=OK")

if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"ERROR_TRANSFERENCIAS={error}")
        print("ESTADO=ERROR")
        raise SystemExit(1)