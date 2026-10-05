import ast
import shutil
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent

def cambiar(texto, anterior, nuevo):
    cantidad = texto.count(anterior)
    if cantidad != 1:
        raise RuntimeError(
            "Version diferente del programa; no se modificaron archivos. "
            f"Coincidencias: {cantidad}. Fragmento: {anterior[:90]}"
        )
    return texto.replace(anterior, nuevo, 1)

def preparar_ieric(texto):
    cambios = [
        (
            '"periodos": [periodo_destino],',
            '"periodos": periodos_fuente,'
        ),
        (
            '"Regla del Observatorio: todo el bloque se registra en el período laboral. "',
            '"Empresas en su mes oficial; empleo y salarios en su mes laboral. "'
        ),
        (
            '    fila = localizar_fila(ws, laboral["anio"], laboral["mes_numero"])',
            '    filas = {}'
        ),
        (
            '    valores = {}',
            '    cargas = {}'
        ),
        (
            '        celda = ws.cell(fila, columnas[clave])',
            '''        datos_serie = resultado["series"][serie]
        fila = localizar_fila(ws, datos_serie["anio"], datos_serie["mes_numero"])
        filas[serie] = fila
        fecha = f'{datos_serie["anio"]:04d}-{datos_serie["mes_numero"]:02d}-01'
        carga = cargas.setdefault(fecha, {
            "hoja": HOJA,
            "row_number": fila,
            "periodo": fecha,
            "valores": {}
        })
        celda = ws.cell(fila, columnas[clave])'''
        ),
        (
            '        valores[titulo_real] = valor',
            '        carga["valores"][titulo_real] = valor'
        ),
        (
            '"filas": {nombre: fila for nombre in resultado["series"]},',
            '"filas": filas,'
        ),
        (
            '"cargas_google": [{"hoja": HOJA, "row_number": fila,\n'
            '                                "periodo": periodo_destino, "valores": valores}]',
            '"cargas_google": list(cargas.values())'
        )
    ]
    for anterior, nuevo in cambios:
        texto = cambiar(texto, anterior, nuevo)
    ast.parse(texto)
    return texto

def preparar_general(texto):
    texto = cambiar(
        texto,
        '    if not any(r.get("fuente") == "dnrpa" and r.get("hubo_cambios")\n'
        '               for r in resultados):\n'
        '        seleccion = [fuente for fuente in seleccion if fuente != "dnrpa"]',
        ''
    )
    texto = cambiar(
        texto,
        '        if not (resultado.get("hubo_cambios") or forzar):\n'
        '            continue',
        ''
    )
    arbol = ast.parse(texto)
    funciones = [
        nodo for nodo in arbol.body
        if isinstance(nodo, ast.FunctionDef)
        and nodo.name == "ejecutar_fuente"
    ]
    if len(funciones) != 1:
        raise RuntimeError("No se identifica ejecutar_fuente.")
    llamadas = 0
    manejadores = 0
    for nodo in ast.walk(funciones[0]):
        if (
            isinstance(nodo, ast.Call)
            and isinstance(nodo.func, ast.Attribute)
            and isinstance(nodo.func.value, ast.Name)
            and nodo.func.value.id == "subprocess"
            and nodo.func.attr == "run"
        ):
            if any(k.arg == "timeout" for k in nodo.keywords):
                raise RuntimeError("Esta version ya tiene timeout; requiere revision.")
            nodo.keywords.append(
                ast.keyword(arg="timeout", value=ast.Constant(value=300))
            )
            llamadas += 1
        if (
            isinstance(nodo, ast.ExceptHandler)
            and isinstance(nodo.type, ast.Name)
            and nodo.type.id == "OSError"
        ):
            nodo.type = ast.Name(id="Exception", ctx=ast.Load())
            manejadores += 1
    if llamadas != 1 or manejadores != 1:
        raise RuntimeError("La estructura de generales es diferente.")
    corregido = ast.unparse(ast.fix_missing_locations(arbol)) + "\n"
    ast.parse(corregido)
    return corregido

def main():
    ieric = RAIZ / "scripts" / "ieric" / "ejecutar_ieric.py"
    general = RAIZ / "ejecutar_general.py"
    for ruta in (ieric, general):
        if not ruta.is_file():
            raise FileNotFoundError(f"No existe: {ruta}")

    # Valida los dos cambios antes de escribir.
    trabajos = [
        (ieric, preparar_ieric(ieric.read_text(encoding="utf-8-sig"))),
        (general, preparar_general(general.read_text(encoding="utf-8-sig")))
    ]
    marca = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    respaldos = []
    for ruta, contenido in trabajos:
        respaldo = ruta.with_name(f"{ruta.stem}.respaldo_{marca}.py")
        shutil.copy2(ruta, respaldo)
        respaldos.append((ruta, respaldo))
    try:
        for ruta, contenido in trabajos:
            temporal = ruta.with_suffix(".correccion.tmp")
            temporal.write_text(contenido, encoding="utf-8")
            temporal.replace(ruta)
    except Exception:
        for ruta, respaldo in respaldos:
            shutil.copy2(respaldo, ruta)
        raise
    print("CORRECCION_PROGRAMAS=OK")
    print("IERIC: empresas en su mes; empleo y salarios en su mes.")
    print("GENERALES: limite de 5 minutos por extractor.")
    print("No se ejecutaron indicadores ni se modifico el Sheet.")
    for ruta, respaldo in respaldos:
        print(f"RESPALDO={respaldo}")

if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"CORRECCION_PROGRAMAS=ERROR: {error}")
        raise SystemExit(1)