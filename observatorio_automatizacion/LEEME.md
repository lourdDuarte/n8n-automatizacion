# Automatización de indicadores — Observatorio (Formosa)

Paquete completo para instalar la automatización en el servidor institucional:
programas en Python + workflows de n8n. Versión del 5 de octubre de 2026.

## Qué hace

Cuatro veces por día (9:30, 11:00, 16:30 y 18:30, hora Argentina) n8n:

1. Ejecuta los programas de Python, que revisan las fuentes oficiales
   (INDEC, IARAF, SIPA, IERIC, DNRPA, SRT, CAMMESA, CECHA, ATP Formosa,
   MAGyP/SAGyP, etc.) y detectan si hay datos nuevos.
2. Carga los datos nuevos en el Google Sheet **BASE DE DATOS** (hoja
   `BASE DE DATOS` y hoja `BASE DE DATOS - Indicadores Productivos`), por
   encabezado y período, sin pisar datos cargados a mano y completando las
   columnas de fórmulas.
3. Avisa por Telegram al grupo del Observatorio: por indicador, "X de Y
   datos en el Sheet" y, si falta algo, el motivo.
4. Si algo falla, el workflow de errores manda el diagnóstico por Telegram.

**El Google Sheet es la base oficial.** El Excel de `base/` es una copia de
trabajo: al comienzo de cada revisión se actualiza solo desde el Sheet.

## Contenido

```
observatorio_automatizacion/
├── LEEME.md                    ← este archivo
├── requirements.txt            ← librerías de Python
├── .env.ejemplo                ← variables de entorno del servicio
├── ejecutar_general.py         ← indicadores generales (IPC, SIPA, IERIC, IARAF, DNRPA, ...)
├── ejecutar_complejos.py       ← EMAE, salarios, ICC, supermercados, pobreza
├── ejecutar_productivos.py     ← comercialización, pecuaria, estimaciones agrícolas
├── base/                       ← copia de trabajo del Excel (se sincroniza con el Sheet)
├── scripts/                    ← un módulo por fuente + scripts/comun (Sheet, Telegram, fórmulas)
│   └── reportes/               ← último reporte de cada fuente (estado; la faena de Formosa lo necesita)
├── reportes/                   ← último reporte de IARAF / transferencias
├── salidas_n8n/  logs/  fuentes/   ← se llenan solos al ejecutar
├── n8n/
│   ├── plantillas/             ← los 3 workflows tal como funcionan hoy (Windows)
│   ├── configurar_n8n.py       ← adapta rutas y comandos al servidor
│   └── importar_linux/         ← ya generados para /opt/observatorio (ver paso 4)
└── documentacion/              ← criterios y correcciones aplicadas
```

Workflows de n8n:

| Archivo | Nombre en n8n | Función |
|---|---|---|
| `1_PRINCIPAL_Telegram.json` | PRINCIPAL - ORDENADO - TELEGRAM MEJORADO | Horarios, ejecución de Python, aviso Telegram |
| `2_Actualizar_Base_del_Observatorio.json` | Actualizar Base del Observatorio | Sub‑workflow que escribe en Google Sheets |
| `3_Aviso_de_errores.json` | AVISO DE ERRORES - DIAGNÓSTICO MEJORADO | Diagnóstico de errores por Telegram |

## Instalación en el servidor (Linux)

Ejemplo con la carpeta `/opt/observatorio`. Si usan otra, cambien las rutas
en todos los pasos y vuelvan a generar los workflows (paso 4).

### 1. Programas y Python

```bash
sudo mkdir -p /opt/observatorio/n8n-files /opt/observatorio/credenciales
sudo unzip observatorio_automatizacion.zip -d /opt/observatorio
sudo mv /opt/observatorio/observatorio_automatizacion /opt/observatorio/automatizacion_indicadores
python3 -m venv /opt/observatorio/venv
/opt/observatorio/venv/bin/pip install -r /opt/observatorio/automatizacion_indicadores/requirements.txt
```

El usuario que ejecuta n8n debe poder **leer y escribir** en
`/opt/observatorio/automatizacion_indicadores` y en `/opt/observatorio/n8n-files`.

### 2. Credencial de Google para Python

1. En Google Cloud, crear una **cuenta de servicio** con la API de Google
   Sheets habilitada y descargar su clave JSON.
2. Guardarla en `/opt/observatorio/credenciales/google_cuenta_servicio.json`.
3. **Compartir el Google Sheet BASE DE DATOS con el correo de la cuenta de
   servicio, como Editor.**
4. Definir la variable `GOOGLE_APPLICATION_CREDENTIALS` en el servicio de n8n
   (paso 3). Sin esta variable, Python no puede escribir en el Sheet.

### 3. n8n

Variables de entorno del servicio de n8n: ver `.env.ejemplo`. Las más importantes:

- `GOOGLE_APPLICATION_CREDENTIALS` (paso 2).
- `GENERIC_TIMEZONE=America/Argentina/Buenos_Aires`.
- `NODES_EXCLUDE=[]` — en n8n 2.x el nodo **Execute Command** viene
  deshabilitado y la automatización lo necesita.
- `N8N_RESTRICT_FILE_ACCESS_TO=/opt/observatorio/n8n-files`.

Si n8n corre en Docker: montar `/opt/observatorio` dentro del contenedor con
la misma ruta y usar una imagen que tenga Python 3 (o instalarlo en ella).

### 4. Importar los workflows

Si usan las rutas del ejemplo, los workflows ya están listos en
`n8n/importar_linux/`. Si usan otras rutas:

```bash
cd /opt/observatorio/automatizacion_indicadores/n8n
python3 configurar_n8n.py --so linux \
    --raiz /opt/observatorio/automatizacion_indicadores \
    --python /opt/observatorio/venv/bin/python \
    --n8n-files /opt/observatorio/n8n-files
```

Importar **con la línea de comandos** (conserva los IDs, así el principal
encuentra al sub‑workflow y al de errores sin reconfigurar nada):

```bash
n8n import:workflow --separate --input=/opt/observatorio/automatizacion_indicadores/n8n/importar_linux/
```

Si se importan desde la pantalla de n8n ("Import from file"), n8n les asigna
IDs nuevos y hay que volver a elegir, en el principal, el workflow del nodo
**Call 'Actualizar Base del Observatorio'** y, en *Settings → Error workflow*
de los dos primeros, el workflow **AVISO DE ERRORES - DIAGNÓSTICO MEJORADO**.

### 5. Credenciales dentro de n8n (no viajan en el paquete)

| Credencial (nombre exacto) | Tipo | Usada en |
|---|---|---|
| `Telegram account` | Telegram API (token del bot del Observatorio) | Aviso al grupo y avisos de error |
| `Google Sheets account` | Google Sheets OAuth2 (cuenta con edición del Sheet) | Sub‑workflow Actualizar Base |

Después de crearlas, abrir cada nodo de Telegram / Google Sheets y elegir la
credencial. Chat del grupo: `-5396908101`. El workflow de errores también
avisa a un chat personal (`8829533357`); cámbienlo o desactiven ese nodo si
corresponde.

### 6. Activar y probar

1. Activar los 3 workflows (el principal y el de errores deben quedar activos).
2. Prueba puntual: en el principal, nodo **Elegir indicador puntual**, poner
   una fuente (por ejemplo `iaraf`) y ejecutar desde **Ejecución manual**.
   Valores posibles: `generales`, `complejos`, `productivos` o una fuente
   (`eph, pbi, iaraf, atp, canasta, sipa, ipi, cammesa, dnrpa, ieric, ipc,
   internet, transferencias_fiscales, srt, combustibles, emae, salarios, icc,
   supermercados, pobreza, comercializacion, pecuaria, estimaciones`).
3. Revisar que llegue el mensaje de Telegram con "X de Y datos en el Sheet".

## Instalación en un servidor Windows

Funciona igual que en la PC actual. Generar los workflows con:

```
python configurar_n8n.py --so windows --raiz "D:\observatorio\automatizacion_indicadores" --python "C:\ruta\python.exe" --n8n-files "C:\Users\<usuario>\.n8n-files"
```

y usar la carpeta `n8n\importar_windows\`. En Windows con Microsoft Excel
instalado, además se recalcula el Excel local al final de cada corrida.

## Notas técnicas

- **Sin Excel en Linux:** el recálculo con Excel se omite automáticamente.
  No afecta al Sheet: la comparación se hace sobre las fórmulas y Google
  Sheets calcula sus propias fórmulas.
- **Combustibles (CECHA):** la tabla viene como imagen; en Linux se lee con
  RapidOCR (incluido en `requirements.txt`).
- **Correcciones únicas** (`ieric_correccion`, `correccion_transferencias`,
  `limpieza_motos`, `srt_historico`) solo corren si se piden a mano en *Elegir indicador
  puntual*; la revisión diaria no las repite.
- **Protecciones:** nunca se pisa un dato cargado a mano distinto del oficial
  (se avisa en Telegram), no se escribe en una columna que no existe, y un
  0 o celda vacía en el Sheet nunca borra datos.
- **Control de BASE DE DATOS** (`scripts/comun/control_base.py`): al inicio de
  cada revisión busca cálculos o valores en meses futuros o fuera de la tabla
  (columnas sin encabezado, filas sin período): guarda una copia en
  `reportes/control_base/`, los borra y avisa por Telegram. Dentro de la tabla,
  una fórmula que mira meses que no corresponden (más de 12 atrás o más de 1
  adelante) u otra pestaña se avisa, sin tocarla. El equipo puede seguir
  cargando valores y completando los cálculos de la tabla; la hoja no se
  bloquea (`PROTEGER = False`; con `True` solo editan la automatización y
  `EDITORES_BASE`).
- `scripts/reportes/productivos/pecuaria/estado_faena_formosa.json` guarda el
  acumulado de faena de Formosa del último mes; **no borrarlo** (con él se
  calcula el dato mensual del mes siguiente).

## Limpieza realizada para este paquete

Se quitaron copias viejas y archivos sin uso (siguen en la PC original):

- Copias de respaldo: `ejecutar_atp.antes_atp.py`, `finalizar_excel.antes_atp.py`,
  `ejecutar_dnrpa.antes_reconsulta_….py`.
- `scripts/eph/` (versión vieja; la EPH la hace `scripts/mercado_trabajo/ejecutar_eph.py`).
- `scripts/productivos/pecuaria/ejecutar_productivos.py` (copia vieja del ejecutor de productivos).
- `ejecutar_puntual.py` y `corregir_programas.py` (ya no los usa n8n).
- `config/` (archivos vacíos), `.env` (no lo usaba ningún programa),
  `documentacion/codigos_n8n` (copias del código que ya está dentro de los workflows).
- Datos generados: logs, respaldos del Excel, PDFs/Excel descargados y
  reportes históricos (se dejó solo el último reporte de cada fuente).
- La carpeta `scripts/ATP` pasó a `scripts/atp` (Linux distingue mayúsculas).
