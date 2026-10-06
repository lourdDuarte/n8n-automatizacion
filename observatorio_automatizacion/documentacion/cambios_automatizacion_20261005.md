# Automatización Observatorio — cambios del 5/10/2026

## Horarios
Revisión diaria (n8n, nodo de horario existente): 9:30, 11:00, 16:30 y 18:30. En la PC solo corre si está prendida con n8n abierto; en el servidor institucional corre siempre.

## El Sheet manda (Excel local = copia automática)
Al inicio de cada revisión de generales, `espejar_desde_sheet` (scripts/comun/sheet_directo.py) copia al Excel local los valores numéricos del Sheet (hoja BASE DE DATOS), por encabezado y período. No toca fórmulas, no borra datos locales si el Sheet está vacío o en 0. Primera corrida: 481 celdas sincronizadas.

## Aviso de Telegram
Por indicador: "X de Y datos en el Sheet" y, si falta alguno, el motivo (dato manual distinto de la fuente, columna inexistente, error). Resumen arriba: revisados / nuevos / sin novedades / pendientes / con error.

## Correcciones aplicadas
- Transferencias IARAF, acumulado Nación 2026: la fórmula no se reiniciaba en enero (ene–ago inflados). Se corrigió la fórmula de enero (XP86 = monto de enero); feb–ago se recalcularon; septiembre 58.092.792 coincide con IARAF. Fuente manual: `correccion_transferencias`.
- Comercialización septiembre 2026: el Sheet tenía copiados los precios de agosto; quedaron los oficiales (maíz 293.319,27; sorgo 281.978,16; soja 558.850; trigo 346.366,68; FOB trigo 260).
- Pecuaria agosto 2026: datos con el formato histórico de cada columna (consumo "41,15"; faena porcina/huevos/aves en miles con 3 decimales; carne aviar en miles de tn: 187).

## Nuevo: Producción agrícola por cultivo (Estimaciones MAGyP)
scripts/productivos/estimaciones/estimaciones.py — dataset oficial datosestimaciones.magyp.gob.ar; Formosa y total país por cultivo y campaña desde 2020/21; corrige revisiones, agrega campañas nuevas (Girasol y Sorgo 2025/26), copia la fórmula de % de participación del equipo. Corregidos: Cártamo 2025/26 producción Formosa 1.223 → 13.481; Sorgo 2024/25 sup. sembrada 51.100 → 5.100.

## EMAE
Se cargan también los 16 "Números índice, base 2004=100" por sector (hoja "Tabla Letras" del INDEC), con control contra las variaciones publicadas.

## Faena bovina Nación (resuelto)
Fuente según la guía del equipo: SAGyP → Bovinos → Información sectorial → Tablero de faena bovina (FAENA, "Faena (cabezas)" del mes). Agosto 2026 = 1.011.334, igual a lo cargado. La automatización ahora toma ese valor del tablero (respaldo: planilla de faena). Los datos son provisorios y SAGyP los revisa (ej. abril: 960.871 en el informe de abril).
Nota: el consumo per cápita de la columna del Sheet es el mensual (planilla de indicadores), no el promedio móvil del tablero de consumo; se mantiene el criterio histórico.

## Migración al servidor
Ver LEEME.md en la raíz del paquete.
