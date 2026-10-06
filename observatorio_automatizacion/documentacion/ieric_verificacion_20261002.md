# IERIC — verificación contra informes oficiales (2/10/2026)

Fuente: Informes de Coyuntura IERIC (https://www.ieric.org.ar/estadistica/informes-de-coyuntura/?AAAA), Cuadro VI (puestos y salario por provincia) y Cuadro VIII (salario real por provincia). Formosa = F, Total país = N. Se revisaron los 69 informes de nov-2020 a jul-2026.

## Criterio
- Cada mes se toma del informe que lo publica por primera vez.
- IERIC revisa meses anteriores en informes posteriores; esas diferencias (sobre todo puestos Formosa y 2021-2022) NO se tocaron.
- SIN estacionalidad: igual al salario con estacionalidad, salvo junio y diciembre (IERIC: "El salario correspondiente a Diciembre y Junio incluye el medio aguinaldo") = (mes anterior + mes siguiente)/2, como fórmula. Todo el histórico ya cumplía la regla salvo jun-2026.

## Corrección única (fuente manual `ieric_correccion`, archivo scripts/ieric/correccion_20261002.json)
- Salario real i.a.: nov-20 y dic-20 (vacíos), feb-21 N (-6 → -16,2), dic-24 F (85 → 8,5), ene-25 N, feb-25 N, oct-25 F/N, nov-25 F/N, jun-26 F, jul-26 F (26,2 → 1,1) y N.
- Var. i.a. salario nominal: nov-20 y dic-20 (vacíos), may-26 F/N (tenían la inflación en lugar de la variación).
- Puestos N con dígitos mal tipeados: abr-21, may-21, mar-23, abr-23 (422.xxx → 442.xxx), ago-24.
- Salario N abr-26: 1.360.132,8 → 1.447.712,6.
- SIN estacionalidad jun-26 F/N: fórmula que quita el aguinaldo.

## Cambios en la automatización diaria (ejecutar_ieric.py)
- Carga también la Var. i.a. del salario real (Cuadro VIII), con control cruzado del salario.
- Completa SIN estacionalidad; al llegar julio/enero completa junio/diciembre con la fórmula.
- Tolera el título "Puestos de trabajo de trabajo registrados…" (errata en informes 244-246).
