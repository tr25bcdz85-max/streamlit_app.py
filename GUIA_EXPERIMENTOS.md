# Guía de experimentos y reflexión — Conciliación CxP

Qué hace cada agente, qué reglas lo protegen y qué se midió al calibrar.
(Datos de control verificados el 2026-09-26 con corte 2026-08-31 y
tolerancia 1, salvo que se indique otro valor.)

## Experimento 1 · Calibrar la tolerancia (R4/R5)

| Tolerancia | PAGADA | PAGADA CON DIFERENCIA | Qué cambió |
|---|---|---|---|
| 1 (default) | 51 | 2 | FAC-0012 (+₡3.600) y FAC-0034 (−₡208.000) en diferencia |
| 5.000 | 52 | 1 | FAC-0012 pasa a PAGADA; FAC-0034 sigue en diferencia |
| 250.000 | 53 | 0 | Ambas pasan a PAGADA |

Conclusión: la tolerancia solo mueve facturas entre PAGADA y PAGADA CON
DIFERENCIA; ningún otro estado cambia. Decisión: default 1 colón
(estricto); el auditor decide los casos con diferencia.

## Experimento 2 · Mover la fecha de corte (R7/R8)

| Corte | VENCIDA | PENDIENTE | Monto vencido |
|---|---|---|---|
| 2026-08-31 | 3 | 2 | ₡5.097.100 |
| 2026-09-30 | 5 | 0 | ₡6.566.900 |

FAC-0051 (₡1.353.700) y FAC-0058 (₡116.100) pasan de PENDIENTE a VENCIDA.
Solo cambian facturas sin pagos; las pagadas no se mueven.

## Experimento 3 · Filtro por proveedor

Cafetería y Alimentos Sabor: 3 alertas (FAC-0034 diferencia, FAC-0047
duplicado, FAC-0051 pendiente). El resumen se recalcula sobre el filtro;
la exportación siempre sale completa (todas las facturas).

## Experimento 4 · Documentos ficticios (`pruebas_ficticias/`)

| Caso | Contenido | Resultado esperado (= obtenido) |
|---|---|---|
| PRUEBA_01_Duplicado_R6 | 4 facturas, una con 2 pagos | 2 PAGADA, 1 PAGO DUPLICADO (+₡50.000), 1 VENCIDA |
| PRUEBA_02_Diferencia_SinFactura | diferencia + pago huérfano | 1 PAGADA, 1 DIFERENCIA (−₡5.000), 1 PENDIENTE, 1 R9 |
| PRUEBA_03_Limpieza_R1 | minúsculas y espacios | 3 PAGADA (sin R1 quedarían sin pago) |

## Preguntas de reflexión (con respuesta)

**1. ¿Por qué los agentes usan temperatura 0?**
Porque aplican reglas duras: con los mismos datos siempre dan el mismo
resultado. Subir la temperatura solo agregaría variación sin valor; el
criterio (por qué FAC-0034 se pagó con ₡208.000 de menos) queda para la
persona, que ahora tiene tiempo de investigarlo.

**2. ¿Qué error evita la validación estricta del Auditor?**
Que un Excel con columnas distintas o montos vacíos se concilie "a
ciegas" y contamine el cierre. El pipeline se detiene con el número de
fila exacto en vez de adivinar.

**3. ¿Qué parte del flujo sí requeriría criterio (y un agente con
temperatura > 0)?**
Redactar el correo al proveedor cuando el auditor rechaza una factura, o
decidir si una diferencia es error o descuento negociado. Por eso el
cuarto agente elegido fue el Detector de duplicados (regla dura) y no el
redactor: mantiene todo el pipeline auditable.
