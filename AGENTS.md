# AGENTS.md — conciliacion_cxp

App de conciliación mensual de facturas contra pagos — Banco Horizonte (Costa Rica), Contabilidad y Cuentas por Pagar: qué está pagado, qué falta, qué se pagó con monto distinto y qué se pagó dos veces. Lógica R1-R9 en `app/motor_cxp.py` (sin UI); `app/conciliacion_cxp_app.py` es la ventana Tkinter y `streamlit_app.py` la vista web. No duplicar reglas entre UIs.

## Layout y comandos

- Escritorio: `python3 app/conciliacion_cxp_app.py`. Web local: `streamlit run streamlit_app.py`. (`python3` aquí es 3.12; deps: pandas, openpyxl, reportlab, matplotlib, streamlit; Tkinter es stdlib. Deps de despliegue en `requirements.txt`.)
- `insumos/` = entradas read-only. `reportes/` = salidas generadas. `dist/` = artefactos de build. `instalador/` = fuentes del `.pkg` macOS. `pruebas_ficticias/` = 3 casos de prueba (cada uno con hojas `Facturas`+`Pagos`). Nunca escribir salidas junto a ni sobre los `.xlsx` de entrada.
- Sin README, manifests, tests, lint ni CI. No es un repo git; no asumir historial.

## Archivos de entrada (read-only, nunca modificar)

- `insumos/Facturas_Proveedores.xlsx` (hoja `Facturas`, 60 filas): `Num_Factura | Proveedor | Fecha_Factura | Fecha_Vencimiento | Monto_CRC | Centro_Costo | Categoria`.
- `insumos/Pagos_Banco.xlsx` (hoja `Pagos`, 58 filas): `Referencia_Pago | Fecha_Pago | Num_Factura | Monto_Pagado_CRC | Cuenta_Bancaria`.

Clave de cruce: `Num_Factura` (facturas) ↔ `Num_Factura` (pagos). Montos en CRC.

## Reglas de negocio (obligatorias, aplicar al pie de la letra)

- R1 Limpieza: en ambos archivos normalizar `Num_Factura` (quitar espacios al inicio y al final, pasar a mayúsculas) en una columna nueva. Nunca modificar las columnas originales.
- R2 Fecha de corte configurable, por defecto 2026-08-31.
- R3 Cruzar por `Num_Factura` limpio; por factura contar los pagos y sumar el monto pagado.
- R4 PAGADA: exactamente 1 pago y `|pagado - facturado| <= 1` colón.
- R5 PAGADA CON DIFERENCIA: exactamente 1 pago y la diferencia es mayor a 1 colón. Reportar `Diferencia = pagado - facturado`.
- R6 PAGO DUPLICADO: 2 o más pagos. Reportar el monto pagado de más (`total pagado - facturado`).
- R7 VENCIDA: sin pagos y `Fecha_Vencimiento < fecha de corte`.
- R8 PENDIENTE: sin pagos y `Fecha_Vencimiento >= fecha de corte`.
- R9 Pagos sin factura: pagos cuyo `Num_Factura` limpio no existe en facturas. Listarlos aparte como alerta.

## Salidas

- `Conciliacion_CxP_<corte>.xlsx` con hojas `Facturas_Conciliadas` (todas las facturas + `Num_Pagos`, `Monto_Pagado`, `Diferencia`, `Estado`), `Pagos_Sin_Factura`, `Resumen` (conteo y monto por estado) y `Por_Proveedor` (facturado, pagado, vencido y pendiente por proveedor).
- `Reporte_Conciliacion_<corte>.pdf` con el resumen ejecutivo y las alertas.

## Gotchas (verificados en los datos)

- `Num_Factura` no es limpio: hay minúsculas (`fac-0022` en pagos, `fac-0008` en facturas) y espacios (`' FAC-0023 '`). Normalizar según R1 antes de cruzar.
- Las fechas de Excel cargan como datetimes; los montos como ints.
- Los conteos difieren (60 facturas vs 58 pagos) — los casos sin cruce son la tarea de conciliación, no un bug de carga.

## Decisiones de UI confirmadas (no cambiar sin preguntar)

- Tolerancia configurable en la UI (default 1); sobrescribe R4/R5.
- Validación estricta: bloquear CONCILIAR ante fechas/montos inválidos o columnas faltantes.
- Resumen = monto facturado por estado; el filtro de proveedor recalcula el resumen.
- El botón Exportar genera `Conciliacion_CxP_<corte>.xlsx` + `Reporte_Conciliacion_<corte>.pdf` (siempre la conciliación completa, sobrescribe sin preguntar).
- Al cargar cada Excel se puede elegir la hoja (defaults: `Facturas` / `Pagos`).
- Gráfico de barras con monto por estado bajo el resumen; botón Copiar resumen al portapapeles.
- Salida: `./reportes` si se puede escribir; si no (el .app con doble clic arranca en `/` de solo lectura), `~/Documents/Conciliacion_CxP/reportes` — ver `resolver_dir_reportes()` en el código.

## Agentes (pipeline secuencial, ver `app/agentes_cxp.py`)

Orquestación: Extractor → Conciliador → Auditor → DetectorDuplicados
(`ejecutar_pipeline`). Todos con temperatura 0 (reglas duras: mismos
datos, mismo resultado).

- **Extractor** (R1): entradas = 2 Excel + hoja de cada uno; salidas = DataFrames validados con `Num_Factura_Limpio` + conteo de limpieza. Falla ante columnas faltantes, fechas/montos inválidos, `Num_Factura` vacío o duplicado.
- **Conciliador** (R2-R8): entradas = DataFrames + corte + tolerancia; salidas = facturas con `Estado` + resumen por estado + totales clave.
- **Auditor** (R9): entradas = salida del Conciliador; salidas = alertas (no PAGADA) ordenadas por monto + pagos sin factura. Nunca falla: sin alertas devuelve tablas vacías.
- **DetectorDuplicados** (R6, cuarto agente): entradas = conciliadas + pagos; salidas = por factura duplicada, pagos individuales (referencia, fecha, monto, cuenta) + monto de más, ordenado descendente.

## App web (Streamlit, para la presentación)

- Desplegar: Streamlit Community Cloud → New app → repo `conciliacion_cxp`, archivo `streamlit_app.py`. Los Excel se suben por `file_uploader` (se leen por bytes, nunca del disco del servidor).
- Descargas: Excel + PDF con la fecha de corte en el nombre; resumen para correo con `st.code`; gráfico con `st.bar_chart`.

## Instalador macOS (sin administrador)

- `dist/Conciliacion_CxP_Instalador.pkg` (~39 MB): instala en `~/Applications` + acceso directo `~/Desktop/Conciliacion CxP`. Solo dominio de usuario (`CurrentUserHome`); no pide clave; Python va empaquetado (verificado con `otool`: solo frameworks del sistema).
- Probar sin sudo: `installer -pkg dist/Conciliacion_CxP_Instalador.pkg -target CurrentUserHomeDirectory`.
- Reconstruir: 1) PyInstaller `--onefile --windowed --name Conciliacion_CxP app/conciliacion_cxp_app.py`; 2) copiar el `.app` a `instalador/staging/`; 3) `pkgbuild` de `instalador/staging` → `CompApp.pkg` (`/Applications`) y de `instalador/staging-desktop` (symlink relativo a `../Applications/...`) → `CompDesktop.pkg` (`/Desktop`); 4) `productbuild --distribution instalador/Distribution.xml --package-path instalador dist/...`.
- `Distribution.xml` debe salir de `productbuild --synthesize` + dominio de usuario: escrito a mano, `productbuild` no incrusta los componentes (y `xar -C` de este sistema está roto; ensamblar manual no lo lee el Instalador).
- Gotcha: si ya existe otra copia del `.app` en el disco, el instalador la actualiza en su lugar (relocation) y NO crea `~/Applications`. Para prueba limpia, apartar las copias (`instalador/staging/`, `dist/`).

## Instalador Windows

- Script listo: `instalador/windows/Conciliacion_CxP_windows.iss` (Inno Setup, sin admin: instala en `%LOCALAPPDATA%\ConciliacionCxP` + acceso directo en el escritorio).
- El `.exe` debe generarse en una máquina Windows (PyInstaller no hace cross-build, no intentar compilarlo desde macOS): `pyinstaller --onefile --windowed --name Conciliacion_CxP app\conciliacion_cxp_app.py`, copiar el `.exe` a `dist\` junto al `.iss` y compilar con Inno Setup.
- El código no necesita cambios de rutas: `resolver_dir_reportes()` usa `expanduser("~/Documents/...")`, que también resuelve en Windows.
