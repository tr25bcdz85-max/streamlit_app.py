"""Conciliación CxP — Banco Horizonte (Streamlit, versión web desplegable).

Misma lógica R1-R9 que la app Tkinter: usa app/motor_cxp.py y el
pipeline de 4 agentes de app/agentes_cxp.py.

Ejecución local:
    streamlit run streamlit_app.py
Despliegue: Streamlit Community Cloud (ver requirements.txt).
"""

import io
import tempfile

import pandas as pd
import streamlit as st

try:  # estructura con carpeta app/ (repo local)
    from app.agentes_cxp import (
        TEMPERATURA,
        agente_auditor,
        agente_conciliador,
        agente_detector_duplicados,
        agente_extractor,
        hojas_del_excel,
    )
    from app.motor_cxp import (
        CORTE_DEFAULT,
        ESTADOS,
        HOJA_FACTURAS_DEFAULT,
        HOJA_PAGOS_DEFAULT,
        TODOS,
        TOLERANCIA_DEFAULT,
        DatosInvalidosError,
        crc,
        exportar_excel,
        exportar_pdf,
        fecha_corta,
        texto_resumen_correo,
        totales_clave,
    )
except ImportError:  # archivos sueltos en la raíz (upload web sin carpetas)
    from agentes_cxp import (
        TEMPERATURA,
        agente_auditor,
        agente_conciliador,
        agente_detector_duplicados,
        agente_extractor,
        hojas_del_excel,
    )
    from motor_cxp import (
        CORTE_DEFAULT,
        ESTADOS,
        HOJA_FACTURAS_DEFAULT,
        HOJA_PAGOS_DEFAULT,
        TODOS,
        TOLERANCIA_DEFAULT,
        DatosInvalidosError,
        crc,
        exportar_excel,
        exportar_pdf,
        fecha_corta,
        texto_resumen_correo,
        totales_clave,
    )

st.set_page_config(page_title="Conciliación CxP — Banco Horizonte",
                   layout="wide")
st.title("Banco Horizonte — Conciliación de Cuentas por Pagar")
st.caption("Pipeline de 4 agentes determinísticos "
           f"(temperatura {TEMPERATURA}): Extractor → Conciliador → "
           "Auditor → DetectorDuplicados.")

# --------------------------------------------------------------------------
# Agente 1 — Extractor
# --------------------------------------------------------------------------
st.header("1 · Extractor: cargar Excel")
col_a, col_b = st.columns(2)
with col_a:
    up_fact = st.file_uploader("Facturas (Excel)", type=["xlsx", "xls"],
                               key="fact")
with col_b:
    up_pagos = st.file_uploader("Pagos (Excel)", type=["xlsx", "xls"],
                                key="pagos")

if not (up_fact and up_pagos):
    st.info("Sube los dos archivos para conciliar. "
            "Puedes usar insumos/Facturas_Proveedores.xlsx y "
            "insumos/Pagos_Banco.xlsx del repositorio.")
    st.stop()

bytes_fact, bytes_pagos = up_fact.getvalue(), up_pagos.getvalue()
hojas_fact = hojas_del_excel(io.BytesIO(bytes_fact))
hojas_pag = hojas_del_excel(io.BytesIO(bytes_pagos))
col_a, col_b = st.columns(2)
with col_a:
    hoja_fact = st.selectbox("Hoja de facturas", hojas_fact,
                             index=hojas_fact.index(HOJA_FACTURAS_DEFAULT)
                             if HOJA_FACTURAS_DEFAULT in hojas_fact else 0)
with col_b:
    hoja_pag = st.selectbox("Hoja de pagos", hojas_pag,
                            index=hojas_pag.index(HOJA_PAGOS_DEFAULT)
                            if HOJA_PAGOS_DEFAULT in hojas_pag else 0)

try:
    ext = agente_extractor(io.BytesIO(bytes_fact), hoja_fact,
                           io.BytesIO(bytes_pagos), hoja_pag)
except DatosInvalidosError as exc:
    st.error(f"Datos inválidos: {exc}")
    st.stop()

st.success(f"Extractor: {ext['filas_facturas']} facturas + "
           f"{ext['filas_pagos']} pagos "
           f"({ext['limpieza_facturas'] + ext['limpieza_pagos']} "
           "con limpieza R1).")

# --------------------------------------------------------------------------
# Parámetros (re-concilia solo al cambiarlos)
# --------------------------------------------------------------------------
st.header("2 · Parámetros")
c1, c2, c3 = st.columns(3)
with c1:
    corte_txt = st.text_input("Fecha de corte (AAAA-MM-DD)",
                              value=CORTE_DEFAULT)
with c2:
    tolerancia = st.number_input("Tolerancia en colones (₡)",
                                 min_value=0, value=TOLERANCIA_DEFAULT,
                                 step=1)
with c3:
    proveedores = sorted(ext["df_facturas"]["Proveedor"].dropna()
                         .unique().tolist())
    proveedor = st.selectbox("Proveedor", [TODOS] + proveedores)

try:
    corte = pd.Timestamp(corte_txt).date()
except ValueError:
    st.error("Fecha de corte inválida. Usa el formato AAAA-MM-DD.")
    st.stop()

# --------------------------------------------------------------------------
# Agente 2 — Conciliador
# --------------------------------------------------------------------------
con = agente_conciliador(ext["df_facturas"], ext["df_pagos"], corte,
                         int(tolerancia))
df_full, df_sin, resumen, tot = (con["df_conciliado"],
                                 con["df_sin_factura"], con["resumen"],
                                 con["totales"])
df = (df_full if proveedor == TODOS else
      df_full[df_full["Proveedor"] == proveedor])

st.header("3 · Conciliador: resumen por estado")
cols = st.columns(len(ESTADOS))
for col, estado in zip(cols, ESTADOS):
    sub = df[df["Estado"] == estado]
    col.metric(estado, f"{len(sub)} facturas", crc(sub["Monto_CRC"].sum()))
st.write(f"Total vencido: **{crc(tot['vencido'])}** · "
         f"Total pendiente: **{crc(tot['pendiente'])}** · "
         f"Pagado de más por duplicados: **{crc(tot['duplicado_de_mas'])}** · "
         f"Pagos sin factura: **{tot['sin_factura_conteo']} "
         f"({crc(tot['sin_factura_monto'])})**")

st.subheader("Monto facturado por Estado")
graf = pd.DataFrame(
    {"Estado": ESTADOS,
     "Monto": [float(df.loc[df["Estado"] == e, "Monto_CRC"].sum())
               for e in ESTADOS]}).set_index("Estado")
st.bar_chart(graf)

# --------------------------------------------------------------------------
# Agente 3 — Auditor
# --------------------------------------------------------------------------
st.header("4 · Auditor: alertas")
aud = agente_auditor(df, df_sin, resumen)
st.write(f"**{aud['n_alertas']}** facturas no PAGADA + "
         f"**{aud['n_pagos_sin_factura']}** pagos sin factura (R9).")
tabla = pd.DataFrame([{
    "Num_Factura": str(f["Num_Factura"]),
    "Proveedor": str(f["Proveedor"]),
    "Vence": fecha_corta(f["Fecha_Vencimiento"]),
    "Facturado": crc(f["Monto_CRC"]),
    "Pagado": crc(f["Monto_Pagado"]),
    "Diferencia": crc(f["Diferencia"]) if f["Num_Pagos"] else "—",
    "Estado": f["Estado"],
} for _, f in aud["df_alertas"].iterrows()])
for _, p in aud["df_pagos_sin_factura"].iterrows():
    tabla.loc[len(tabla)] = [str(p["Num_Factura"]),
                             f"{p['Referencia_Pago']} — {p['Cuenta_Bancaria']}",
                             fecha_corta(p["Fecha_Pago"]), "—",
                             crc(p["Monto_Pagado_CRC"]), "—",
                             "PAGO SIN FACTURA"]
st.dataframe(tabla, use_container_width=True)

# --------------------------------------------------------------------------
# Agente 4 — Detector de duplicados
# --------------------------------------------------------------------------
st.header("5 · Detector de duplicados (R6)")
dup = agente_detector_duplicados(df_full, ext["df_pagos"])
st.write(f"**{dup['n_facturas_duplicadas']}** facturas duplicadas · "
         f"**{crc(dup['total_pagado_de_mas'])}** pagados de más.")
for d in dup["detalle"]:
    with st.expander(f"{d['factura']} — {d['proveedor']} "
                     f"(de más: {crc(d['pagado_de_mas'])})"):
        st.table(pd.DataFrame(d["pagos"]))

# --------------------------------------------------------------------------
# Exportar + copiar resumen
# --------------------------------------------------------------------------
st.header("6 · Exportar")
res_filtro = {e: {"conteo": int((df["Estado"] == e).sum()),
                  "monto": float(df.loc[df["Estado"] == e,
                                        "Monto_CRC"].sum())}
              for e in ESTADOS}
sufijo = pd.Timestamp(corte).strftime("%Y-%m-%d")
with tempfile.TemporaryDirectory() as tmp:
    xlsx_path = f"{tmp}/Conciliacion_CxP_{sufijo}.xlsx"
    pdf_path = f"{tmp}/Reporte_Conciliacion_{sufijo}.pdf"
    exportar_excel(xlsx_path, df_full, df_sin, con["resumen"])
    exportar_pdf(pdf_path, df_full, df_sin, con["resumen"], corte,
                 int(tolerancia), proveedor)
    with open(xlsx_path, "rb") as fh:
        datos_xlsx = fh.read()
    with open(pdf_path, "rb") as fh:
        datos_pdf = fh.read()
c1, c2 = st.columns(2)
with c1:
    st.download_button("Descargar Excel",
                       datos_xlsx,
                       file_name=f"Conciliacion_CxP_{sufijo}.xlsx")
with c2:
    st.download_button("Descargar PDF",
                       datos_pdf,
                       file_name=f"Reporte_Conciliacion_{sufijo}.pdf")

texto = texto_resumen_correo(corte, int(tolerancia), proveedor, res_filtro,
                             totales_clave(df, df_sin),
                             int((df["Estado"] != "PAGADA").sum()))
st.subheader("Resumen listo para correo")
st.code(texto)
