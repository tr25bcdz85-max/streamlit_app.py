"""Motor de conciliación CxP — Banco Horizonte (módulo puro, sin UI).

Contiene las reglas R1-R9 del AGENTS.md: carga y validación de Excel,
cruce facturas vs pagos, resúmenes y exportación a Excel/PDF.
Lo usan tanto la app Tkinter (app/conciliacion_cxp_app.py) como la
app Streamlit (streamlit_app.py): una sola fuente de verdad.

Dependencias: pandas, openpyxl, reportlab.
"""

import os
from datetime import date, datetime

import pandas as pd

# --------------------------------------------------------------------------
# Constantes
# --------------------------------------------------------------------------

AZUL_OSCURO = "#1F3A5F"

CORTE_DEFAULT = "2026-08-31"
TOLERANCIA_DEFAULT = 1  # R4 usa 1 colón; la UI permite cambiarlo.

HOJA_FACTURAS_DEFAULT = "Facturas"
HOJA_PAGOS_DEFAULT = "Pagos"

COLS_FACTURAS = [
    "Num_Factura",
    "Proveedor",
    "Fecha_Factura",
    "Fecha_Vencimiento",
    "Monto_CRC",
    "Centro_Costo",
    "Categoria",
]
COLS_PAGOS = [
    "Referencia_Pago",
    "Fecha_Pago",
    "Num_Factura",
    "Monto_Pagado_CRC",
    "Cuenta_Bancaria",
]

COL_LIMPIO = "Num_Factura_Limpio"
COL_FLAG_LIMPIEZA = "Requirio_Limpieza"

ESTADOS = [
    "PAGADA",
    "PAGADA CON DIFERENCIA",
    "PAGO DUPLICADO",
    "VENCIDA",
    "PENDIENTE",
]

DIR_REPORTES = "reportes"
XLSX_BASE = "Conciliacion_CxP"
PDF_BASE = "Reporte_Conciliacion"

TODOS = "Todos"


class DatosInvalidosError(ValueError):
    """Error de validación de los Excel de entrada (bloquea CONCILIAR)."""


def resolver_dir_reportes():
    """Carpeta de salida que sí se pueda escribir.

    Al abrir el .app con doble clic, el directorio de trabajo es /
    (solo lectura) y ./reportes falla. Se intenta ./reportes primero
    (respeta AGENTS.md al correr desde terminal) y si no se puede,
    ~/Documents/Conciliacion_CxP/reportes.
    """
    candidatos = [
        os.path.join(os.getcwd(), DIR_REPORTES),
        os.path.join(os.path.expanduser("~"), "Documents",
                     "Conciliacion_CxP", DIR_REPORTES),
    ]
    ultimo_error = None
    for ruta in candidatos:
        try:
            os.makedirs(ruta, exist_ok=True)
            # exist_ok no basta: verifica escritura real.
            prueba = os.path.join(ruta, ".prueba_escritura.tmp")
            with open(prueba, "w") as fh:
                fh.write("ok")
            os.remove(prueba)
            return ruta
        except OSError as exc:
            ultimo_error = exc
    raise DatosInvalidosError(
        f"No se pudo crear ninguna carpeta de salida {candidatos}: "
        f"{ultimo_error}")


def nombres_export(corte, base):
    """Nombres con fecha de corte: Conciliacion_CxP_2026-08-31.xlsx."""
    sufijo = pd.Timestamp(corte).strftime("%Y-%m-%d")
    return (os.path.join(base, f"{XLSX_BASE}_{sufijo}.xlsx"),
            os.path.join(base, f"{PDF_BASE}_{sufijo}.pdf"))


# --------------------------------------------------------------------------
# Utilidades de formato
# --------------------------------------------------------------------------

def crc(valor):
    """Formato colones con separador de miles: ₡1.362.800."""
    return f"₡{valor:,.0f}".replace(",", ".")


def fecha_corta(valor):
    """YYYY-MM-DD para datetimes/Timestamps; '—' si es nulo."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return "—"
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return valor.strftime("%Y-%m-%d")
    return str(valor)


# --------------------------------------------------------------------------
# R1: limpieza
# --------------------------------------------------------------------------

def limpiar_num_factura(valor):
    """Normaliza Num_Factura: str, sin espacios extremos, en mayúsculas."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return str(valor).strip().upper()


# --------------------------------------------------------------------------
# Carga y validación de Excel
# --------------------------------------------------------------------------

def hojas_del_excel(path):
    """Lista las hojas de un .xlsx (para permitir elegirla al cargar)."""
    return pd.ExcelFile(path, engine="openpyxl").sheet_names


def _validar_columnas(df, requeridas, etiqueta):
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise DatosInvalidosError(
            f"{etiqueta}: faltan columnas {faltantes}. "
            f"Se esperaban: {requeridas}."
        )


def cargar_facturas(path, hoja):
    """Carga y valida facturas. Devuelve DataFrame con R1 aplicada."""
    try:
        df = pd.read_excel(path, sheet_name=hoja, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001 - se reporta tal cual al usuario
        raise DatosInvalidosError(f"No se pudo leer la hoja '{hoja}': {exc}")
    _validar_columnas(df, COLS_FACTURAS, "Facturas")
    df = df.copy()

    df[COL_LIMPIO] = df["Num_Factura"].map(limpiar_num_factura)
    df[COL_FLAG_LIMPIEZA] = df["Num_Factura"].astype(str) != df[COL_LIMPIO]

    vacios = df.index[df[COL_LIMPIO] == ""].tolist()
    if vacios:
        raise DatosInvalidosError(
            f"Facturas: Num_Factura vacío en filas Excel {[i + 2 for i in vacios]}."
        )
    dup = df[df.duplicated(COL_LIMPIO, keep=False)][COL_LIMPIO].unique().tolist()
    if len(dup):
        raise DatosInvalidosError(
            f"Facturas: Num_Factura duplicado tras limpieza (R1): {dup}."
        )

    df["Fecha_Vencimiento"] = pd.to_datetime(df["Fecha_Vencimiento"], errors="coerce")
    malas = df.index[df["Fecha_Vencimiento"].isna()].tolist()
    if malas:
        raise DatosInvalidosError(
            "Facturas: Fecha_Vencimiento inválida (ni fecha ni texto "
            f"parseable) en filas Excel {[i + 2 for i in malas]}."
        )
    df["Fecha_Factura"] = pd.to_datetime(df["Fecha_Factura"], errors="coerce")

    df["Monto_CRC"] = pd.to_numeric(df["Monto_CRC"], errors="coerce")
    malos = df.index[df["Monto_CRC"].isna()].tolist()
    if malos:
        raise DatosInvalidosError(
            f"Facturas: Monto_CRC vacío o no numérico en filas Excel {[i + 2 for i in malos]}."
        )
    return df


def cargar_pagos(path, hoja):
    """Carga y valida pagos. Devuelve DataFrame con R1 aplicada."""
    try:
        df = pd.read_excel(path, sheet_name=hoja, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001 - se reporta tal cual al usuario
        raise DatosInvalidosError(f"No se pudo leer la hoja '{hoja}': {exc}")
    _validar_columnas(df, COLS_PAGOS, "Pagos")
    df = df.copy()

    df[COL_LIMPIO] = df["Num_Factura"].map(limpiar_num_factura)
    df[COL_FLAG_LIMPIEZA] = df["Num_Factura"].astype(str) != df[COL_LIMPIO]

    df["Monto_Pagado_CRC"] = pd.to_numeric(df["Monto_Pagado_CRC"], errors="coerce")
    malos = df.index[df["Monto_Pagado_CRC"].isna()].tolist()
    if malos:
        raise DatosInvalidosError(
            f"Pagos: Monto_Pagado_CRC vacío o no numérico en filas Excel {[i + 2 for i in malos]}."
        )
    df["Fecha_Pago"] = pd.to_datetime(df["Fecha_Pago"], errors="coerce")
    return df


# --------------------------------------------------------------------------
# Motor de conciliación (R2-R9)
# --------------------------------------------------------------------------

def conciliar(df_fact, df_pag, corte, tolerancia=TOLERANCIA_DEFAULT):
    """Cruza facturas vs pagos y clasifica cada factura.

    Devuelve (df_conciliado, df_pagos_sin_factura, resumen).
    df_conciliado trae las columnas originales + Num_Pagos, Monto_Pagado,
    Diferencia, Estado (+ columnas de limpieza R1).
    """
    # R3: contar pagos y sumar monto por factura (Num_Factura limpio).
    pagos_validos = df_pag[df_pag[COL_LIMPIO] != ""]
    agg = (
        pagos_validos.groupby(COL_LIMPIO)["Monto_Pagado_CRC"]
        .agg(Num_Pagos="size", Monto_Pagado="sum")
        .reset_index()
    )
    df = df_fact.merge(agg, on=COL_LIMPIO, how="left")
    df["Num_Pagos"] = df["Num_Pagos"].fillna(0).astype(int)
    df["Monto_Pagado"] = df["Monto_Pagado"].fillna(0)
    df["Diferencia"] = df["Monto_Pagado"] - df["Monto_CRC"]

    corte_ts = pd.Timestamp(corte)

    def clasificar(fila):
        # R4/R5: un pago, tolerancia configurable (default 1 colón).
        if fila["Num_Pagos"] == 1:
            if abs(fila["Diferencia"]) <= tolerancia:
                return "PAGADA"
            return "PAGADA CON DIFERENCIA"
        # R6: dos o más pagos (Diferencia = monto pagado de más).
        if fila["Num_Pagos"] >= 2:
            return "PAGO DUPLICADO"
        # R7/R8: sin pagos, según vencimiento vs corte.
        if fila["Fecha_Vencimiento"] < corte_ts:
            return "VENCIDA"
        return "PENDIENTE"

    df["Estado"] = df.apply(clasificar, axis=1)

    # R9: pagos cuyo Num_Factura limpio no existe en facturas (incluye vacíos).
    nums_fact = set(df_fact[COL_LIMPIO])
    df_sin = df_pag[~df_pag[COL_LIMPIO].isin(nums_fact)].copy()

    resumen = {}
    for estado in ESTADOS:
        sub = df[df["Estado"] == estado]
        resumen[estado] = {
            "conteo": int(len(sub)),
            "monto": float(sub["Monto_CRC"].sum()),
        }
    return df, df_sin, resumen


def totales_clave(df_conciliado, df_sin_factura):
    """Total vencido, pendiente, pagado de más y pagos sin factura."""
    venc = df_conciliado.loc[df_conciliado["Estado"] == "VENCIDA", "Monto_CRC"].sum()
    pend = df_conciliado.loc[df_conciliado["Estado"] == "PENDIENTE", "Monto_CRC"].sum()
    dupl = df_conciliado.loc[
        df_conciliado["Estado"] == "PAGO DUPLICADO", "Diferencia"].sum()
    return {
        "vencido": float(venc),
        "pendiente": float(pend),
        "duplicado_de_mas": float(dupl),
        "sin_factura_conteo": int(len(df_sin_factura)),
        "sin_factura_monto": float(df_sin_factura["Monto_Pagado_CRC"].sum())
        if len(df_sin_factura) else 0.0,
    }


# --------------------------------------------------------------------------
# Resúmenes
# --------------------------------------------------------------------------

def resumen_por_proveedor(df_conciliado):
    """Monto facturado, pagado, vencido y pendiente por proveedor."""
    filas = []
    for prov, sub in df_conciliado.groupby("Proveedor"):
        filas.append({
            "Proveedor": prov,
            "Monto_Facturado": float(sub["Monto_CRC"].sum()),
            "Monto_Pagado": float(sub["Monto_Pagado"].sum()),
            "Monto_Vencido": float(
                sub.loc[sub["Estado"] == "VENCIDA", "Monto_CRC"].sum()),
            "Monto_Pendiente": float(
                sub.loc[sub["Estado"] == "PENDIENTE", "Monto_CRC"].sum()),
        })
    df = pd.DataFrame(filas)
    return df.sort_values("Monto_Facturado", ascending=False).reset_index(drop=True)


def texto_resumen_correo(corte, tolerancia, proveedor, resumen, tot, n_alertas):
    """Texto del resumen listo para pegar en un correo."""
    lineas = [
        "Banco Horizonte — Conciliación de Cuentas por Pagar",
        f"Fecha de corte: {corte} | Tolerancia: {crc(tolerancia)} | "
        f"Proveedor: {proveedor}",
        "",
        "Resumen por estado (monto facturado):",
    ]
    for estado in ESTADOS:
        r = resumen[estado]
        lineas.append(f"- {estado}: {r['conteo']} facturas — {crc(r['monto'])}")
    lineas += [
        "",
        f"Total vencido: {crc(tot['vencido'])}",
        f"Total pendiente: {crc(tot['pendiente'])}",
        f"Pagado de más por duplicados: {crc(tot['duplicado_de_mas'])}",
        f"Pagos sin factura: {tot['sin_factura_conteo']} "
        f"({crc(tot['sin_factura_monto'])})",
        f"Facturas en alerta: {n_alertas}",
    ]
    return "\n".join(lineas)


# --------------------------------------------------------------------------
# Exportación a Excel
# --------------------------------------------------------------------------

def exportar_excel(path_xlsx, df_conciliado, df_sin_factura, resumen):
    """Genera Conciliacion_CxP_<corte>.xlsx con 4 hojas."""
    cols_conc = (COLS_FACTURAS + [COL_LIMPIO, "Num_Pagos", "Monto_Pagado",
                                  "Diferencia", "Estado"])
    df_c = df_conciliado[cols_conc].copy()
    df_r = pd.DataFrame(
        [{"Estado": e, "Num_Facturas": resumen[e]["conteo"],
          "Monto_Facturado": resumen[e]["monto"]} for e in ESTADOS])
    df_prov = resumen_por_proveedor(df_conciliado)
    with pd.ExcelWriter(path_xlsx, engine="openpyxl") as writer:
        df_c.to_excel(writer, sheet_name="Facturas_Conciliadas", index=False)
        df_sin_factura.to_excel(writer, sheet_name="Pagos_Sin_Factura", index=False)
        df_r.to_excel(writer, sheet_name="Resumen", index=False)
        df_prov.to_excel(writer, sheet_name="Por_Proveedor", index=False)


# --------------------------------------------------------------------------
# Exportación a PDF (reportlab)
# --------------------------------------------------------------------------

def exportar_pdf(path_pdf, df_conciliado, df_sin_factura, resumen,
                 corte, tolerancia, proveedor_filtro=TODOS):
    """Genera Reporte_Conciliacion.pdf con resumen ejecutivo y alertas."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    doc = SimpleDocTemplate(path_pdf, pagesize=letter,
                            title="Reporte de Conciliación CxP")
    estilos = getSampleStyleSheet()
    partes = [
        Paragraph("Banco Horizonte - Conciliación de Cuentas por Pagar",
                  estilos["Title"]),
        Paragraph(f"Fecha de corte: {corte} &nbsp;&nbsp;|&nbsp;&nbsp; "
                  f"Fecha de generación: {date.today().isoformat()} "
                  f"&nbsp;&nbsp;|&nbsp;&nbsp; Tolerancia: {crc(tolerancia)} "
                  f"&nbsp;&nbsp;|&nbsp;&nbsp; Proveedor: {proveedor_filtro}",
                  estilos["Normal"]),
        Spacer(1, 12),
        Paragraph("Resumen por estado (monto facturado)", estilos["Heading2"]),
    ]

    tabla_res = [["Estado", "Facturas", "Monto"]]
    for estado in ESTADOS:
        r = resumen[estado]
        tabla_res.append([estado, str(r["conteo"]), crc(r["monto"])])
    t = Table(tabla_res, colWidths=[220, 80, 150])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(AZUL_OSCURO)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
    ]))
    partes += [t, Spacer(1, 12)]

    alertas = df_conciliado[df_conciliado["Estado"] != "PAGADA"].sort_values(
        "Monto_CRC", ascending=False)
    partes.append(Paragraph(f"Alertas: facturas no PAGADA ({len(alertas)})",
                            estilos["Heading2"]))
    tabla_al = [["Factura", "Proveedor", "Vence", "Facturado", "Pagado",
                 "Diferencia", "Estado"]]
    for _, f in alertas.iterrows():
        tabla_al.append([
            str(f["Num_Factura"]),
            str(f["Proveedor"])[:24],
            fecha_corta(f["Fecha_Vencimiento"]),
            crc(f["Monto_CRC"]),
            crc(f["Monto_Pagado"]),
            crc(f["Diferencia"]) if f["Num_Pagos"] else "—",
            f["Estado"],
        ])
    ta = Table(tabla_al, colWidths=[55, 120, 60, 65, 65, 65, 90], repeatRows=1)
    ta.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7a1f1f")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (3, 1), (5, -1), "RIGHT"),
    ]))
    partes += [ta, Spacer(1, 12)]

    partes.append(Paragraph(f"Pagos sin factura — R9 ({len(df_sin_factura)})",
                            estilos["Heading2"]))
    tabla_r9 = [["Referencia", "Num_Factura", "Monto", "Cuenta"]]
    for _, p in df_sin_factura.iterrows():
        tabla_r9.append([
            str(p["Referencia_Pago"]),
            str(p["Num_Factura"]),
            crc(p["Monto_Pagado_CRC"]),
            str(p["Cuenta_Bancaria"]),
        ])
    tr9 = Table(tabla_r9, colWidths=[90, 90, 90, 150], repeatRows=1)
    tr9.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#5b5b5b")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (2, 1), (2, -1), "RIGHT"),
    ]))
    partes.append(tr9)
    doc.build(partes)
