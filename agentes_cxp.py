"""Agentes del pipeline CxP — Banco Horizonte.

Orquestación secuencial de 4 agentes determinísticos (temperatura 0:
las reglas duras R1-R9 siempre dan el mismo resultado con los mismos
datos, por eso son código y no criterio). Cada agente declara su
CONTRATO (entradas, salidas, reglas que aplica). Ver AGENTS.md.

Orden: Extractor -> Conciliador -> Auditor -> DetectorDuplicados.
"""

from motor_cxp import (
    COL_FLAG_LIMPIEZA,
    COL_LIMPIO,
    TOLERANCIA_DEFAULT,
    DatosInvalidosError,
    cargar_facturas,
    cargar_pagos,
    conciliar,
    hojas_del_excel,
    totales_clave,
)

TEMPERATURA = 0  # Agentes determinísticos: reglas duras, sin muestreo.


def agente_extractor(fuente_facturas, hoja_fact, fuente_pagos, hoja_pagos):
    """Agente 1 — Extractor: carga, valida y aplica limpieza R1."""
    df_fact = cargar_facturas(fuente_facturas, hoja_fact)
    df_pag = cargar_pagos(fuente_pagos, hoja_pagos)
    return {
        "df_facturas": df_fact,
        "df_pagos": df_pag,
        "filas_facturas": len(df_fact),
        "filas_pagos": len(df_pag),
        "limpieza_facturas": int(df_fact[COL_FLAG_LIMPIEZA].sum()),
        "limpieza_pagos": int(df_pag[COL_FLAG_LIMPIEZA].sum()),
    }


agente_extractor.CONTRATO = {
    "entradas": "2 Excel (facturas + pagos) + hoja de cada uno",
    "salidas": "DataFrames validados con Num_Factura_Limpio + conteo de limpieza R1",
    "reglas": ["R1"],
    "temperatura": TEMPERATURA,
    "falla_si": "columnas faltantes, fechas/montos inválidos, Num_Factura vacío o duplicado",
}


def agente_conciliador(df_fact, df_pag, corte, tolerancia=TOLERANCIA_DEFAULT):
    """Agente 2 — Conciliador: cruza y clasifica cada factura (R3-R8)."""
    df_conc, df_sin, resumen = conciliar(df_fact, df_pag, corte, tolerancia)
    return {
        "df_conciliado": df_conc,
        "df_sin_factura": df_sin,
        "resumen": resumen,
        "totales": totales_clave(df_conc, df_sin),
    }


agente_conciliador.CONTRATO = {
    "entradas": "DataFrames del Extractor + fecha de corte + tolerancia",
    "salidas": "Facturas con Estado + resumen por estado + totales clave",
    "reglas": ["R2", "R3", "R4", "R5", "R6", "R7", "R8"],
    "temperatura": TEMPERATURA,
    "falla_si": "corte con formato inválido o tolerancia negativa (lo valida la UI)",
}


def agente_auditor(df_conc, df_sin, resumen):
    """Agente 3 — Auditor: verifica y lista todas las alertas (R9 + no PAGADA)."""
    alertas = df_conc[df_conc["Estado"] != "PAGADA"].sort_values(
        "Monto_CRC", ascending=False)
    return {
        "n_alertas": int(len(alertas)),
        "df_alertas": alertas,
        "df_pagos_sin_factura": df_sin.sort_values(
            "Monto_Pagado_CRC", ascending=False),
        "n_pagos_sin_factura": int(len(df_sin)),
        "resumen": resumen,
    }


agente_auditor.CONTRATO = {
    "entradas": "Salida del Conciliador",
    "salidas": "Tabla de alertas ordenada por monto + pagos R9",
    "reglas": ["R9"],
    "temperatura": TEMPERATURA,
    "falla_si": "nada: si no hay alertas devuelve tablas vacías",
}


def agente_detector_duplicados(df_conc, df_pag):
    """Agente 4 — Detector de duplicados: detalla cada factura con 2+ pagos (R6)."""
    dupls = df_conc[df_conc["Estado"] == "PAGO DUPLICADO"]
    detalle = []
    for _, f in dupls.iterrows():
        pagos = df_pag[df_pag[COL_LIMPIO] == f[COL_LIMPIO]].sort_values(
            "Fecha_Pago")
        detalle.append({
            "factura": str(f["Num_Factura"]),
            "proveedor": str(f["Proveedor"]),
            "facturado": float(f["Monto_CRC"]),
            "n_pagos": int(f["Num_Pagos"]),
            "total_pagado": float(f["Monto_Pagado"]),
            "pagado_de_mas": float(f["Diferencia"]),
            "pagos": [
                {"referencia": str(p["Referencia_Pago"]),
                 "fecha": str(p["Fecha_Pago"].date())
                 if hasattr(p["Fecha_Pago"], "date") else str(p["Fecha_Pago"]),
                 "monto": float(p["Monto_Pagado_CRC"]),
                 "cuenta": str(p["Cuenta_Bancaria"])}
                for _, p in pagos.iterrows()
            ],
        })
    detalle.sort(key=lambda d: d["pagado_de_mas"], reverse=True)
    return {
        "n_facturas_duplicadas": len(detalle),
        "total_pagado_de_mas": float(sum(d["pagado_de_mas"] for d in detalle)),
        "detalle": detalle,
    }


agente_detector_duplicados.CONTRATO = {
    "entradas": "Facturas conciliadas + pagos del Extractor",
    "salidas": "Por cada factura duplicada: pagos individuales + monto de más",
    "reglas": ["R6"],
    "temperatura": TEMPERATURA,
    "falla_si": "nada: si no hay duplicados devuelve detalle vacío",
}


def ejecutar_pipeline(fuente_facturas, hoja_fact, fuente_pagos, hoja_pagos,
                      corte, tolerancia=TOLERANCIA_DEFAULT):
    """Orquestación secuencial: Extractor -> Conciliador -> Auditor -> Detector.

    Devuelve dict con la salida de cada agente + bitácora del flujo.
    """
    bitacora = []
    ext = agente_extractor(fuente_facturas, hoja_fact, fuente_pagos, hoja_pagos)
    bitacora.append(
        f"Extractor: {ext['filas_facturas']} facturas + {ext['filas_pagos']} "
        f"pagos ({ext['limpieza_facturas'] + ext['limpieza_pagos']} con limpieza R1)")
    con = agente_conciliador(ext["df_facturas"], ext["df_pagos"], corte,
                             tolerancia)
    bitacora.append(
        f"Conciliador: {con['resumen']['PAGADA']['conteo']} PAGADA con "
        f"corte {corte} y tolerancia {tolerancia}")
    aud = agente_auditor(con["df_conciliado"], con["df_sin_factura"],
                         con["resumen"])
    bitacora.append(
        f"Auditor: {aud['n_alertas']} alertas + "
        f"{aud['n_pagos_sin_factura']} pagos sin factura")
    dup = agente_detector_duplicados(con["df_conciliado"], ext["df_pagos"])
    bitacora.append(
        f"DetectorDuplicados: {dup['n_facturas_duplicadas']} facturas "
        f"duplicadas")
    return {"extractor": ext, "conciliador": con, "auditor": aud,
            "duplicados": dup, "bitacora": bitacora}


__all__ = ["TEMPERATURA", "agente_extractor", "agente_conciliador",
           "agente_auditor", "agente_detector_duplicados",
           "ejecutar_pipeline", "hojas_del_excel", "DatosInvalidosError"]
