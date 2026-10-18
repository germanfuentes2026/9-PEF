"""
App Streamlit: Plan Económico Financiero (PEF) de concesión vial.

Ejecutar localmente:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import io
import os

import numpy as np
import pandas as pd
import streamlit as st

from pef_model import Inputs, calcular, cargar_inputs_excel, N

st.set_page_config(page_title="PEF Concesión Vial", page_icon="🛣️", layout="wide")

ARCHIVO_DEFECTO = os.path.join(os.path.dirname(__file__), "data", "PEF_X.xlsx")
MILLON = 1e6


# ---------------------------------------------------------------------------
# Utilidades de formato
# ---------------------------------------------------------------------------
def fmt_num(x, dec=0):
    if x is None or (isinstance(x, float) and not np.isfinite(x)):
        return "—"
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(x, dec=2):
    if x is None or isinstance(x, str):
        return "No calculable" if x is None else x
    return fmt_num(x * 100, dec) + " %"


def fmt_m(x):
    return "$ " + fmt_num(x / MILLON, 0) + " M"


def tabla_millones(df: pd.DataFrame) -> pd.DataFrame:
    out = df / MILLON
    out["TOTAL"] = out.sum(axis=1)
    return out


def estilo(df: pd.DataFrame, dec=0):
    return df.style.format(lambda v: fmt_num(v, dec))


# ---------------------------------------------------------------------------
# Carga de datos
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def leer_excel(contenido: bytes, nombre: str) -> dict:
    return cargar_inputs_excel(io.BytesIO(contenido), nombre=nombre).to_dict()


def buscar_archivo_defecto():
    """Busca el Excel por defecto en data/ o en la carpeta de la app (cualquier .xlsx)."""
    carpeta = os.path.dirname(os.path.abspath(__file__))
    candidatos = [ARCHIVO_DEFECTO, os.path.join(carpeta, "PEF_X.xlsx")]
    for sub in ("data", "Data", "DATA", ""):
        d = os.path.join(carpeta, sub)
        if os.path.isdir(d):
            candidatos += sorted(os.path.join(d, f) for f in os.listdir(d)
                                 if f.lower().endswith(".xlsx") and not f.startswith("~$"))
    for c in candidatos:
        if os.path.isfile(c):
            return c
    return None


def inputs_base():
    up = st.session_state.get("archivo_subido")
    if up is not None:
        return Inputs.from_dict(leer_excel(up.getvalue(), up.name))
    ruta = buscar_archivo_defecto()
    if ruta is None:
        return None
    with open(ruta, "rb") as fh:
        return Inputs.from_dict(leer_excel(fh.read(), os.path.basename(ruta)))


# ---------------------------------------------------------------------------
# Barra lateral
# ---------------------------------------------------------------------------
st.sidebar.title("🛣️ PEF Concesión Vial")
st.sidebar.file_uploader(
    "Cargar un PEF (.xlsx con hojas CONTROL / FLUJO / INPUTS)",
    type=["xlsx"], key="archivo_subido",
    help="Si no cargás ninguno, se usa PEF_X.xlsx incluido en la app.",
)
try:
    base = inputs_base()
except Exception as e:  # archivo con otra estructura
    st.error("No se pudo leer el archivo. Tiene que ser un PEF con las hojas CONTROL, FLUJO, INPUTS, "
             f"HIPOTESIS, IMPUESTOS y PRESTAMO. Detalle: {type(e).__name__}: {e}")
    st.stop()
if base is None:
    st.title("Plan Económico Financiero — Concesión Vial")
    st.info("No se encontró el archivo de datos por defecto (data/PEF_X.xlsx) en el repositorio. "
            "Cargá un PEF desde la barra lateral, o subí PEF_X.xlsx a GitHub dentro de una carpeta llamada "
            "`data` (o en la misma carpeta que app.py).")
    st.stop()
v = f"{base.nombre}"  # sufijo de claves: al cambiar de archivo se reinician los controles

if st.sidebar.button("↺ Restablecer valores del archivo"):
    for k in list(st.session_state.keys()):
        if k.endswith(v):
            del st.session_state[k]
    st.rerun()

st.sidebar.caption(f"Archivo: **{base.nombre}**")


def num(label, valor, key, **kw):
    return st.sidebar.number_input(label, value=valor, key=f"{key}|{v}", **kw)


def pct(label, valor, key, paso=0.1, **kw):
    """Input en % que devuelve la fracción."""
    r = st.sidebar.number_input(label + " (%)", value=round(valor * 100, 4), step=paso, key=f"{key}|{v}",
                                format="%.2f", **kw)
    return r / 100


inp = base.copia()

st.sidebar.subheader("Escenario")
inp.escenario = 2 if st.sidebar.radio("Plazo de concesión", ["20 años", "15 años"],
                                      index=0 if base.escenario == 2 else 1, horizontal=True,
                                      key=f"esc|{v}") == "20 años" else 1

st.sidebar.subheader("Tránsito y tarifa")
inp.uteq = num("UTEQUIS año base", float(base.uteq), "uteq", step=100_000.0, format="%.0f")
inp.tarifa = num("Tarifa ($ por UTEQUI)", float(base.tarifa), "tarifa", step=10.0, format="%.0f")
inp.crecimiento = pct("Crecimiento anual del tránsito", base.crecimiento, "crec")
inp.uteq_adicional = num("UTEQUIS de peajes adicionales", float(base.uteq_adicional), "adic", step=100_000.0, format="%.0f")
usar_d3 = st.sidebar.checkbox("Base de tránsito distinta desde el año 3",
                              value=base.uteq_base_desde_anio3 is not None, key=f"d3chk|{v}")
inp.uteq_base_desde_anio3 = (num("UTEQUIS base desde el año 3", float(base.uteq_base_desde_anio3 or base.uteq),
                                 "d3", step=100_000.0, format="%.0f") if usar_d3 else None)
inp.otros_ingresos_pct = pct("Otros ingresos (% de peajes)", base.otros_ingresos_pct, "otros")
with st.sidebar.expander("Sensibilidad nuevos peajes (años 2 a 5)"):
    inp.nuevos_peajes = [
        st.number_input(f"Año {i + 2}", value=float(base.nuevos_peajes[i]), step=0.05, key=f"np{i}|{v}")
        for i in range(4)
    ]
    st.caption("Igual que en el Excel, el factor del año 5 se arrastra a los años siguientes.")

st.sidebar.subheader("Sensibilidades de costos")
inp.sens_obras = pct("Obras", base.sens_obras, "s_obras", paso=1.0)
inp.sens_repav = pct("Repavimentación", base.sens_repav, "s_rep", paso=1.0)
inp.sens_cym = pct("Conservación y mantenimiento", base.sens_cym, "s_cym", paso=1.0)
inp.sens_atu = pct("Atención al usuario", base.sens_atu, "s_atu", paso=1.0)
inp.sens_adm = pct("Administración", base.sens_adm, "s_adm", paso=1.0)
inp.sens_expl = pct("Explotación", base.sens_expl, "s_expl", paso=1.0)

st.sidebar.subheader("Financiamiento")
inp.tasa_descuento = pct("Tasa de descuento del VAN", base.tasa_descuento, "td")
inp.pct_deuda = pct("Deuda sobre CAPEX total", base.pct_deuda, "deuda", paso=1.0)
inp.tna = pct("Tasa del préstamo (TNA)", base.tna, "tna")
inp.cuotas = int(num("Cantidad de cuotas", int(base.cuotas), "cuotas", min_value=1, max_value=20, step=1))
inp.sistema = 1 if st.sidebar.radio("Sistema de amortización del préstamo", ["Francés", "Alemán"],
                                    index=0 if base.sistema == 1 else 1, horizontal=True,
                                    key=f"sis|{v}") == "Francés" else 2
inp.gracia = st.sidebar.checkbox("Año de gracia (capitaliza el interés del año 1)", value=base.gracia, key=f"gr|{v}")

with st.sidebar.expander("Impuestos"):
    def p2(label, val, key):
        return st.number_input(label + " (%)", value=round(val * 100, 4), step=0.1, format="%.2f", key=f"{key}|{v}") / 100
    inp.iibb = p2("Ingresos Brutos", base.iibb, "iibb")
    inp.ganancias = p2("Impuesto a las Ganancias", base.ganancias, "gan")
    inp.idc = p2("Débitos y créditos bancarios", base.idc, "idc")
    inp.iva = p2("IVA", base.iva, "iva")
    inp.iva_capex = p2("IVA CAPEX", base.iva_capex, "ivacap")
    inp.iva_cym = p2("IVA contenido Conservación y Mant.", base.iva_cym, "ivacym")
    inp.iva_atu = p2("IVA contenido Atención al Usuario", base.iva_atu, "ivaatu")
    inp.iva_adm = p2("IVA contenido Administración", base.iva_adm, "ivaadm")
    inp.iva_expl = p2("IVA contenido Explotación", base.iva_expl, "ivaexpl")
    inp.sellos = p2("Sellos", base.sellos, "sellos")
    inp.municipales = p2("Tasas municipales", base.municipales, "mun")
    inp.pct_deduccion_intereses = p2("Tope deducción de intereses (art. 85)", base.pct_deduccion_intereses, "dedint")

with st.sidebar.expander("Amortización impositiva"):
    opciones = {1: "Obras en el plazo restante + repavimentación en su vida útil",
                2: "Obras y repavimentación en vida útil corta",
                3: "Instantánea (100% en el año de ejecución)"}
    inp.variante_amort = st.selectbox("Variante", list(opciones), format_func=lambda k: opciones[k],
                                      index=[1, 2, 3].index(base.variante_amort if base.variante_amort in (1, 2, 3) else 1),
                                      key=f"vam|{v}")
    inp.vida_util_repav = int(st.number_input("Vida útil repavimentación (años)", value=int(base.vida_util_repav),
                                              min_value=1, max_value=20, key=f"vur|{v}"))
    inp.vida_util_obras_5 = int(st.number_input("Vida útil obras en variante corta (años)", value=int(base.vida_util_obras_5),
                                                min_value=1, max_value=20, key=f"vuo|{v}"))
    inp.pct_amortizable = st.number_input("Porcentaje de la inversión que se amortiza (%)",
                                          value=round(base.pct_amortizable * 100, 2), step=5.0, key=f"pam|{v}") / 100

with st.sidebar.expander("Garantías"):
    inp.gar_pr_contrato = st.number_input("Garantía a 1er requerimiento — contrato ($)", value=float(base.gar_pr_contrato), step=1e9, format="%.0f", key=f"g1|{v}")
    inp.gar_pr_obras_1a5 = st.number_input("Garantía a 1er requerimiento — obras años 1 a 5 ($)", value=float(base.gar_pr_obras_1a5), step=1e9, format="%.0f", key=f"g2|{v}")
    inp.gar_pr_obras_6a20 = st.number_input("Garantía a 1er requerimiento — obras años 6 a 20 ($)", value=float(base.gar_pr_obras_6a20), step=1e9, format="%.0f", key=f"g3|{v}")
    inp.caucion_costo_obras = st.number_input("Costo póliza de caución obras (%)", value=base.caucion_costo_obras * 100, step=0.1, key=f"g4|{v}") / 100
    inp.caucion_costo_contrato = st.number_input("Costo póliza de caución contrato (%)", value=base.caucion_costo_contrato * 100, step=0.1, key=f"g5|{v}") / 100


# ---------------------------------------------------------------------------
# Edición de series anuales (CAPEX / OPEX)
# ---------------------------------------------------------------------------
cols_anio = [f"Año {i}" for i in range(1, N + 1)]
SERIES = [("obras", "Obras (puesta en valor + obligatorias)"), ("repav", "Repavimentación"),
          ("operacion", "Explotación / Operación"), ("mantenimiento", "Conservación y mantenimiento"),
          ("servicios", "Atención al usuario"), ("administracion", "Administración")]


def editor_series(tab):
    with tab:
        st.markdown("**Montos anuales con IVA, en millones de pesos.** Podés editar cualquier celda: el modelo se recalcula al instante.")
        df = pd.DataFrame([np.asarray(getattr(base, k)) / MILLON for k, _ in SERIES],
                          index=[lab for _, lab in SERIES], columns=cols_anio)
        ed = st.data_editor(df, key=f"ed_series|{v}", use_container_width=True,
                            column_config={c: st.column_config.NumberColumn(c, format="%.1f") for c in cols_anio})
        st.markdown("**Factores de ajuste** (1 = 100%). Por ejemplo, la rampa de mantenimiento de los primeros años.")
        fac = pd.DataFrame([getattr(base, f"{k}_f1") for k, _ in SERIES],
                           index=[lab for _, lab in SERIES], columns=cols_anio)
        ef = st.data_editor(fac, key=f"ed_fac|{v}", use_container_width=True,
                            column_config={c: st.column_config.NumberColumn(c, format="%.2f") for c in cols_anio})
        for (k, _), (_, fila) in zip(SERIES, ed.iterrows()):
            setattr(inp, k, list(np.asarray(fila.values, dtype=float) * MILLON))
        for (k, _), (_, fila) in zip(SERIES, ef.iterrows()):
            setattr(inp, f"{k}_f1", list(np.asarray(fila.values, dtype=float)))
        with st.expander("Factores avanzados (hoja HIPOTESIS)"):
            HIP = [("h_peajes", "Peajes existentes"), ("h_peajes_nuevos", "Peajes nuevos"), ("h_otros", "Otros ingresos"),
                   ("h_cym", "Conservación y mantenimiento"), ("h_atu", "Atención al usuario"), ("h_adm", "Administración"),
                   ("h_expl", "Explotación"), ("h_pv", "Obras de puesta en valor"), ("h_oo", "Obras obligatorias"),
                   ("h_rep", "Repavimentación")]
            hdf = pd.DataFrame([getattr(base, k) for k, _ in HIP], index=[l for _, l in HIP], columns=cols_anio)
            eh = st.data_editor(hdf, key=f"ed_hip|{v}", use_container_width=True,
                                column_config={c: st.column_config.NumberColumn(c, format="%.2f") for c in cols_anio})
            for (k, _), (_, fila) in zip(HIP, eh.iterrows()):
                setattr(inp, k, list(np.asarray(fila.values, dtype=float)))
            st.caption("El factor de peajes de un año se arrastra a los años siguientes (desde el año 6), igual que en el Excel.")


# ---------------------------------------------------------------------------
# Página principal
# ---------------------------------------------------------------------------
st.title("Plan Económico Financiero — Concesión Vial")
st.caption(f"Modelo: {base.nombre} · Plazo {inp.plazo} años · Año 1 = {inp.anio_inicio} · Valores en pesos con IVA")

tabs = st.tabs(["📊 Resumen", "💵 Flujo de fondos", "🏗️ CAPEX y OPEX", "🧾 Impuestos",
                "📉 Amortizaciones", "🏦 Préstamo", "🎯 Sensibilidad", "⬇️ Descargar"])
editor_series(tabs[2])          # define las series editadas antes de calcular

res = calcular(inp)
k = res.kpis

# ------------------------------- Resumen -----------------------------------
with tabs[0]:
    c = st.columns(4)
    c[0].metric("TIR del proyecto (sin préstamo)", fmt_pct(k["TIR del proyecto (sin préstamo)"]))
    c[1].metric("TIR con préstamo", fmt_pct(k["TIR (con IVA, incluye préstamo)"]))
    c[2].metric("TIR modificada", fmt_pct(k["TIR modificada"]))
    c[3].metric("VAN", fmt_m(k["VAN"]), help=f"Tasa de descuento {fmt_pct(inp.tasa_descuento)}")
    c = st.columns(4)
    c[0].metric("VAN del proyecto (sin préstamo)", fmt_m(k["VAN del proyecto (sin préstamo)"]))
    c[1].metric("VAI / VAE", fmt_num(k["VAI / VAE"], 3) if k["VAI / VAE"] is not None else "—")
    c[2].metric("Payback", f'{k["Payback (años)"]} años' if not isinstance(k["Payback (años)"], str) else k["Payback (años)"])
    c[3].metric("Payback descontado", f'{k["Payback descontado (años)"]} años' if not isinstance(k["Payback descontado (años)"], str) else k["Payback descontado (años)"])
    for a in res.avisos:
        st.warning(a)

    fl = res.flujo
    graf = pd.DataFrame({
        "Flujo neto": fl.loc["FLUJO NETO"].values / MILLON,
        "Flujo neto acumulado": fl.loc["FLUJO NETO ACUMULADO"].values / MILLON,
    }, index=[str(inp.anio_inicio + i) for i in range(N)])
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Flujo neto anual (millones de $)**")
        st.bar_chart(graf[["Flujo neto"]])
    with g2:
        st.markdown("**Flujo neto acumulado (millones de $)**")
        st.line_chart(graf[["Flujo neto acumulado"]])

    comp = pd.DataFrame({
        "Ingresos por peaje": [(fl.loc["Peajes"].sum() + fl.loc["Peajes adicionales"].sum() + fl.loc["Otros ingresos"].sum()) / MILLON],
        "OPEX operativo": [fl.loc[["Conservación y mantenimiento", "Atención al usuario", "Administración", "Explotación"]].values.sum() / MILLON],
        "CAPEX": [fl.loc["TOTAL CAPEX"].sum() / MILLON],
        "Impuestos": [fl.loc["Impuestos"].sum() / MILLON],
        "Servicio de la deuda": [fl.loc["Servicio de la deuda"].sum() / MILLON],
        "Garantías": [fl.loc["Garantías"].sum() / MILLON],
    }, index=["Total concesión"]).T
    st.markdown("**Totales de la concesión (millones de $)**")
    st.bar_chart(comp, horizontal=True)

# ------------------------------- Flujo -------------------------------------
with tabs[1]:
    st.markdown("**Flujo de fondos (millones de $, con IVA)**")
    st.dataframe(estilo(tabla_millones(res.flujo)), use_container_width=True, height=720)

with tabs[2]:
    st.markdown("**CAPEX y OPEX resultantes (millones de $, con IVA)**")
    st.dataframe(estilo(tabla_millones(res.capex_opex)), use_container_width=True)

with tabs[3]:
    st.markdown("**Impuestos (millones de $)**")
    st.dataframe(estilo(tabla_millones(res.impuestos)), use_container_width=True, height=720)

with tabs[4]:
    st.markdown("**Amortizaciones impositivas (millones de $, sin IVA)**")
    am = tabla_millones(res.amortizaciones)
    st.dataframe(estilo(am), use_container_width=True)
    total_base = am.loc[["Base obras (sin IVA)", "Base repavimentación (sin IVA)"], "TOTAL"].sum()
    total_am = am.loc["Amortización usada", "TOTAL"]
    if abs(total_base - total_am) < 0.01:
        st.success(f"Control OK: se amortiza el 100% de la inversión ({fmt_num(total_am, 1)} M).")
    else:
        st.error(f"Control: se amortizan {fmt_num(total_am, 1)} M de {fmt_num(total_base, 1)} M invertidos.")

with tabs[5]:
    st.markdown(f"**Préstamo:** crédito de {fmt_m(k['Crédito'])} "
                f"({fmt_pct(inp.pct_deuda, 1)} del CAPEX total), TNA {fmt_pct(inp.tna)}, "
                f"{inp.cuotas} cuotas, sistema {'francés' if inp.sistema == 1 else 'alemán'}"
                f"{', con año de gracia' if inp.gracia else ''}.")
    pr = res.prestamo.copy()
    for col in ["Cuota", "Interés", "Capital", "Saldo"]:
        pr[col] = pr[col] / MILLON
    st.dataframe(pr.style.format({c: (lambda x: fmt_num(x, 1)) for c in ["Cuota", "Interés", "Capital", "Saldo"]}),
                 use_container_width=True, hide_index=True)

# ------------------------------- Sensibilidad ------------------------------
with tabs[6]:
    st.markdown("**TIR del proyecto (sin préstamo) según tarifa y tránsito**")
    pasos = [-0.10, -0.05, 0.0, 0.05, 0.10]
    filas = {}
    for dt in pasos:
        fila = {}
        for du in pasos:
            r = calcular(inp.copia(tarifa=inp.tarifa * (1 + dt), uteq=inp.uteq * (1 + du), pct_deuda=0.0,
                                   uteq_base_desde_anio3=(inp.uteq_base_desde_anio3 * (1 + du)
                                                          if inp.uteq_base_desde_anio3 is not None else None)),
                         _sin_deuda_interno=True)
            fila[f"Tránsito {du:+.0%}"] = r.kpis["TIR (con IVA, incluye préstamo)"]
        filas[f"Tarifa {dt:+.0%} (${fmt_num(inp.tarifa * (1 + dt))})"] = fila
    st.dataframe(pd.DataFrame(filas).T.style.format(lambda x: fmt_pct(x)), use_container_width=True)

    st.markdown("**Impacto de ±10% en cada variable sobre la TIR del proyecto y el VAN**")
    variables = [("Tarifa", "tarifa"), ("Tránsito", "uteq"), ("Obras", "sens_obras"), ("Repavimentación", "sens_repav"),
                 ("Conservación y mantenimiento", "sens_cym"), ("Administración", "sens_adm"), ("Explotación", "sens_expl")]
    base_sd = calcular(inp.copia(pct_deuda=0.0), _sin_deuda_interno=True).kpis
    rows = []
    for nombre, campo in variables:
        out = {"Variable": nombre}
        for s, etiqueta in [(-0.10, "−10%"), (0.10, "+10%")]:
            if campo in ("tarifa", "uteq"):
                cambio = {campo: getattr(inp, campo) * (1 + s)}
                if campo == "uteq" and inp.uteq_base_desde_anio3 is not None:
                    cambio["uteq_base_desde_anio3"] = inp.uteq_base_desde_anio3 * (1 + s)
            else:
                cambio = {campo: (1 + getattr(inp, campo)) * (1 + s) - 1}
            r = calcular(inp.copia(pct_deuda=0.0, **cambio), _sin_deuda_interno=True).kpis
            out[f"TIR {etiqueta}"] = r["TIR (con IVA, incluye préstamo)"]
            out[f"VAN {etiqueta} (M$)"] = r["VAN"] / MILLON
        rows.append(out)
    sdf = pd.DataFrame(rows).set_index("Variable")
    st.caption(f"Base: TIR del proyecto {fmt_pct(base_sd['TIR (con IVA, incluye préstamo)'])} · VAN {fmt_m(base_sd['VAN'])}")
    st.dataframe(sdf.style.format({c: (fmt_pct if c.startswith("TIR") else (lambda x: fmt_num(x))) for c in sdf.columns}),
                 use_container_width=True)

# ------------------------------- Descargar ---------------------------------
with tabs[7]:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({"Indicador": list(k.keys()), "Valor": [x if not isinstance(x, (list, dict)) else str(x) for x in k.values()]}
                     ).to_excel(xw, sheet_name="Indicadores", index=False)
        tabla_millones(res.flujo).to_excel(xw, sheet_name="Flujo (M$)")
        tabla_millones(res.capex_opex).to_excel(xw, sheet_name="CAPEX OPEX (M$)")
        tabla_millones(res.impuestos).to_excel(xw, sheet_name="Impuestos (M$)")
        tabla_millones(res.amortizaciones).to_excel(xw, sheet_name="Amortizaciones (M$)")
        res.prestamo.to_excel(xw, sheet_name="Préstamo", index=False)
        pd.DataFrame([(a, str(b)) for a, b in inp.to_dict().items()], columns=["Parámetro", "Valor"]
                     ).to_excel(xw, sheet_name="Supuestos", index=False)
    st.download_button("Descargar resultados en Excel", buf.getvalue(),
                       file_name=f"Resultados_{os.path.splitext(base.nombre)[0]}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Incluye indicadores, flujo de fondos, CAPEX/OPEX, impuestos, amortizaciones, préstamo y los supuestos usados.")
