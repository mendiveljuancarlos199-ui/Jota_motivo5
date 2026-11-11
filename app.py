
import io
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PANADERÍA FLORES - FLOW 360
# DASHBOARD DE BALANCEO DE CAPACIDAD
# HERRAMIENTAS: YAMAZUMI + TAKT TIME + PHVA
# ============================================================

st.set_page_config(
    page_title="Panadería Flores | Dashboard",
    page_icon="🥐",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# 1. DISEÑO VISUAL
# ============================================================

NARANJA = "#FFB454"
VERDE = "#34D399"
ROJO = "#FB7185"
AZUL = "#60A5FA"
MORADO = "#A78BFA"

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(125deg, #09111f, #17263e);
        color: #edf2ff;
    }

    [data-testid="stSidebar"] {
        background: #0d1728;
    }

    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #18283f, #263c56);
        border: 1px solid #354e6e;
        padding: 16px;
        border-radius: 17px;
    }

    .block-container {
        padding-top: 1.5rem;
    }

    h1, h2, h3 {
        color: #f8fafc !important;
    }

    .banner {
        background: linear-gradient(105deg, #233653, #64412f);
        padding: 25px;
        border-radius: 19px;
        border: 1px solid #795841;
        margin-bottom: 18px;
    }

    .banner h1 {
        margin: 0;
        font-size: 2.3rem;
    }

    .banner p {
        color: #e5e7eb;
        margin-bottom: 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 2. FUNCIONES
# ============================================================

def mostrar_grafico(fig, alto=400):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=alto,
        margin=dict(l=15, r=15, t=60, b=25),
        legend=dict(orientation="h", y=-0.25),
    )

    st.plotly_chart(fig, use_container_width=True)


def calcular_porcentaje(parte, total):
    if total <= 0:
        return 0.0

    return parte / total * 100


def crear_excel(hojas):
    archivo = io.BytesIO()

    with pd.ExcelWriter(
        archivo,
        engine="openpyxl",
    ) as writer:

        for nombre, tabla in hojas.items():
            tabla.to_excel(
                writer,
                sheet_name=nombre[:31],
                index=False,
            )

    return archivo.getvalue()


def validar_atenciones(tabla):
    tabla = tabla.copy()

    for columna in ["llegaron", "atendidos"]:
        tabla[columna] = pd.to_numeric(
            tabla[columna],
            errors="coerce",
        )

    valido = (
        not tabla[["llegaron", "atendidos"]]
        .isna()
        .any()
        .any()
        and (tabla["llegaron"] >= 0).all()
        and (tabla["atendidos"] >= 0).all()
        and (
            tabla["atendidos"] <= tabla["llegaron"]
        ).all()
    )

    return valido, tabla


# ============================================================
# 3. DATOS BASE
# ============================================================

FRANJAS = [
    "06:00-07:00",
    "07:00-08:00",
    "08:00-09:00",
    "09:00-10:00",
    "10:00-11:00",
    "11:00-12:00",
    "12:00-13:00",
    "13:00-14:00",
    "14:00-15:00",
    "15:00-16:00",
    "16:00-17:00",
    "17:00-18:00",
    "18:00-19:00",
    "19:00-20:00",
    "20:00-21:00",
]

DEMANDA_LV = [
    12, 22, 14, 6, 5,
    5, 6, 5, 4, 4,
    6, 13, 18, 12, 3,
]

DEMANDA_SD = [
    17, 30, 19, 8, 6,
    5, 6, 4, 4, 4,
    7, 17, 23, 14, 3,
]

DIAS = [
    "31/08", "01/09", "02/09", "03/09",
    "04/09", "05/09", "06/09", "07/09",
    "08/09", "09/09", "10/09", "11/09",
    "12/09", "13/09",
]

TIPOS = [
    "Laborable",
    "Laborable",
    "Laborable",
    "Laborable",
    "Laborable",
    "Fin de semana",
    "Fin de semana",
    "Laborable",
    "Laborable",
    "Laborable",
    "Laborable",
    "Laborable",
    "Fin de semana",
    "Fin de semana",
]

AS_LLEGARON = [
    88, 90, 86, 92, 96, 118, 112,
    89, 87, 91, 90, 98, 120, 110,
]

AS_ATENDIDOS = [
    80, 81, 78, 82, 85, 98, 94,
    81, 79, 82, 81, 86, 99, 93,
]

TO_LLEGARON = [
    90, 92, 88, 94, 98, 120, 114,
    91, 89, 93, 92, 100, 122, 112,
]

TO_ATENDIDOS = [
    86, 88, 84, 90, 93, 110, 105,
    87, 85, 89, 88, 95, 112, 103,
]

CAUSAS = [
    "Cola excesiva",
    "Concentración de demanda",
    "Cobro lento",
    "Producto agotado",
    "Demora percibida",
    "Otros",
]

FRECUENCIAS = [62, 41, 28, 18, 12, 7]

ELEMENTOS = [
    "E1 - Pedido",
    "E2 - Selección",
    "E3 - Embolsado",
    "E4 - Cálculo",
    "E5 - Cobro",
    "E6 - Registro",
    "E7 - Entrega",
]

TIEMPOS_AS = [
    11.5, 27.6, 8.1, 13.8,
    24.2, 34.5, 6.9,
]

TIEMPOS_TO = [
    11.5, 27.6, 8.1, 0,
    17.3, 9.2, 5.8,
]

PUESTOS_TO = [
    "Vendedor / Caja",
    "Apoyo / Despacho",
    "Apoyo / Despacho",
    "Eliminado",
    "Vendedor / Caja",
    "Vendedor / Caja",
    "Vendedor / Caja",
]

AUDITORIA_ITEMS = [
    "Activación del apoyo con 3+ clientes",
    "Apoyo ejecuta E2 y E3",
    "Vendedor ejecuta E1, E5, E6 y E7",
    "Cumplimiento del Takt",
    "Personal conoce la regla",
]

AUDITORIA_CUMPLE = [45, 47, 48, 46, 50]


# ============================================================
# 4. INICIALIZACIÓN
# ============================================================

if "as_is" not in st.session_state:

    st.session_state.as_is = pd.DataFrame({
        "dia": DIAS,
        "tipo": TIPOS,
        "llegaron": AS_LLEGARON,
        "atendidos": AS_ATENDIDOS,
    })


if "to_be" not in st.session_state:

    st.session_state.to_be = pd.DataFrame({
        "dia": DIAS,
        "tipo": TIPOS,
        "llegaron": TO_LLEGARON,
        "atendidos": TO_ATENDIDOS,
    })


if "auditoria" not in st.session_state:

    st.session_state.auditoria = pd.DataFrame({
        "criterio": AUDITORIA_ITEMS,
        "cumple": AUDITORIA_CUMPLE,
        "muestra": [50] * 5,
    })


# ============================================================
# 5. MENÚ LATERAL CON DASHBOARD
# ============================================================

st.sidebar.markdown("## 🥐 PANADERÍA FLORES")
st.sidebar.caption("FLOW 360 | Sistema de gestión")

st.sidebar.divider()

st.sidebar.markdown("### 📌 MENÚ PRINCIPAL")

MENU = [
    "🏠 Dashboard",
    "📈 Análisis de demanda",
    "🎯 Diagrama de Pareto",
    "⏱️ Takt Time",
    "🏗️ Gráfico Yamazumi",
    "🧪 Simulador de capacidad",
    "📋 PHVA y auditoría",
    "💰 Evaluación económica",
    "📥 Datos y reportes",
]

pagina = st.sidebar.radio(
    "Selecciona una sección",
    MENU,
    index=0,
)

st.sidebar.divider()

st.sidebar.markdown("### ⚙️ PARÁMETROS")

demanda_rafaga = st.sidebar.slider(
    "Clientes en ráfaga de 30 min",
    min_value=5,
    max_value=35,
    value=19,
)

cola = st.sidebar.slider(
    "Clientes esperando",
    min_value=0,
    max_value=15,
    value=4,
)

umbral_apoyo = st.sidebar.number_input(
    "Activar apoyo desde",
    min_value=1,
    max_value=10,
    value=3,
)

tiempo_disponible = st.sidebar.number_input(
    "Tiempo disponible (s)",
    min_value=300,
    max_value=7200,
    value=1800,
    step=300,
)

ciclo_as = st.sidebar.number_input(
    "Ciclo actual AS-IS (s)",
    min_value=1.0,
    max_value=600.0,
    value=138.0,
)

ticket = st.sidebar.number_input(
    "Ticket promedio S/",
    min_value=0.1,
    value=4.5,
    step=0.5,
)

meta = st.sidebar.slider(
    "Meta de atención (%)",
    80,
    100,
    90,
)

st.sidebar.divider()

st.sidebar.success(
    "Objetivo: alcanzar al menos "
    "90% de atención en horas pico."
)


# ============================================================
# 6. CÁLCULOS
# ============================================================

as_df = st.session_state.as_is.copy()
to_df = st.session_state.to_be.copy()

for tabla in [as_df, to_df]:

    tabla["no_atendidos"] = (
        tabla["llegaron"] - tabla["atendidos"]
    )

    tabla["atencion_pct"] = np.where(
        tabla["llegaron"] > 0,
        tabla["atendidos"]
        / tabla["llegaron"].replace(0, np.nan)
        * 100,
        0,
    )


as_total = int(as_df["llegaron"].sum())
as_atendidos = int(as_df["atendidos"].sum())
as_perdidos = as_total - as_atendidos

to_total = int(to_df["llegaron"].sum())
to_atendidos = int(to_df["atendidos"].sum())
to_perdidos = to_total - to_atendidos

as_pct = calcular_porcentaje(
    as_atendidos, as_total
)

to_pct = calcular_porcentaje(
    to_atendidos, to_total
)

takt = tiempo_disponible / demanda_rafaga

vendedor_to = sum(
    t
    for t, puesto in zip(
        TIEMPOS_TO, PUESTOS_TO
    )
    if puesto == "Vendedor / Caja"
)

apoyo_to = sum(
    t
    for t, puesto in zip(
        TIEMPOS_TO, PUESTOS_TO
    )
    if puesto == "Apoyo / Despacho"
)

ciclo_to = sum(TIEMPOS_TO)

ciclo_efectivo = max(
    vendedor_to,
    apoyo_to,
)

apoyo_activo = cola >= umbral_apoyo

ciclo_simulado = (
    ciclo_efectivo
    if apoyo_activo
    else ciclo_to
)

saturacion_as = ciclo_as / takt * 100

saturacion_to = (
    ciclo_simulado / takt * 100
)

capacidad_simulada = (
    tiempo_disponible / ciclo_simulado
)


# ============================================================
# 7. ENCABEZADO
# ============================================================

st.markdown(
    """
    <div class="banner">
        <h1>🥐 PANADERÍA FLORES</h1>
        <p>
            FLOW 360 | Dashboard de inteligencia
            operativa y balanceo de capacidad
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PÁGINA 1: DASHBOARD
# ============================================================

if pagina == "🏠 Dashboard":

    st.title("🏠 DASHBOARD GENERAL")

    st.write(
        "Centro de control de atención al cliente "
        "en horas pico de Panadería Flores."
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Atención AS-IS",
        f"{as_pct:.1f}%",
    )

    c2.metric(
        "Atención TO-BE",
        f"{to_pct:.1f}%",
        delta=f"{to_pct-as_pct:+.1f} pp",
    )

    c3.metric(
        "Clientes no atendidos AS-IS",
        as_perdidos,
    )

    reduccion = calcular_porcentaje(
        as_perdidos - to_perdidos,
        as_perdidos,
    )

    c4.metric(
        "Reducción proyectada",
        f"{reduccion:.1f}%",
    )

    st.progress(
        min(to_pct / meta, 1),
        text=(
            f"Atención proyectada: {to_pct:.1f}% "
            f"| Meta: {meta}%"
        ),
    )

    izquierda, derecha = st.columns(2)

    with izquierda:

        escenarios = pd.DataFrame({
            "Escenario": [
                "AS-IS",
                "TO-BE",
                "Meta",
            ],
            "Atención": [
                as_pct,
                to_pct,
                meta,
            ],
        })

        fig = px.bar(
            escenarios,
            x="Escenario",
            y="Atención",
            color="Escenario",
            text="Atención",
            title="Comparación de escenarios",
            color_discrete_sequence=[
                ROJO,
                VERDE,
                NARANJA,
            ],
        )

        fig.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="outside",
        )

        fig.update_yaxes(range=[0, 110])

        mostrar_grafico(fig)

    with derecha:

        fig = go.Figure(
            go.Indicator(
                mode="gauge+number",
                value=to_pct,
                number={"suffix": "%"},
                title={
                    "text": "Atención proyectada"
                },
                gauge={
                    "axis": {
                        "range": [0, 100]
                    },
                    "bar": {
                        "color": VERDE
                    },
                    "steps": [
                        {
                            "range": [0, 80],
                            "color": "#512C39",
                        },
                        {
                            "range": [80, meta],
                            "color": "#66532C",
                        },
                        {
                            "range": [meta, 100],
                            "color": "#215D4B",
                        },
                    ],
                    "threshold": {
                        "line": {
                            "color": "white",
                            "width": 4,
                        },
                        "value": meta,
                    },
                },
            )
        )

        mostrar_grafico(fig)

    st.subheader("📈 Evolución diaria")

    comparacion_diaria = pd.DataFrame({
        "Día": DIAS,
        "AS-IS": as_df["atencion_pct"],
        "TO-BE": to_df["atencion_pct"],
    })

    fig = px.line(
        comparacion_diaria,
        x="Día",
        y=["AS-IS", "TO-BE"],
        markers=True,
        title="Atención diaria por escenario",
        color_discrete_sequence=[
            ROJO,
            VERDE,
        ],
    )

    fig.add_hline(
        y=meta,
        line_dash="dash",
        line_color=NARANJA,
        annotation_text="Meta",
    )

    mostrar_grafico(fig)

    st.subheader("🚦 Estado de la operación")

    a, b, c = st.columns(3)

    a.metric(
        "Takt de la ráfaga",
        f"{takt:.1f} s",
    )

    b.metric(
        "Saturación AS-IS",
        f"{saturacion_as:.1f}%",
    )

    c.metric(
        "Saturación simulada",
        f"{saturacion_to:.1f}%",
    )

    if apoyo_activo:
        st.success(
            "🟢 Apoyo de despacho activado."
        )
    else:
        st.warning(
            "🟠 Operación con un solo puesto."
        )

    if to_pct >= meta:
        st.success(
            "La proyección TO-BE cumple "
            "la meta de atención."
        )
    else:
        st.warning(
            "La proyección TO-BE aún no "
            "alcanza la meta."
        )

    st.caption(
        "AS-IS corresponde al diagnóstico; "
        "TO-BE es un escenario proyectado."
    )


# ============================================================
# PÁGINA 2: DEMANDA
# ============================================================

elif pagina == "📈 Análisis de demanda":

    st.title("📈 ANÁLISIS DE DEMANDA")

    demanda = pd.DataFrame({
        "Franja": FRANJAS,
        "Laborable": DEMANDA_LV,
        "Fin de semana": DEMANDA_SD,
    })

    demanda["Promedio ponderado"] = (
        demanda["Laborable"] * 5
        + demanda["Fin de semana"] * 2
    ) / 7

    promedio = demanda[
        "Promedio ponderado"
    ].mean()

    umbral_pico = promedio * 1.2
    umbral_valle = promedio * 0.8

    demanda["Clasificación"] = np.select(
        [
            demanda["Promedio ponderado"]
            > umbral_pico,

            demanda["Promedio ponderado"]
            < umbral_valle,
        ],
        [
            "Hora pico",
            "Hora valle",
        ],
        default="Hora normal",
    )

    a, b, c = st.columns(3)

    a.metric(
        "Promedio",
        f"{promedio:.1f} clientes/h",
    )

    b.metric(
        "Umbral pico",
        f"{umbral_pico:.1f}",
    )

    c.metric(
        "Umbral valle",
        f"{umbral_valle:.1f}",
    )

    fig = px.bar(
        demanda,
        x="Franja",
        y=[
            "Laborable",
            "Fin de semana",
        ],
        barmode="group",
        title="Demanda por hora",
    )

    fig.add_hline(
        y=umbral_pico,
        line_dash="dash",
        line_color=ROJO,
    )

    mostrar_grafico(fig, 470)

    fig = px.bar(
        demanda,
        x="Franja",
        y="Promedio ponderado",
        color="Clasificación",
        title="Horas pico y valle",
        color_discrete_map={
            "Hora pico": ROJO,
            "Hora valle": VERDE,
            "Hora normal": AZUL,
        },
    )

    mostrar_grafico(fig)

    st.dataframe(
        demanda.round(2),
        hide_index=True,
        use_container_width=True,
    )

    st.info(
        "Las franjas prioritarias son "
        "06:00–09:00 y 17:00–20:00."
    )


# ============================================================
# PÁGINA 3: PARETO
# ============================================================

elif pagina == "🎯 Diagrama de Pareto":

    st.title("🎯 DIAGRAMA DE PARETO")

    pareto = pd.DataFrame({
        "Causa": CAUSAS,
        "Frecuencia": FRECUENCIAS,
    })

    pareto["Acumulado"] = (
        pareto["Frecuencia"].cumsum()
        / pareto["Frecuencia"].sum()
        * 100
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=pareto["Causa"],
            y=pareto["Frecuencia"],
            name="Frecuencia",
            marker_color=NARANJA,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=pareto["Causa"],
            y=pareto["Acumulado"],
            name="Acumulado (%)",
            mode="lines+markers",
            yaxis="y2",
            line=dict(
                color=VERDE,
                width=3,
            ),
        )
    )

    fig.update_layout(
        title="Pareto de causas de no atención",
        yaxis=dict(
            title="Frecuencia",
        ),
        yaxis2=dict(
            title="Acumulado (%)",
            overlaying="y",
            side="right",
            range=[0, 110],
        ),
    )

    mostrar_grafico(fig, 470)

    principales = sum(FRECUENCIAS[:3])

    st.metric(
        "Participación de C1, C2 y C3",
        f"{calcular_porcentaje(principales,sum(FRECUENCIAS)):.1f}%",
    )

    st.dataframe(
        pareto.round(2),
        hide_index=True,
        use_container_width=True,
    )

    st.warning(
        "Las causas prioritarias son "
        "cola excesiva, concentración "
        "de demanda y cobro lento."
    )


# ============================================================
# PÁGINA 4: TAKT TIME
# ============================================================

elif pagina == "⏱️ Takt Time":

    st.title("⏱️ LABORATORIO TAKT TIME")

    st.latex(
        r"Takt = \frac{Tiempo\ disponible}{Demanda}"
    )

    st.latex(
        r"Saturacion = \frac{Tiempo\ de\ ciclo}{Takt}"
        r"\times 100"
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "Takt",
        f"{takt:.1f} s",
    )

    b.metric(
        "Ciclo AS-IS",
        f"{ciclo_as:.1f} s",
    )

    c.metric(
        "Saturación AS-IS",
        f"{saturacion_as:.1f}%",
    )

    d.metric(
        "Saturación simulada",
        f"{saturacion_to:.1f}%",
    )

    escenarios = pd.DataFrame([
        ["Pico L-V 07-08", 3600, 22],
        ["Pico S-D 07-08", 3600, 30],
        ["Pico S-D 18-19", 3600, 23],
        ["Ráfaga L-V", 1800, 14],
        ["Ráfaga S-D", 1800, 19],
    ], columns=[
        "Escenario",
        "Tiempo",
        "Demanda",
    ])

    escenarios["Takt"] = (
        escenarios["Tiempo"]
        / escenarios["Demanda"]
    )

    escenarios["AS-IS (%)"] = (
        ciclo_as / escenarios["Takt"] * 100
    )

    escenarios["TO-BE (%)"] = (
        ciclo_efectivo
        / escenarios["Takt"]
        * 100
    )

    fig = px.bar(
        escenarios,
        x="Escenario",
        y=[
            "AS-IS (%)",
            "TO-BE (%)",
        ],
        barmode="group",
        title="Saturación por escenario",
        color_discrete_sequence=[
            ROJO,
            VERDE,
        ],
    )

    fig.add_hline(
        y=100,
        line_dash="dash",
        line_color=NARANJA,
        annotation_text="Límite 100%",
    )

    mostrar_grafico(fig, 470)

    st.dataframe(
        escenarios.round(2),
        hide_index=True,
        use_container_width=True,
    )


# ============================================================
# PÁGINA 5: YAMAZUMI
# ============================================================

elif pagina == "🏗️ Gráfico Yamazumi":

    st.title("🏗️ GRÁFICO YAMAZUMI")

    elementos = pd.DataFrame({
        "Elemento": ELEMENTOS,
        "AS-IS (s)": TIEMPOS_AS,
        "TO-BE (s)": TIEMPOS_TO,
        "Puesto TO-BE": PUESTOS_TO,
    })

    st.dataframe(
        elementos,
        hide_index=True,
        use_container_width=True,
    )

    diferencia = ciclo_as - sum(TIEMPOS_AS)

    if abs(diferencia) > 0.01:

        st.warning(
            f"Los tiempos AS-IS desglosados "
            f"suman {sum(TIEMPOS_AS):.1f} s, "
            f"mientras que el ciclo declarado "
            f"es {ciclo_as:.1f} s. "
            f"Diferencia: {diferencia:+.1f} s."
        )

    st.subheader("Yamazumi AS-IS")

    fig_as = go.Figure()

    for elemento, segundos in zip(
        ELEMENTOS,
        TIEMPOS_AS,
    ):

        fig_as.add_trace(
            go.Bar(
                x=["Vendedor único"],
                y=[segundos],
                name=elemento,
            )
        )

    if diferencia > 0:

        fig_as.add_trace(
            go.Bar(
                x=["Vendedor único"],
                y=[diferencia],
                name="Tiempo no desglosado",
                marker_color="#64748B",
            )
        )

    fig_as.update_layout(
        barmode="stack",
        title="Carga del vendedor AS-IS",
        yaxis_title="Segundos",
    )

    fig_as.add_hline(
        y=takt,
        line_dash="dash",
        line_color=ROJO,
        annotation_text=f"Takt {takt:.1f} s",
    )

    mostrar_grafico(fig_as, 470)

    st.subheader("Yamazumi TO-BE")

    fig_to = go.Figure()

    for elemento, segundos, puesto in zip(
        ELEMENTOS,
        TIEMPOS_TO,
        PUESTOS_TO,
    ):

        if segundos <= 0:
            continue

        fig_to.add_trace(
            go.Bar(
                x=[puesto],
                y=[segundos],
                name=elemento,
            )
        )

    fig_to.update_layout(
        barmode="stack",
        title="Balanceo entre dos puestos",
        yaxis_title="Segundos",
    )

    fig_to.add_hline(
        y=takt,
        line_dash="dash",
        line_color=ROJO,
        annotation_text=f"Takt {takt:.1f} s",
    )

    mostrar_grafico(fig_to, 470)

    a, b, c = st.columns(3)

    a.metric(
        "Vendedor",
        f"{vendedor_to:.1f} s",
    )

    b.metric(
        "Apoyo",
        f"{apoyo_to:.1f} s",
    )

    c.metric(
        "Ciclo efectivo teórico",
        f"{ciclo_efectivo:.1f} s",
    )

    st.info(
        f"Trabajo total propuesto: "
        f"{ciclo_to:.1f} s. "
        "El ciclo efectivo supone "
        "dos puestos trabajando en paralelo."
    )


# ============================================================
# PÁGINA 6: SIMULADOR
# ============================================================

elif pagina == "🧪 Simulador de capacidad":

    st.title("🧪 SIMULADOR DE CAPACIDAD")

    st.write(
        "Modifica los parámetros del menú lateral "
        "para observar cómo cambia la operación."
    )

    if apoyo_activo:

        st.success(
            "🟢 SEGUNDO PUESTO ACTIVADO"
        )

    else:

        st.warning(
            "🟠 OPERACIÓN CON UN SOLO PUESTO"
        )

    a, b, c = st.columns(3)

    a.metric(
        "Capacidad teórica",
        f"{capacidad_simulada:.1f} clientes",
    )

    b.metric(
        "Demanda",
        demanda_rafaga,
    )

    c.metric(
        "Diferencia",
        f"{capacidad_simulada-demanda_rafaga:+.1f}",
    )

    comparacion = pd.DataFrame({
        "Modelo": [
            "AS-IS",
            "Un puesto mejorado",
            "Dos puestos mejorados",
        ],
        "Capacidad": [
            tiempo_disponible / ciclo_as,
            tiempo_disponible / ciclo_to,
            tiempo_disponible / ciclo_efectivo,
        ],
    })

    fig = px.bar(
        comparacion,
        x="Modelo",
        y="Capacidad",
        color="Modelo",
        text="Capacidad",
        title="Comparación de capacidad",
    )

    fig.add_hline(
        y=demanda_rafaga,
        line_dash="dash",
        line_color=ROJO,
        annotation_text="Demanda",
    )

    mostrar_grafico(fig)

    st.subheader("Regla de activación")

    st.code(
        f"SI clientes_en_cola >= {umbral_apoyo}:\n"
        "    activar_apoyo_despacho()\n"
        "SINO:\n"
        "    operar_con_un_puesto()",
        language="python",
    )

    st.caption(
        "Modelo simplificado de capacidad. "
        "No predice directamente los tiempos "
        "reales de espera ni los abandonos."
    )


# ============================================================
# PÁGINA 7: PHVA
# ============================================================

elif pagina == "📋 PHVA y auditoría":

    st.title("📋 CICLO PHVA Y AUDITORÍA")

    phva = pd.DataFrame([
        [1, "Identificar horas pico", "Planear"],
        [2, "Medir AS-IS", "Planear"],
        [3, "Pareto de causas", "Planear"],
        [4, "Calcular Takt", "Planear"],
        [5, "Construir Yamazumi", "Hacer"],
        [6, "Redistribuir tareas", "Hacer"],
        [7, "Auditar TO-BE", "Verificar"],
        [8, "Evaluar economía", "Actuar"],
    ], columns=[
        "Paso",
        "Actividad",
        "Fase PHVA",
    ])

    st.dataframe(
        phva,
        hide_index=True,
        use_container_width=True,
    )

    st.subheader("Auditoría del balanceo")

    audit_edit = st.data_editor(
        st.session_state.auditoria,
        hide_index=True,
        use_container_width=True,
        key="editor_auditoria",
    )

    if st.button("Guardar auditoría"):

        nueva = audit_edit.copy()

        for columna in ["cumple", "muestra"]:
            nueva[columna] = pd.to_numeric(
                nueva[columna],
                errors="coerce",
            )

        valida = (
            not nueva[["cumple", "muestra"]]
            .isna().any().any()
            and (nueva["muestra"] > 0).all()
            and (nueva["cumple"] >= 0).all()
            and (
                nueva["cumple"] <= nueva["muestra"]
            ).all()
        )

        if valida:
            st.session_state.auditoria = nueva
            st.success("Auditoría guardada.")
            st.rerun()
        else:
            st.error("Revisa los valores.")

    audit = st.session_state.auditoria.copy()

    audit["Cumplimiento (%)"] = (
        audit["cumple"]
        / audit["muestra"]
        * 100
    )

    global_pct = calcular_porcentaje(
        audit["cumple"].sum(),
        audit["muestra"].sum(),
    )

    st.metric(
        "Cumplimiento global",
        f"{global_pct:.1f}%",
    )

    fig = px.bar(
        audit,
        x="criterio",
        y="Cumplimiento (%)",
        color="Cumplimiento (%)",
        title="Auditoría por criterio",
        color_continuous_scale="Tealgrn",
    )

    fig.update_yaxes(range=[0, 110])

    mostrar_grafico(fig)


# ============================================================
# PÁGINA 8: ECONOMÍA
# ============================================================

elif pagina == "💰 Evaluación económica":

    st.title("💰 EVALUACIÓN ECONÓMICA")

    a, b, c = st.columns(3)

    inversion = a.number_input(
        "Inversión inicial S/",
        min_value=0.0,
        value=500.0,
    )

    mantenimiento = b.number_input(
        "Mantenimiento anual S/",
        min_value=0.0,
        value=120.0,
    )

    dias_anuales = c.number_input(
        "Días operativos al año",
        min_value=1,
        max_value=366,
        value=360,
    )

    recuperados_dia = (
        as_perdidos / len(as_df)
        - to_perdidos / len(to_df)
    )

    beneficio = (
        recuperados_dia
        * ticket
        * dias_anuales
    )

    costo = inversion + mantenimiento

    bc = beneficio / costo if costo > 0 else 0

    roi = (
        (beneficio - costo) / costo * 100
        if costo > 0
        else 0
    )

    recuperacion = (
        inversion / (beneficio / 12)
        if beneficio > 0
        else float("inf")
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Ingreso recuperado anual",
        f"S/ {beneficio:,.2f}",
    )

    c2.metric(
        "B/C simplificado",
        f"{bc:.2f}",
    )

    c3.metric(
        "ROI simplificado",
        f"{roi:,.1f}%",
    )

    c4.metric(
        "Recuperación",
        (
            f"{recuperacion:.2f} meses"
            if np.isfinite(recuperacion)
            else "No recuperable"
        ),
    )

    economia = pd.DataFrame({
        "Concepto": [
            "Ingreso recuperado",
            "Inversión",
            "Mantenimiento",
        ],
        "Monto": [
            beneficio,
            inversion,
            mantenimiento,
        ],
    })

    fig = px.bar(
        economia,
        x="Concepto",
        y="Monto",
        color="Concepto",
        title="Análisis económico",
    )

    mostrar_grafico(fig)

    st.info(
        "El ingreso recuperado es una proyección "
        "de ventas brutas, no de utilidad neta. "
        "Para una evaluación financiera completa "
        "deben incluirse costos variables, margen "
        "y costo de oportunidad del personal."
    )


# ============================================================
# PÁGINA 9: REPORTES
# ============================================================

elif pagina == "📥 Datos y reportes":

    st.title("📥 DATOS Y REPORTES")

    st.subheader("Datos AS-IS")

    as_edit = st.data_editor(
        st.session_state.as_is,
        hide_index=True,
        use_container_width=True,
        key="editor_as",
    )

    if st.button("Guardar datos AS-IS"):

        valido, nuevo = validar_atenciones(
            as_edit
        )

        if valido:
            st.session_state.as_is = nuevo
            st.success("Datos guardados.")
            st.rerun()
        else:
            st.error(
                "Los atendidos no pueden superar "
                "los clientes que llegaron."
            )

    st.subheader("Datos TO-BE")

    to_edit = st.data_editor(
        st.session_state.to_be,
        hide_index=True,
        use_container_width=True,
        key="editor_to",
    )

    if st.button("Guardar datos TO-BE"):

        valido, nuevo = validar_atenciones(
            to_edit
        )

        if valido:
            st.session_state.to_be = nuevo
            st.success("Datos guardados.")
            st.rerun()
        else:
            st.error(
                "Los atendidos no pueden superar "
                "los clientes que llegaron."
            )

    st.subheader("Descargar reporte")

    indicadores = pd.DataFrame({
        "Indicador": [
            "Atención AS-IS (%)",
            "Atención TO-BE (%)",
            "Meta (%)",
            "Takt (s)",
            "Saturación AS-IS (%)",
            "Saturación simulada (%)",
        ],
        "Valor": [
            as_pct,
            to_pct,
            meta,
            takt,
            saturacion_as,
            saturacion_to,
        ],
    })

    demanda_excel = pd.DataFrame({
        "Franja": FRANJAS,
        "Laborable": DEMANDA_LV,
        "Fin de semana": DEMANDA_SD,
    })

    yamazumi_excel = pd.DataFrame({
        "Elemento": ELEMENTOS,
        "AS-IS": TIEMPOS_AS,
        "TO-BE": TIEMPOS_TO,
        "Puesto TO-BE": PUESTOS_TO,
    })

    pareto_excel = pd.DataFrame({
        "Causa": CAUSAS,
        "Frecuencia": FRECUENCIAS,
    })

    archivo_excel = crear_excel({
        "Indicadores": indicadores,
        "Demanda": demanda_excel,
        "AS_IS": as_df,
        "TO_BE": to_df,
        "Pareto": pareto_excel,
        "Yamazumi": yamazumi_excel,
        "Auditoria": st.session_state.auditoria,
    })

    st.download_button(
        "⬇️ Descargar Excel completo",
        data=archivo_excel,
        file_name="PANADERIA_FLORES_FLOW_360.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    st.subheader("Importar datos desde Excel")

    archivo_subido = st.file_uploader(
        "Selecciona un archivo Excel",
        type=["xlsx"],
    )

    if archivo_subido is not None:

        try:

            libro = pd.ExcelFile(archivo_subido)

            hoja = st.selectbox(
                "Hoja del Excel",
                libro.sheet_names,
            )

            importado = pd.read_excel(
                libro,
                sheet_name=hoja,
            )

            st.dataframe(
                importado.head(30),
                use_container_width=True,
            )

            if hoja in ["AS_IS", "TO_BE"]:

                if st.button("Importar hoja"):

                    columnas = [
                        "dia",
                        "tipo",
                        "llegaron",
                        "atendidos",
                    ]

                    if all(
                        col in importado.columns
                        for col in columnas
                    ):

                        valido, nuevo = validar_atenciones(
                            importado[columnas]
                        )

                        if valido:

                            destino = (
                                "as_is"
                                if hoja == "AS_IS"
                                else "to_be"
                            )

                            st.session_state[destino] = nuevo

                            st.success(
                                "Importación completada."
                            )

                            st.rerun()

                        else:
                            st.error(
                                "Los conteos no son válidos."
                            )

                    else:
                        st.error(
                            "Faltan columnas obligatorias."
                        )

        except Exception as error:

            st.error(
                f"Error al leer Excel: {error}"
            )


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.divider()

st.caption(
    "🥐 PANADERÍA FLORES | FLOW 360 | "
    "Dashboard académico de Ingeniería Industrial | "
    "Yamazumi + Takt Time + PHVA | "
    f"Consulta: {date.today().strftime('%d/%m/%Y')}"
)

# ===================== FIN =====================
