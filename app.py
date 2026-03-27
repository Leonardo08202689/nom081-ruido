"""
app.py — Estudio de Ruido de Fuente Fija NOM-081-SEMARNAT-1994
Desplegado en Streamlit Cloud
"""

import io
import math
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

from estudio_ruido_nom import (
    EstudioRuidoNOM,
    cargar_csv,
    generar_templates,
    PERIODOS_FUENTE,
    PERIODOS_FONDO,
    N_LECTURAS,
)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  CONFIGURACIÓN DE PÁGINA                                        ║
# ╚══════════════════════════════════════════════════════════════════╝

st.set_page_config(
    page_title="NOM-081 | Ruido de Fuente Fija",
    page_icon="🔊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personalizado ─────────────────────────────────────────────
st.markdown("""
<style>
/* Fuente global */
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }

/* Banner de título */
.banner {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    border-radius: 10px;
    padding: 20px 28px 16px 28px;
    margin-bottom: 24px;
    color: white;
}
.banner h1 { margin: 0; font-size: 1.7rem; font-weight: 700; }
.banner p  { margin: 4px 0 0 0; opacity: .85; font-size: .9rem; }

/* KPI cards */
.kpi-card {
    background: #F0F6FF;
    border: 1px solid #C3DAF5;
    border-radius: 10px;
    padding: 16px 10px 12px 10px;
    text-align: center;
}
.kpi-card .label { font-size: .78rem; color: #555; font-weight: 600;
                   text-transform: uppercase; letter-spacing: .04em; }
.kpi-card .value { font-size: 2rem; font-weight: 800;
                   color: #1F4E79; line-height: 1.1; }
.kpi-card .unit  { font-size: .85rem; color: #888; }

.kpi-highlight { background: #E8F5E9; border-color: #A5D6A7; }
.kpi-highlight .value { color: #1B5E20; }

/* Sección con borde lateral */
.section-box {
    border-left: 4px solid #2E75B6;
    padding-left: 14px;
    margin: 10px 0 18px 0;
}
.section-box h3 { margin: 0 0 4px 0; color: #1F4E79; font-size: 1rem; }

/* Tabla de resultados */
.result-row   { background: #F8FBFF; }
.result-label { font-weight: 600; color: #1F4E79; }

/* Badge norma */
.badge {
    display: inline-block;
    background: #E3F2FD;
    color: #1565C0;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: .78rem;
    font-weight: 600;
    margin-left: 8px;
}

/* Alerta verde resultado */
.result-final {
    background: linear-gradient(135deg,#E8F5E9,#C8E6C9);
    border: 1.5px solid #81C784;
    border-radius: 10px;
    padding: 16px 22px;
    margin-top: 12px;
}
.result-final .big { font-size: 2.4rem; font-weight: 900;
                     color: #1B5E20; }
.result-final .sub { color: #2E7D32; font-size: .9rem; margin-top: 2px; }

/* Sidebar */
section[data-testid="stSidebar"] { background: #1F4E79 !important; }
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stButton button {
    background: #2E75B6; border: none; color: white;
    border-radius: 8px; width: 100%;
}
section[data-testid="stSidebar"] .stButton button:hover {
    background: #4490D0;
}
</style>
""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════════════════════════════╗
# ║  HELPERS                                                        ║
# ╚══════════════════════════════════════════════════════════════════╝

@st.cache_data(show_spinner=False)
def _template_bytes(periodos: list, nombre: str) -> bytes:
    """Genera CSV plantilla en memoria."""
    df = pd.DataFrame({"lectura": range(1, N_LECTURAS + 1)})
    for p in periodos:
        df[p] = ""
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode()


def _cargar_upload(uploaded) -> dict | None:
    """Carga un UploadedFile de Streamlit como dict de datos."""
    if uploaded is None:
        return None
    with tempfile.NamedTemporaryFile(
        delete=False, suffix=".csv", mode="wb"
    ) as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = tmp.name
    try:
        return cargar_csv(tmp_path)
    except (FileNotFoundError, ValueError) as e:
        st.error(f"❌ {e}")
        return None
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _excel_bytes(estudio: EstudioRuidoNOM) -> bytes:
    """Genera el Excel en memoria y devuelve bytes."""
    with tempfile.NamedTemporaryFile(
        delete=False, suffix=".xlsx"
    ) as tmp:
        tmp_path = tmp.name
    estudio.exportar_excel(tmp_path)
    data = Path(tmp_path).read_bytes()
    Path(tmp_path).unlink(missing_ok=True)
    return data


def _tabla_periodo(bloque: dict, periodos: list) -> pd.DataFrame:
    """Convierte resultados de bloque a DataFrame para mostrar."""
    rows = []
    for indicador, key in [
        ("N50 (dB)",  "N50"),
        ("σ (dB)",    "sigma"),
        ("N10 (dB)",  "N10"),
        ("Neq (dB)",  "Neq"),
    ]:
        row = {"Indicador": indicador}
        for p in periodos:
            row[p] = round(bloque["periodos"][p][key], 2)
        row["Promedio"] = round({
            "N50":   bloque["N50_prom"],
            "sigma": bloque["sigma_prom"],
            "N10":   bloque["N10_prom"],
            "Neq":   bloque["Neq_eq"],
        }[key], 2)
        rows.append(row)
    return pd.DataFrame(rows).set_index("Indicador")


def _kpi(label: str, value: str, unit: str = "dB",
         highlight: bool = False):
    cls = "kpi-card kpi-highlight" if highlight else "kpi-card"
    st.markdown(f"""
    <div class="{cls}">
        <div class="label">{label}</div>
        <div class="value">{value}</div>
        <div class="unit">{unit}</div>
    </div>""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════════════════════════════╗
# ║  SIDEBAR                                                        ║
# ╚══════════════════════════════════════════════════════════════════╝

with st.sidebar:
    st.markdown("## 🔊 NOM-081")
    st.markdown("**Ruido de Fuente Fija**")
    st.markdown("---")

    st.markdown("### 📥 Plantillas CSV")
    st.markdown("Descarga, rellena con tus 35 lecturas y sube.")

    st.download_button(
        label="⬇ fuente_template.csv",
        data=_template_bytes(PERIODOS_FUENTE, "fuente"),
        file_name="fuente_template.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.download_button(
        label="⬇ fondo_template.csv",
        data=_template_bytes(PERIODOS_FONDO, "fondo"),
        file_name="fondo_template.csv",
        mime="text/csv",
        use_container_width=True,
    )

    st.markdown("---")
    st.markdown("### 📋 Fórmulas clave")
    st.markdown("""
- `N50` = media aritmética  
- `σ` = desv. estándar (ddof=1)  
- `N10 = N50 + 1.2817·σ`  
- `Neq = 10·log₁₀(⟨10^(Nᵢ/10)⟩)`  
- `Ce = 0.9023·σ_prom`  
- `N'50 = N50_prom + Ce`  
- `Nff = max(N'50, (Neq)eq)`  
""")

    st.markdown("---")
    st.caption("SEMARNAT · NOM-081-SEMARNAT-1994")


# ╔══════════════════════════════════════════════════════════════════╗
# ║  BANNER                                                         ║
# ╚══════════════════════════════════════════════════════════════════╝

st.markdown("""
<div class="banner">
  <h1>🔊 Estudio de Ruido de Fuente Fija
    <span class="badge">NOM-081-SEMARNAT-1994</span>
  </h1>
  <p>Carga los CSV de mediciones, ingresa los metadatos y obtén el reporte
     completo con exportación a Excel.</p>
</div>
""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════════════════════════════╗
# ║  PESTAÑAS PRINCIPALES                                           ║
# ╚══════════════════════════════════════════════════════════════════╝

tab_datos, tab_resultados, tab_ayuda = st.tabs([
    "📂  Datos de entrada",
    "📊  Resultados",
    "❓  Ayuda / Fórmulas",
])


# ══════════════════════════════════════════════════════════════════
# TAB 1 — Datos de entrada
# ══════════════════════════════════════════════════════════════════
with tab_datos:

    # ── Metadatos ─────────────────────────────────────────────────
    st.markdown('<div class="section-box"><h3>📝 Información del Estudio</h3></div>',
                unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        meta_sitio   = st.text_input("Sitio / Zona",         value="ZC1 Honda Hermosillo")
        meta_fecha   = st.text_input("Fecha del estudio",    value="", placeholder="dd/mm/aaaa")
    with c2:
        meta_resp    = st.text_input("Responsable técnico",  value="")
        meta_exp     = st.text_input("N° de expediente",     value="")
    with c3:
        meta_limite  = st.number_input(
            "Límite permisible (dB)", min_value=40.0,
            max_value=120.0, value=68.0, step=0.5,
            help="Límite de la norma aplicable al giro de la empresa"
        )
        meta_giro    = st.text_input("Giro / Actividad",     value="")

    st.markdown("---")

    # ── Carga de archivos ─────────────────────────────────────────
    st.markdown('<div class="section-box"><h3>📤 Archivos CSV de Lecturas</h3></div>',
                unsafe_allow_html=True)

    col_f, col_b = st.columns(2)

    with col_f:
        st.markdown("**Ruido de Fuente** — Periodos A, B, C, D, E")
        up_fuente = st.file_uploader(
            "fuente.csv", type=["csv"],
            key="up_fuente", label_visibility="collapsed"
        )
        if up_fuente:
            st.success(f"✅ {up_fuente.name}  ({up_fuente.size/1024:.1f} KB)")

    with col_b:
        st.markdown("**Ruido de Fondo** — Periodos I, II, III, IV, V")
        up_fondo = st.file_uploader(
            "fondo.csv", type=["csv"],
            key="up_fondo", label_visibility="collapsed"
        )
        if up_fondo:
            st.success(f"✅ {up_fondo.name}  ({up_fondo.size/1024:.1f} KB)")

    st.markdown("---")

    # ── Vista previa de datos ─────────────────────────────────────
    if up_fuente or up_fondo:
        st.markdown("**Vista previa de datos cargados**")
        prev_col1, prev_col2 = st.columns(2)

        if up_fuente:
            try:
                df_prev = pd.read_csv(up_fuente)
                up_fuente.seek(0)
                with prev_col1:
                    st.caption("Fuente (primeras 5 filas)")
                    st.dataframe(df_prev.head(), use_container_width=True,
                                 height=210)
            except Exception:
                pass

        if up_fondo:
            try:
                df_prev = pd.read_csv(up_fondo)
                up_fondo.seek(0)
                with prev_col2:
                    st.caption("Fondo (primeras 5 filas)")
                    st.dataframe(df_prev.head(), use_container_width=True,
                                 height=210)
            except Exception:
                pass

        st.markdown("")

    # ── Botón Calcular ────────────────────────────────────────────
    btn_calcular = st.button(
        "🧮  Calcular Estudio",
        type="primary",
        use_container_width=True,
        disabled=(up_fuente is None or up_fondo is None),
    )

    if up_fuente is None or up_fondo is None:
        st.info("⬆ Sube ambos archivos CSV para habilitar el cálculo. "
                "Usa las plantillas del menú lateral si no tienes el formato.")

    # ── Ejecutar cálculo ──────────────────────────────────────────
    if btn_calcular:
        with st.spinner("Procesando mediciones…"):
            up_fuente.seek(0)
            up_fondo.seek(0)
            datos_f = _cargar_upload(up_fuente)
            datos_b = _cargar_upload(up_fondo)

        if datos_f and datos_b:
            metadata = {
                "Sitio":        meta_sitio,
                "Fecha":        meta_fecha,
                "Responsable":  meta_resp,
                "Expediente":   meta_exp,
                "Giro":         meta_giro,
                "Límite (dB)":  str(meta_limite),
            }
            metadata = {k: v for k, v in metadata.items() if v}

            estudio = EstudioRuidoNOM(datos_f, datos_b, metadata)
            res     = estudio.calcular()

            st.session_state["res"]     = res
            st.session_state["estudio"] = estudio
            st.session_state["limite"]  = meta_limite
            st.success("✅ Cálculo completado. Ve a la pestaña **📊 Resultados**.")
            st.balloons()


# ══════════════════════════════════════════════════════════════════
# TAB 2 — Resultados
# ══════════════════════════════════════════════════════════════════
with tab_resultados:

    if "res" not in st.session_state:
        st.info("🔎 Aún no hay resultados. "
                "Ve a **📂 Datos de entrada**, sube los CSV y presiona Calcular.")
        st.stop()

    res     = st.session_state["res"]
    estudio = st.session_state["estudio"]
    limite  = st.session_state.get("limite", 68.0)

    # ── KPIs ──────────────────────────────────────────────────────
    st.markdown("### Indicadores globales")
    kc = st.columns(4)
    with kc[0]: _kpi("N50 fuente",    f"{res['fuente']['N50_prom']:.2f}")
    with kc[1]: _kpi("(Neq)eq fuente",f"{res['fuente']['Neq_eq']:.2f}")
    with kc[2]: _kpi("N50 fondo",     f"{res['fondo']['N50_prom']:.2f}")
    with kc[3]: _kpi("(Neq)eq fondo", f"{res['fondo']['Neq_eq']:.2f}")

    st.markdown("")

    kc2 = st.columns(4)
    with kc2[0]: _kpi("Ce",            f"{res['Ce']:.2f}")
    with kc2[1]: _kpi("Δ50",           f"{res['delta50']:.2f}")
    with kc2[2]: _kpi("N'50",          f"{res['N50_corr']:.2f}")
    with kc2[3]:
        cf_str = f"{res['Cf']:.2f}" if res["Cf_aplica"] else "N/A"
        _kpi("Cf", cf_str, unit="dB" if res["Cf_aplica"] else "—")

    # ── Resultado final ────────────────────────────────────────────
    excede = res["Nff_corr"] > limite
    color  = "#C62828" if excede else "#1B5E20"
    bg1    = "#FFEBEE" if excede else "#E8F5E9"
    bg2    = "#FFCDD2" if excede else "#C8E6C9"
    borde  = "#EF9A9A" if excede else "#81C784"
    icono  = "⚠️" if excede else "✅"
    estado = f"EXCEDE el límite ({limite:.1f} dB)" if excede \
             else f"CUMPLE el límite ({limite:.1f} dB)"

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{bg1},{bg2});
                border:2px solid {borde}; border-radius:12px;
                padding:18px 26px; margin:18px 0 10px 0;">
      <span style="font-size:.85rem;color:#555;font-weight:600;
                   text-transform:uppercase;letter-spacing:.05em;">
        Nivel de Fuente Fija — Resultado Final
      </span>
      <div style="font-size:3rem;font-weight:900;color:{color};
                  line-height:1.05;margin:4px 0;">
        (N')ff = {res['Nff_corr']:.2f} dB
      </div>
      <div style="color:{color};font-weight:600;font-size:1rem;">
        {icono} {estado}
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ── Tablas por bloque ─────────────────────────────────────────
    col_tbl1, col_tbl2 = st.columns(2)

    with col_tbl1:
        st.markdown("#### 🔉 Ruido de Fuente  (A – E)")
        df_f = _tabla_periodo(res["fuente"], res["fuente"]["periodos"].keys())
        st.dataframe(
            df_f.style
            .format("{:.2f}")
            .highlight_max(axis=1, color="#D6E4F0")
            .highlight_min(axis=1, color="#FFF9C4"),
            use_container_width=True,
        )

    with col_tbl2:
        st.markdown("#### 🔇 Ruido de Fondo  (I – V)")
        df_b = _tabla_periodo(res["fondo"], res["fondo"]["periodos"].keys())
        st.dataframe(
            df_b.style
            .format("{:.2f}")
            .highlight_max(axis=1, color="#D6E4F0")
            .highlight_min(axis=1, color="#FFF9C4"),
            use_container_width=True,
        )

    st.markdown("---")

    # ── Gráfica comparativa ───────────────────────────────────────
    st.markdown("#### 📈 Comparativo por Periodo")

    import altair as alt

    # Fuente
    per_f = list(res["fuente"]["periodos"].keys())
    rows_f = [{"Periodo": p, "N50": res["fuente"]["periodos"][p]["N50"],
               "N10": res["fuente"]["periodos"][p]["N10"],
               "Neq": res["fuente"]["periodos"][p]["Neq"],
               "Bloque": "Fuente"} for p in per_f]

    per_b = list(res["fondo"]["periodos"].keys())
    rows_b = [{"Periodo": p, "N50": res["fondo"]["periodos"][p]["N50"],
               "N10": res["fondo"]["periodos"][p]["N10"],
               "Neq": res["fondo"]["periodos"][p]["Neq"],
               "Bloque": "Fondo"} for p in per_b]

    # Combinar con etiqueta única
    for i, r in enumerate(rows_f):
        r["PeriodoLabel"] = f"F-{r['Periodo']}"
    for i, r in enumerate(rows_b):
        r["PeriodoLabel"] = f"B-{r['Periodo']}"

    df_chart = pd.DataFrame(rows_f + rows_b)

    # N50 + N10 + Neq lines
    df_melt = df_chart.melt(
        id_vars=["PeriodoLabel", "Bloque"],
        value_vars=["N50", "N10", "Neq"],
        var_name="Indicador", value_name="Valor"
    )

    chart = alt.Chart(df_melt).mark_line(point=True).encode(
        x=alt.X("PeriodoLabel:N", title="Periodo",
                 axis=alt.Axis(labelAngle=-30)),
        y=alt.Y("Valor:Q", title="Nivel (dB)",
                 scale=alt.Scale(zero=False)),
        color=alt.Color("Indicador:N",
                         scale=alt.Scale(
                             domain=["N50", "N10", "Neq"],
                             range=["#2E75B6", "#C00000", "#70AD47"])),
        strokeDash=alt.StrokeDash("Bloque:N",
                                   scale=alt.Scale(
                                       domain=["Fuente", "Fondo"],
                                       range=[[1, 0], [4, 2]])),
        tooltip=["PeriodoLabel", "Bloque", "Indicador",
                  alt.Tooltip("Valor:Q", format=".2f")],
    ).properties(height=320)

    # Línea de límite
    limite_line = alt.Chart(pd.DataFrame({"y": [limite]})).mark_rule(
        color="#FF0000", strokeDash=[6, 3], strokeWidth=1.5
    ).encode(y="y:Q")

    st.altair_chart(chart + limite_line, use_container_width=True)
    st.caption("Línea roja punteada = límite permisible. "
               "Línea continua = Fuente | Línea guionada = Fondo")

    st.markdown("---")

    # ── Resumen de correcciones ────────────────────────────────────
    st.markdown("#### 🔧 Detalle de Correcciones")

    corr_data = {
        "Parámetro": [
            "Ce — Correc. por valores extremos",
            "Δ50 — Diferencia fuente − fondo",
            "N'50 = N50_prom + Ce",
            "Cf — Correc. por fondo",
            "Nff = max(N'50, (Neq)eq)",
            "(N')ff — Resultado Final",
        ],
        "Fórmula": [
            "0.9023 × σ_prom",
            "N50_fuente − N50_fondo",
            f"{res['fuente']['N50_prom']:.2f} + {res['Ce']:.2f}",
            "−(Δ50+9)+3√(4Δ50−3)" if res["Cf_aplica"] else "No aplica (Δ50 < 0.75)",
            f"max({res['N50_corr']:.2f}, {res['fuente']['Neq_eq']:.2f})",
            f"{res['Nff']:.2f} + {res['Cf']:.2f}" if res["Cf_aplica"] else f"= Nff",
        ],
        "Valor (dB)": [
            f"{res['Ce']:.2f}",
            f"{res['delta50']:.2f}",
            f"{res['N50_corr']:.2f}",
            f"{res['Cf']:.2f}" if res["Cf_aplica"] else "—",
            f"{res['Nff']:.2f}",
            f"{res['Nff_corr']:.2f}",
        ],
    }
    df_corr = pd.DataFrame(corr_data)
    st.dataframe(df_corr, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Exportar Excel ────────────────────────────────────────────
    st.markdown("#### 📥 Exportar Resultados")
    col_dl1, col_dl2 = st.columns([1, 2])
    with col_dl1:
        xlsx_bytes = _excel_bytes(estudio)
        nombre_xlsx = f"NOM081_{meta_sitio.replace(' ','_')}.xlsx" \
                      if "meta_sitio" in dir() else "estudio_ruido_nom081.xlsx"
        st.download_button(
            label="⬇  Descargar Excel completo",
            data=xlsx_bytes,
            file_name=nombre_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    with col_dl2:
        st.caption(
            "El Excel incluye tres hojas: **Fuente** (35 lecturas + cálculos por periodo), "
            "**Fondo** (ídem) y **Resumen** (correcciones y resultado final con formato NOM-081)."
        )


# ══════════════════════════════════════════════════════════════════
# TAB 3 — Ayuda
# ══════════════════════════════════════════════════════════════════
with tab_ayuda:
    col_h1, col_h2 = st.columns([1, 1])

    with col_h1:
        st.markdown("### 📋 Guía rápida de uso")
        st.markdown("""
**Paso 1 — Obtener plantillas**  
Descarga `fuente_template.csv` y `fondo_template.csv` desde el menú lateral.

**Paso 2 — Llenar datos**  
Completa los 35 valores de dB medidos en campo para cada periodo.  
- Fuente: columnas `A`, `B`, `C`, `D`, `E`  
- Fondo:  columnas `I`, `II`, `III`, `IV`, `V`

**Paso 3 — Cargar archivos**  
Sube los CSV en la pestaña *Datos de entrada* y llena los metadatos del estudio.

**Paso 4 — Calcular**  
Presiona el botón **Calcular Estudio**.

**Paso 5 — Revisar y exportar**  
Revisa los resultados en la pestaña *Resultados* y descarga el Excel.

---
### 📁 Formato CSV
```
lectura,A,B,C,D,E          ← fuente.csv
1,62.5,63.1,64.0,61.8,65.2
2,63.0,62.8,63.5,62.4,64.1
...
35,63.3,64.2,64.8,63.0,65.5
```
```
lectura,I,II,III,IV,V      ← fondo.csv
1,63.8,64.2,65.0,64.5,64.7
...
```
""")

    with col_h2:
        st.markdown("### 🔢 Fórmulas implementadas")
        st.markdown("""
**Por cada periodo (35 lecturas):**

| Indicador | Fórmula |
|-----------|---------|
| N50  | media aritmética |
| σ    | desv. estándar muestral (ddof=1) |
| N10  | N50 + 1.2817 × σ |
| Neq  | 10·log₁₀[(1/35)·Σ10^(Nᵢ/10)] |

**Promedios globales (5 periodos):**

| Indicador | Fórmula |
|-----------|---------|
| N50_prom  | media de los 5 N50 |
| σ_prom    | media de las 5 σ |
| N10_prom  | N50_prom + 1.2817·σ_prom |
| (Neq)eq   | 10·log₁₀[(1/5)·Σ10^(Neqᵢ/10)] |

**Correcciones:**

| Símbolo | Fórmula |
|---------|---------|
| Ce | 0.9023 × σ_prom |
| Δ50 | N50_fuente − N50_fondo |
| N'50 | N50_prom + Ce |
| Cf | −(Δ50+9)+3√(4Δ50−3) si Δ50≥0.75 |
| Nff | max(N'50, (Neq)eq) |
| (N')ff | Nff + Cf (si aplica) |
""")

        st.markdown("---")
        st.markdown("### ✅ Valores de validación (ZC1 Honda Hermosillo)")
        val_data = {
            "Parámetro":        ["N50_prom", "σ_prom", "N10_prom","(Neq)eq",
                                  "Ce", "Δ50", "N'50", "Cf", "(N')ff"],
            "Fuente esperado":  ["63.58","2.12","66.29","64.12",
                                  "1.91","−0.84","65.49","N/A","65.49"],
            "Fondo esperado":   ["64.42","1.49","66.34","64.67",
                                  "—","—","—","—","—"],
        }
        st.dataframe(pd.DataFrame(val_data), use_container_width=True,
                     hide_index=True)
