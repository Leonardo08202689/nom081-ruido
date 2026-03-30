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
    page_icon=":sound:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personalizado ─────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }

.page-header {
    border-left: 5px solid #1F4E79;
    padding: 10px 0 10px 18px;
    margin-bottom: 20px;
}
.page-header h1 { margin: 0; font-size: 1.45rem; font-weight: 700; color: #1F4E79; }
.page-header p  { margin: 4px 0 0; font-size: .85rem; color: #555; }

.kpi-card {
    background: #F7F9FC;
    border: 1px solid #D0DCEA;
    border-top: 3px solid #2E75B6;
    border-radius: 4px;
    padding: 14px 10px 10px;
    text-align: center;
}
.kpi-card .label { font-size: .73rem; color: #666; font-weight: 600;
                   text-transform: uppercase; letter-spacing: .05em; }
.kpi-card .value { font-size: 1.75rem; font-weight: 700; color: #1F4E79; line-height: 1.15; }
.kpi-card .unit  { font-size: .8rem; color: #999; }
.kpi-highlight { border-top-color: #1B7A40; }
.kpi-highlight .value { color: #1B7A40; }

.result-box {
    border: 2px solid;
    border-radius: 6px;
    padding: 18px 24px;
    margin: 16px 0;
}
.section-box {
    border-left: 3px solid #2E75B6;
    padding-left: 12px;
    margin: 8px 0 14px;
}
.section-box h3 { margin: 0; color: #1F4E79; font-size: .95rem; }

section[data-testid="stSidebar"] { background: #1F4E79 !important; }
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stButton button {
    background: #2E75B6; border: none; border-radius: 4px; width: 100%;
}
section[data-testid="stSidebar"] .stButton button:hover { background: #3A8FD1; }
</style>
""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════════════════════════════╗
# ║  HELPERS                                                        ║
# ╚══════════════════════════════════════════════════════════════════╝

@st.cache_data(show_spinner=False)
def _template_bytes(periodos: tuple, nombre: str) -> bytes:
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
        st.error(f" {e}")
        return None
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _excel_bytes(estudio: EstudioRuidoNOM) -> bytes:
    """Genera el Excel en memoria y devuelve bytes."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name
    try:
        estudio.exportar_excel(tmp_path)
        data = Path(tmp_path).read_bytes()
        return data
    except Exception as e:
        st.error(f"Error al generar Excel: {e}")
        return b""
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _word_bytes(res: dict) -> bytes:
    """
    Genera la Memoria de Calculo Word (NOM-081) usando generar_memoria_word.py.
    100 % Python — sin Node.js. Funciona en Streamlit Cloud.
    """
    from generar_memoria_word import generar_word
    return generar_word(res)


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
    st.markdown("## NOM-081")
    st.markdown("**Ruido de Fuente Fija**")
    st.markdown("---")

    st.markdown("### Plantillas CSV")
    st.markdown("Descarga, rellena con tus 35 lecturas y sube.")

    st.download_button(
        label="Descargar fuente_template.csv",
        data=_template_bytes(tuple(PERIODOS_FUENTE), "fuente"),
        file_name="fuente_template.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.download_button(
        label="Descargar fondo_template.csv",
        data=_template_bytes(tuple(PERIODOS_FONDO), "fondo"),
        file_name="fondo_template.csv",
        mime="text/csv",
        use_container_width=True,
    )

    st.markdown("---")
    st.markdown("### Fórmulas clave")
    st.markdown("""
`N50` — media aritmética  
`N10 = N50 + 1.2817 σ`  
`Ce  = 0.9023 σ_prom`  
`Nff = max(N'50, Neq_eq)`  
""")

    st.markdown("---")
    st.caption("SEMARNAT · NOM-081-SEMARNAT-1994")


# ╔══════════════════════════════════════════════════════════════════╗
# ║  BANNER                                                         ║
# ╚══════════════════════════════════════════════════════════════════╝

st.markdown("""
<div class="page-header">
  <h1>Estudio de Ruido de Fuente Fija &nbsp;&mdash;&nbsp; NOM-081-SEMARNAT-1994</h1>
  <p>Cargue los archivos CSV de mediciones, ingrese los metadatos del estudio y obtenga el reporte completo.</p>
</div>
""", unsafe_allow_html=True)


# ╔══════════════════════════════════════════════════════════════════╗
# ║  PESTAÑAS PRINCIPALES                                           ║
# ╚══════════════════════════════════════════════════════════════════╝

tab_datos, tab_resultados, tab_ayuda = st.tabs([
    "Datos de entrada",
    "Resultados",
    "Ayuda / Fórmulas",
])


# ══════════════════════════════════════════════════════════════════
# TAB 1 — Datos de entrada
# ══════════════════════════════════════════════════════════════════
with tab_datos:

    # ── Metadatos ─────────────────────────────────────────────────
    st.markdown('<div class="section-box"><h3>Información del Estudio</h3></div>',
                unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        meta_sitio   = st.text_input("Sitio / Zona",         value=st.session_state.get("meta_sitio", ""))
        st.session_state.meta_sitio = meta_sitio
        meta_fecha   = st.text_input("Fecha del estudio",    value=st.session_state.get("meta_fecha", ""), placeholder="dd/mm/aaaa")
        st.session_state.meta_fecha = meta_fecha
    with c2:
        meta_resp    = st.text_input("Responsable técnico",  value=st.session_state.get("meta_resp", ""))
        st.session_state.meta_resp = meta_resp
        meta_exp     = st.text_input("N° de expediente",     value=st.session_state.get("meta_exp", ""))
        st.session_state.meta_exp = meta_exp
    with c3:
        meta_limite  = st.number_input(
            "Límite permisible (dB)", min_value=40.0,
            max_value=120.0, value=st.session_state.get("meta_limite", 68.0), step=0.5,
            help="Límite de la norma aplicable al giro de la empresa"
        )
        st.session_state.meta_limite = meta_limite
        meta_giro    = st.text_input("Giro / Actividad",     value=st.session_state.get("meta_giro", ""))
        st.session_state.meta_giro = meta_giro

    st.markdown("---")

    # ── Carga de archivos ─────────────────────────────────────────
    st.markdown('<div class="section-box"><h3>Archivos CSV de Lecturas</h3></div>',
                unsafe_allow_html=True)

    col_f, col_b = st.columns(2)

    with col_f:
        st.markdown("**Ruido de Fuente** — Periodos A, B, C, D, E")
        up_fuente = st.file_uploader(
            "fuente.csv", type=["csv"],
            key="up_fuente", label_visibility="collapsed"
        )
        if up_fuente:
            st.success(f" {up_fuente.name}  ({up_fuente.size/1024:.1f} KB)")

    with col_b:
        st.markdown("**Ruido de Fondo** — Periodos I, II, III, IV, V")
        up_fondo = st.file_uploader(
            "fondo.csv", type=["csv"],
            key="up_fondo", label_visibility="collapsed"
        )
        if up_fondo:
            st.success(f" {up_fondo.name}  ({up_fondo.size/1024:.1f} KB)")

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
        " Calcular Estudio",
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

            st.session_state["res"]            = res
            st.session_state["estudio"]        = estudio
            st.session_state["limite"]         = meta_limite
            st.session_state["meta_sitio"]     = meta_sitio
            st.session_state["meta_fecha"]     = meta_fecha
            st.session_state["meta_resp"]      = meta_resp
            st.session_state["meta_ubicacion"] = meta_giro
            st.success(" Cálculo completado. Ve a la pestaña **Resultados**.")


# ══════════════════════════════════════════════════════════════════
# TAB 2 — Resultados
# ══════════════════════════════════════════════════════════════════
with tab_resultados:

    if "res" not in st.session_state:
        st.info("Sin resultados aún. "
                "Vaya a la pestaña Datos de entrada, cargue los CSV y presione Calcular.")
        st.stop()

    res     = st.session_state["res"]
    estudio = st.session_state["estudio"]
    limite  = st.session_state.get("limite", 68.0)



    # ── Resultado final ────────────────────────────────────────────
    excede = res["Nff_corr"] > limite
    color_text  = "#9B0000" if excede else "#1B5E20"
    color_border= "#E57373" if excede else "#66BB6A"
    color_bg    = "#FFF5F5" if excede else "#F5FBF6"
    estado = f"EXCEDE el límite ({limite:.1f} dB)" if excede \
             else f"CUMPLE el límite ({limite:.1f} dB)"

    st.markdown(f"""
    <div class="result-box" style="border-color:{color_border}; background:{color_bg};">
        <div style="font-size:.8rem; color:#555; font-weight:600; text-transform:uppercase;
                    letter-spacing:.06em; margin-bottom:4px;">
            Nivel de Fuente Fija — Resultado Final
        </div>
        <div style="font-size:2.8rem; font-weight:800; color:{color_text}; line-height:1.1;">
            {res['Nff_corr']:.2f} dB
        </div>
        <div style="font-size:1rem; font-weight:600; color:{color_text}; margin-top:6px;">
            {estado}
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Resumen de correcciones ────────────────────────────────────
    st.markdown("#### Detalle de Correcciones")

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

    # ── Exportar ──────────────────────────────────────────────────
    st.markdown("#### Exportar Resultados")

    base_name = (
        st.session_state.get("meta_sitio", "Estudio")
        .replace(" ", "_")
        .replace("/", "-")
        or "Estudio"
    )

    col_dl1, col_dl2 = st.columns(2)

    # Excel
    with col_dl1:
        xlsx_bytes = _excel_bytes(estudio)
        if xlsx_bytes:
            st.download_button(
                label="Descargar Excel",
                data=xlsx_bytes,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        else:
            st.error("No se pudo generar el archivo Excel. Revisa los logs.")
        st.caption(
            "Tres hojas: **Fuente**, **Fondo** y **Resumen** "
            "con tablas de datos, cálculos por periodo y resultado final."
        )

    # Word — Memoria de Cálculo
    with col_dl2:
        import copy
        # Ensamblar el dict completo que espera generar_memoria_word.py
        res_full = copy.deepcopy(res)
        res_full["metadata"] = {
            "compania":    st.session_state.get("meta_sitio", ""),
            "ubicacion":   st.session_state.get("meta_ubicacion", ""),
            "evaluadores": st.session_state.get("meta_resp", ""),
            "zona":        st.session_state.get("meta_sitio", ""),
            "evaluacion":  "Diurna",
            "hora_inicio": "",
            "hora_final":  "",
            "fecha":       st.session_state.get("meta_fecha", ""),
            "limite":      float(st.session_state.get("limite", 68.0)),
        }
        res_full["fuente_data"] = {
            p: estudio.df_fuente[p].tolist() for p in estudio.per_fuente
        }
        res_full["fondo_data"] = {
            p: estudio.df_fondo[p].tolist() for p in estudio.per_fondo
        }
        res_full["fuente_stats"] = {
            p: res["fuente"]["periodos"][p] for p in estudio.per_fuente
        }
        res_full["fondo_stats"] = {
            p: res["fondo"]["periodos"][p] for p in estudio.per_fondo
        }
        for p in estudio.per_fuente:
            res_full["fuente_stats"][p]["suma"] = round(float(estudio.df_fuente[p].sum()), 2)
        for p in estudio.per_fondo:
            res_full["fondo_stats"][p]["suma"]  = round(float(estudio.df_fondo[p].sum()),  2)
        res_full["promedios"] = {
            "fuente": {"N50": res["fuente"]["N50_prom"], "sigma": res["fuente"]["sigma_prom"],
                       "N10": res["fuente"]["N10_prom"], "Neq":   res["fuente"]["Neq_eq"]},
            "fondo":  {"N50": res["fondo"]["N50_prom"],  "sigma": res["fondo"]["sigma_prom"],
                       "N10": res["fondo"]["N10_prom"],  "Neq":   res["fondo"]["Neq_eq"]},
        }
        res_full["correcciones"] = {
            "Ce":        res["Ce"],
            "delta50":   res["delta50"],
            "N50_corr":  res["N50_corr"],
            "Cf_aplica": bool(res["Cf_aplica"]),
            "Cf":        res["Cf"] if res["Cf_aplica"] else "No Aplica",
        }
        res_full["resultado"] = {"Nff": res["Nff"], "Nff_corr": res["Nff_corr"]}

        try:
            docx_data = _word_bytes(res_full)
            st.download_button(
                label="Descargar Word — Memoria de Cálculo",
                data=docx_data,
                file_name=f"MemoriaCalculo_{base_name}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"
                ),
                use_container_width=True,
                type="primary",
            )
            st.caption(
                "Memoria de Cálculo estilo NOM-081: ficha de identificación, "
                "tablas de 35 lecturas, fórmulas explicadas paso a paso, "
                "correcciones y resultado final."
            )
        except Exception as e:
            st.error(f"Error al generar Word: {e}")


# ══════════════════════════════════════════════════════════════════
# TAB 3 — Ayuda
# ══════════════════════════════════════════════════════════════════
with tab_ayuda:
    col_h1, col_h2 = st.columns([1, 1])

    with col_h1:
        st.markdown("### Guía rápida de uso")
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
### Formato CSV
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
        st.markdown("### Fórmulas implementadas")

        st.markdown("**Por periodo (35 lecturas):**")
        st.latex(r"N_{50} = \frac{1}{n}\sum_{i=1}^{n} N_i")
        st.latex(r"\sigma = \sqrt{\frac{\sum(N_i - N_{50})^2}{n-1}}")
        st.latex(r"N_{10} = N_{50} + 1.2817\,\sigma")
        st.latex(r"N_{eq} = 10\log_{10}\!\left[\frac{1}{n}\sum_{i=1}^{n}10^{N_i/10}\right]")

        st.markdown("**Promedios globales (5 periodos):**")
        st.latex(r"\bar{N}_{50} = \frac{1}{5}\sum N_{50_i} \qquad \bar{\sigma} = \frac{1}{5}\sum \sigma_i")
        st.latex(r"(N_{eq})_{eq} = 10\log_{10}\!\left[\frac{1}{5}\sum_{i=1}^{5}10^{N_{eq_i}/10}\right]")

        st.markdown("**Correcciones:**")
        st.latex(r"C_e = 0.9023\,\bar{\sigma}")
        st.latex(r"N'_{50} = \bar{N}_{50} + C_e \quad (\text{si } \Delta_{50} \geq 0)")
        st.latex(r"\Delta_{50} = N_{50_{\text{fuente}}} - N_{50_{\text{fondo}}}")
        st.latex(r"C_f = -(\Delta_{50}+9) + 3\sqrt{4\Delta_{50}-3} \quad \text{si } \Delta_{50} \geq 0.75")
        st.latex(r"N_{ff} = \max\!\left(N'_{50},\;(N_{eq})_{eq}\right)")
        st.latex(r"(N')_{ff} = N_{ff} + C_f \quad \text{(si aplica)}")
