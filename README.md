# 🔊 Estudio de Ruido de Fuente Fija — NOM-081-SEMARNAT-1994

Aplicación web para automatizar estudios de ruido de fuente fija conforme
a la **NOM-081-SEMARNAT-1994**, desplegada en [Streamlit Cloud](https://streamlit.io/cloud).

---

## 🗂 Estructura del repositorio

```
nom081-ruido/
├── app.py                   ← Aplicación Streamlit (punto de entrada)
├── estudio_ruido_nom.py     ← Módulo de cálculo (importado por app.py)
├── requirements.txt         ← Dependencias Python
├── .streamlit/
│   └── config.toml          ← Tema y configuración de Streamlit
└── README.md
```

---

## 🚀 Despliegue en Streamlit Cloud (paso a paso)

### 1. Crear repositorio en GitHub

```bash
# En tu máquina local
git init nom081-ruido
cd nom081-ruido

# Copia aquí todos los archivos de esta carpeta
git add .
git commit -m "feat: primera versión NOM-081 Streamlit app"

# Crea un repositorio en github.com y luego:
git remote add origin https://github.com/TU_USUARIO/nom081-ruido.git
git push -u origin main
```

### 2. Desplegar en Streamlit Cloud

1. Entra a **[share.streamlit.io](https://share.streamlit.io)** con tu cuenta de GitHub.
2. Haz clic en **"New app"**.
3. Selecciona tu repositorio `nom081-ruido` y la rama `main`.
4. En **"Main file path"** escribe: `app.py`
5. Haz clic en **"Deploy!"**

✅ En 2–3 minutos tendrás una URL pública del tipo:  
`https://TU_USUARIO-nom081-ruido-app-XXXXX.streamlit.app`

---

## 💻 Ejecución local

```bash
# 1. Clonar / descargar el repositorio
git clone https://github.com/TU_USUARIO/nom081-ruido.git
cd nom081-ruido

# 2. Crear entorno virtual (recomendado)
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar
streamlit run app.py
```

Abre tu navegador en `http://localhost:8501`

---

## 📋 Flujo de uso

| Paso | Acción |
|------|--------|
| 1 | Descarga las plantillas CSV desde el menú lateral |
| 2 | Llena las 35 lecturas de dB por periodo |
| 3 | Sube `fuente.csv` y `fondo.csv` en la pestaña **Datos de entrada** |
| 4 | Ingresa los metadatos del estudio (sitio, fecha, responsable…) |
| 5 | Presiona **Calcular Estudio** |
| 6 | Revisa los resultados y descarga el **Excel** con reporte completo |

---

## 🔢 Fórmulas implementadas

**Por periodo (35 lecturas):**
- `N50` = media aritmética
- `σ` = desviación estándar muestral (ddof=1)
- `N10 = N50 + 1.2817 × σ`
- `Neq = 10 × log₁₀[(1/35) × Σ10^(Nᵢ/10)]`

**Promedios globales:**
- `(Neq)eq = 10 × log₁₀[(1/5) × Σ10^(Neqᵢ/10)]`

**Correcciones:**
- `Ce = 0.9023 × σ_prom`
- `Δ50 = N50_fuente − N50_fondo`
- `N'50 = N50_prom + Ce`
- `Cf = −(Δ50+9) + 3√(4Δ50−3)` si `Δ50 ≥ 0.75`, si no **No Aplica**

**Resultado:**
- `Nff = max(N'50, (Neq)eq)`
- `(N')ff = Nff + Cf` (si aplica)

---

## ✅ Valores de validación (ZC1 Honda Hermosillo)

| Parámetro | Valor esperado |
|-----------|---------------|
| N50_prom fuente | 63.58 dB |
| σ_prom fuente | 2.12 dB |
| N10_prom fuente | 66.29 dB |
| (Neq)eq fuente | 64.12 dB |
| Ce | 1.91 dB |
| Δ50 | −0.84 dB |
| N'50 | 65.49 dB |
| Cf | No Aplica |
| **(N')ff** | **65.49 dB** |

---

## 📦 Dependencias

| Librería | Versión mínima | Uso |
|----------|---------------|-----|
| streamlit | 1.32.0 | Framework web |
| pandas | 1.5.0 | Manejo de datos |
| numpy | 1.23.0 | Cálculos numéricos |
| openpyxl | 3.0.10 | Exportación Excel |
| altair | 5.0.0 | Gráficas interactivas |

---

## 📄 Licencia

MIT — Libre para uso interno y comercial.
