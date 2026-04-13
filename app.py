import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO
import re
import unicodedata

st.set_page_config(page_title="Matching GOD MODE", layout="wide")

st.title("🔥 Matching Inteligente NIVEL DIOS")

# -------------------------
# 🧠 DICCIONARIO SEMÁNTICO
# -------------------------

SINONIMOS = {
    "clinica": "hospital",
    "hospital": "hospital",
    "banco": "banco",
    "financial": "banco",
    "corp": "",
    "company": "",
}

STOPWORDS = ["sa", "s.a", "ltda", "inc", "corp", "company", "co", "srl"]

# -------------------------
# 🧠 LIMPIEZA PRO+
# -------------------------

def limpiar_texto(texto):
    texto = str(texto).lower().strip()

    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

    texto = re.sub(r'[^a-z0-9\s]', '', texto)

    palabras = texto.split()

    palabras = [SINONIMOS.get(p, p) for p in palabras]
    palabras = [p for p in palabras if p not in STOPWORDS]

    return " ".join(palabras)

# -------------------------
# 🔥 SCORING NIVEL DIOS
# -------------------------

def calcular_score(a, b):
    score1 = fuzz.token_set_ratio(a, b)
    score2 = fuzz.partial_ratio(a, b)
    score3 = fuzz.token_sort_ratio(a, b)

    score = (score1 * 0.5 + score2 * 0.25 + score3 * 0.25)

    if a in b or b in a:
        score += 10

    len_diff = abs(len(a) - len(b))
    if len_diff > 15:
        score -= 5

    return min(100, max(0, score))

# -------------------------
# 🔥 TOP MATCHES
# -------------------------

def top_matches(base1, base2, top_n=3):
    resultados = []

    for item in base1:
        matches = []

        for candidato in base2:
            score = calcular_score(item, candidato)
            matches.append((candidato, score))

        matches = sorted(matches, key=lambda x: x[1], reverse=True)[:top_n]

        fila = {
            "Base": item
        }

        for i, (m, s) in enumerate(matches):
            fila[f"Match {i+1}"] = m
            fila[f"Score {i+1}"] = round(s, 2)

        resultados.append(fila)

    return pd.DataFrame(resultados)

# -------------------------
# 🎯 AUTO DETECCIÓN COLUMNAS
# -------------------------

def detectar_columnas(df):
    return df.select_dtypes(include="object").columns.tolist()

# -------------------------
# UI
# -------------------------

modo = st.radio("Modo", ["📂 Multi-archivo", "📄 Mismo archivo"])

threshold = st.slider("Umbral mínimo", 0, 100, 70)

# =========================
# 📂 MULTI ARCHIVO
# =========================

if modo == "📂 Multi-archivo":

    archivos = st.file_uploader(
        "Sube archivos Excel",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if archivos and len(archivos) >= 2:

        nombres = [f.name for f in archivos]
        maestro = st.selectbox("Archivo maestro", nombres)

        hojas = {}
        columnas = {}

        st.subheader("📑 Configuración")

        for archivo in archivos:
            excel = pd.ExcelFile(archivo)
            hoja = st.selectbox(f"Hoja {archivo.name}", excel.sheet_names)

            df_temp = pd.read_excel(excel, sheet_name=hoja)

            cols = detectar_columnas(df_temp)

            col = st.selectbox(f"Columna {archivo.name}", cols)

            hojas[archivo.name] = hoja
            columnas[archivo.name] = col

        # Procesar
        archivo_maestro = next(f for f in archivos if f.name == maestro)
        excel_master = pd.ExcelFile(archivo_maestro)

        df_master = pd.read_excel(
            excel_master,
            sheet_name=hojas[maestro]
        )

        base1 = df_master[columnas[maestro]].dropna().apply(limpiar_texto)

        resultados = []

        for archivo in archivos:
            if archivo.name == maestro:
                continue

            excel = pd.ExcelFile(archivo)
            df = pd.read_excel(excel, sheet_name=hojas[archivo.name])

            base2 = df[columnas[archivo.name]].dropna().apply(limpiar_texto)

            df_res = top_matches(base1, base2)

            # filtrar por threshold
            df_res = df_res[df_res["Score 1"] >= threshold]

            df_res["Archivo"] = archivo.name

            resultados.append(df_res)

        if resultados:
            df_final = pd.concat(resultados)

            st.success("✅ Matching completado")
            st.dataframe(df_final)

            # EXPORT PRO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name="Resultados", index=False)

            st.download_button(
                "📥 Descargar Excel PRO",
                output.getvalue(),
                "matching_god.xlsx"
            )

# =========================
# 📄 MISMO ARCHIVO
# =========================

if modo == "📄 Mismo archivo":

    archivo = st.file_uploader("Sube Excel", type=["xlsx"])

    if archivo:

        excel = pd.ExcelFile(archivo)
        hoja = st.selectbox("Hoja", excel.sheet_names)

        df = pd.read_excel(excel, sheet_name=hoja)

        cols = detectar_columnas(df)

        col1 = st.selectbox("Columna 1", cols)
        col2 = st.selectbox("Columna 2", cols)

        base1 = df[col1].dropna().apply(limpiar_texto)
        base2 = df[col2].dropna().apply(limpiar_texto)

        df_res = top_matches(base1, base2)

        df_res = df_res[df_res["Score 1"] >= threshold]

        st.success("✅ Matching completado")
        st.dataframe(df_res)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, index=False)

        st.download_button(
            "📥 Descargar",
            output.getvalue(),
            "matching_god.xlsx"
        )
