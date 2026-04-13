import streamlit as st
import pandas as pd
from rapidfuzz import fuzz, process
from io import BytesIO
import re
import unicodedata

st.set_page_config(page_title="Matching PRO", layout="wide")

st.title("🔍 Matching Inteligente PRO (Sin IA)")

# -------------------------
# 🧠 LIMPIEZA AVANZADA
# -------------------------

STOPWORDS = [
    "sa", "s.a", "ltda", "inc", "corp", "company", "co", "srl"
]

def limpiar_texto(texto):
    texto = str(texto).lower().strip()

    # quitar tildes
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

    # quitar caracteres especiales
    texto = re.sub(r'[^a-z0-9\s]', '', texto)

    # quitar stopwords
    palabras = texto.split()
    palabras = [p for p in palabras if p not in STOPWORDS]

    return " ".join(palabras)

# -------------------------
# 🔥 SCORING PRO
# -------------------------

def calcular_score(a, b):
    score1 = fuzz.token_set_ratio(a, b)
    score2 = fuzz.partial_ratio(a, b)
    score3 = fuzz.token_sort_ratio(a, b)

    # promedio base
    score = (score1 * 0.4 + score2 * 0.3 + score3 * 0.3)

    # bonus si uno contiene al otro
    if a in b or b in a:
        score += 10

    # penalización por diferencia de longitud
    len_diff = abs(len(a) - len(b))
    if len_diff > 10:
        score -= 5

    return min(100, max(0, score))

# -------------------------
# 🔄 MATCHING PRO
# -------------------------

def emparejar_pro(base1, base2, threshold):
    resultados = []
    usados = set()

    for item in base1:
        mejor_match = None
        mejor_score = 0

        for candidato in base2:
            if candidato in usados:
                continue

            score = calcular_score(item, candidato)

            if score > mejor_score:
                mejor_score = score
                mejor_match = candidato

        if mejor_score >= threshold:
            resultados.append((item, mejor_match, round(mejor_score, 2), "Coincidencia"))
            usados.add(mejor_match)
        else:
            resultados.append((item, None, round(mejor_score, 2), "Sin coincidencia"))

    return pd.DataFrame(
        resultados,
        columns=["Base 1", "Base 2", "Score", "Estado"]
    )

# -------------------------
# UI
# -------------------------

modo = st.radio(
    "Modo",
    ["📂 Entre archivos", "📄 Mismo archivo"]
)

threshold = st.slider("Umbral", 0, 100, 80)

# -------------------------
# 📂 MULTI ARCHIVO
# -------------------------

if modo == "📂 Entre archivos":

    archivos = st.file_uploader(
        "Sube archivos",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if archivos and len(archivos) >= 2:

        nombres = [f.name for f in archivos]
        maestro = st.selectbox("Archivo maestro", nombres)

        archivo_maestro = next(f for f in archivos if f.name == maestro)

        df_master = pd.read_excel(archivo_maestro)
        col_master = st.selectbox("Columna maestro", df_master.columns)

        base1 = df_master[col_master].dropna().apply(limpiar_texto)

        resultados = []

        for archivo in archivos:
            if archivo.name == maestro:
                continue

            df = pd.read_excel(archivo)
            col = df.columns[0]

            base2 = df[col].dropna().apply(limpiar_texto)

            df_res = emparejar_pro(base1, base2, threshold)
            df_res["Archivo"] = archivo.name

            resultados.append(df_res)

        df_final = pd.concat(resultados)

        st.dataframe(df_final)

        output = BytesIO()
        df_final.to_excel(output, index=False)

        st.download_button(
            "📥 Descargar",
            output.getvalue(),
            "matching_pro.xlsx"
        )

# -------------------------
# 📄 MISMO ARCHIVO
# -------------------------

if modo == "📄 Mismo archivo":

    archivo = st.file_uploader("Sube Excel", type=["xlsx"])

    if archivo:

        df = pd.read_excel(archivo)

        col1 = st.selectbox("Columna 1", df.columns)
        col2 = st.selectbox("Columna 2", df.columns)

        base1 = df[col1].dropna().apply(limpiar_texto)
        base2 = df[col2].dropna().apply(limpiar_texto)

        df_res = emparejar_pro(base1, base2, threshold)

        st.dataframe(df_res)

        output = BytesIO()
        df_res.to_excel(output, index=False)

        st.download_button(
            "📥 Descargar",
            output.getvalue(),
            "matching_pro.xlsx"
        )
