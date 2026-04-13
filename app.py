import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO
import re
import unicodedata

st.set_page_config(page_title="Matching ULTRA", layout="wide")

st.title("🧠 Matching Inteligente ULTRA (Claro y preciso)")

# -------------------------
# LIMPIEZA PRO
# -------------------------

STOPWORDS = ["sa", "s.a", "ltda", "inc", "corp", "company", "co", "srl"]

def limpiar_texto(texto):
    texto = str(texto).lower().strip()

    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

    texto = re.sub(r'[^a-z0-9\s]', '', texto)

    palabras = texto.split()
    palabras = [p for p in palabras if p not in STOPWORDS]

    return " ".join(palabras)

# -------------------------
# SCORE INTELIGENTE
# -------------------------

def calcular_score(a, b):
    score = (
        fuzz.token_set_ratio(a, b) * 0.5 +
        fuzz.partial_ratio(a, b) * 0.3 +
        fuzz.token_sort_ratio(a, b) * 0.2
    )

    if a in b or b in a:
        score += 10

    return min(100, score)

# -------------------------
# MATCHING LIMPIO
# -------------------------

def matching_ultra(base1, base2):
    resultados = []

    for item in base1:
        mejor = None
        mejor_score = 0

        for candidato in base2:
            score = calcular_score(item, candidato)

            if score > mejor_score:
                mejor_score = score
                mejor = candidato

        # CLASIFICACIÓN
        if mejor_score >= 85:
            estado = "✅ MATCH"
        elif mejor_score >= 65:
            estado = "⚠️ REVISAR"
        else:
            estado = "❌ NO MATCH"
            mejor = None  # importante

        resultados.append({
            "Base 1": item,
            "Match": mejor,
            "Score": round(mejor_score, 2),
            "Estado": estado
        })

    return pd.DataFrame(resultados)

# -------------------------
# UI
# -------------------------

modo = st.radio("Modo", ["📂 Multi-archivo", "📄 Mismo archivo"])

# =========================
# 📄 MISMO ARCHIVO
# =========================

if modo == "📄 Mismo archivo":

    archivo = st.file_uploader("Sube Excel", type=["xlsx"])

    if archivo:
        excel = pd.ExcelFile(archivo)
        hoja = st.selectbox("Hoja", excel.sheet_names)

        df = pd.read_excel(excel, sheet_name=hoja)

        cols = df.select_dtypes(include="object").columns

        col1 = st.selectbox("Columna 1", cols)
        col2 = st.selectbox("Columna 2", cols)

        base1 = df[col1].dropna().apply(limpiar_texto)
        base2 = df[col2].dropna().apply(limpiar_texto)

        df_res = matching_ultra(base1, base2)

        # FILTRO VISUAL 🔥
        filtro = st.multiselect(
            "Filtrar por estado",
            ["✅ MATCH", "⚠️ REVISAR", "❌ NO MATCH"],
            default=["✅ MATCH", "⚠️ REVISAR", "❌ NO MATCH"]
        )

        df_res = df_res[df_res["Estado"].isin(filtro)]

        st.dataframe(df_res)

        # EXPORT
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name="Resultados", index=False)

        st.download_button(
            "📥 Descargar",
            output.getvalue(),
            "matching_ultra.xlsx"
        )

# =========================
# 📂 MULTI ARCHIVO
# =========================

if modo == "📂 Multi-archivo":

    archivos = st.file_uploader(
        "Sube archivos",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if archivos and len(archivos) >= 2:

        nombres = [f.name for f in archivos]
        maestro = st.selectbox("Archivo maestro", nombres)

        hojas = {}
        columnas = {}

        for archivo in archivos:
            excel = pd.ExcelFile(archivo)
            hoja = st.selectbox(f"Hoja {archivo.name}", excel.sheet_names)

            df_temp = pd.read_excel(excel, sheet_name=hoja)
            cols = df_temp.select_dtypes(include="object").columns

            col = st.selectbox(f"Columna {archivo.name}", cols)

            hojas[archivo.name] = hoja
            columnas[archivo.name] = col

        archivo_maestro = next(f for f in archivos if f.name == maestro)
        df_master = pd.read_excel(
            archivo_maestro,
            sheet_name=hojas[maestro]
        )

        base1 = df_master[columnas[maestro]].dropna().apply(limpiar_texto)

        resultados = []

        for archivo in archivos:
            if archivo.name == maestro:
                continue

            df = pd.read_excel(
                archivo,
                sheet_name=hojas[archivo.name]
            )

            base2 = df[columnas[archivo.name]].dropna().apply(limpiar_texto)

            df_res = matching_ultra(base1, base2)
            df_res["Archivo"] = archivo.name

            resultados.append(df_res)

        df_final = pd.concat(resultados)

        st.dataframe(df_final)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)

        st.download_button(
            "📥 Descargar",
            output.getvalue(),
            "matching_ultra.xlsx"
        )
