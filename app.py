import streamlit as st
import pandas as pd
import numpy as np
from openai import OpenAI

# -------------------------
# CONFIG INICIAL
# -------------------------
st.set_page_config(page_title="Matching IA", layout="wide")

if "OPENAI_API_KEY" not in st.secrets:
    st.error("❌ Falta API KEY en Streamlit Cloud")
    st.stop()

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

import streamlit as st
import pandas as pd
import numpy as np
from openai import OpenAI
from io import BytesIO
import re

st.set_page_config(page_title="Matching IA", layout="wide")

st.title("🤖 Matching Inteligente con IA")

# -------------------------
# CONFIG API
# -------------------------
if "OPENAI_API_KEY" not in st.secrets:
    st.error("❌ Falta configurar OPENAI_API_KEY en Secrets")
    st.stop()

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# -------------------------
# FUNCIONES
# -------------------------

def limpiar_texto(texto):
    texto = texto.lower().strip()
    texto = re.sub(r'[^a-z0-9\s]', '', texto)
    return texto

@st.cache_data
def obtener_embeddings(textos):
    response = client.embeddings.create(
        model="text-embedding-3-small",
        input=textos
    )
    return [np.array(e.embedding) for e in response.data]

def cosine_similarity(a, b):
    return np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b))

def emparejar_embeddings(base1, emb1, base2, emb2, threshold):
    resultados = []
    usados = set()

    for i, vec1 in enumerate(emb1):
        mejor_score = -1
        mejor_j = None

        for j, vec2 in enumerate(emb2):
            if j in usados:
                continue

            score = cosine_similarity(vec1, vec2)

            if score > mejor_score:
                mejor_score = score
                mejor_j = j

        if mejor_score >= threshold:
            resultados.append((
                base1[i],
                base2[mejor_j],
                round(float(mejor_score), 4),
                "Coincidencia"
            ))
            usados.add(mejor_j)
        else:
            resultados.append((
                base1[i],
                None,
                round(float(mejor_score), 4),
                "Sin coincidencia"
            ))

    return pd.DataFrame(
        resultados,
        columns=["Base 1", "Base 2", "Similitud", "Estado"]
    )

# -------------------------
# MODO DE USO
# -------------------------

modo = st.radio(
    "Selecciona modo",
    ["📂 Entre archivos", "📄 Dentro del mismo archivo"]
)

threshold = st.slider("Umbral IA", 0.0, 1.0, 0.75)

# =========================
# 🔹 MODO 1: MULTI ARCHIVO
# =========================

if modo == "📂 Entre archivos":

    archivos = st.file_uploader(
        "Sube múltiples archivos Excel",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if archivos and len(archivos) >= 2:

        nombres = [f.name for f in archivos]
        archivo_maestro_nombre = st.selectbox("Archivo maestro", nombres)
        archivo_maestro = next(f for f in archivos if f.name == archivo_maestro_nombre)

        excel_master = pd.ExcelFile(archivo_maestro)
        hoja_master = st.selectbox("Hoja maestro", excel_master.sheet_names)
        df_master = pd.read_excel(excel_master, sheet_name=hoja_master)

        col_master = st.selectbox("Columna maestro", df_master.columns)

        base_master = df_master[col_master].dropna().astype(str).apply(limpiar_texto)

        st.info("⏳ Generando embeddings...")
        emb_master = obtener_embeddings(base_master.tolist())

        resultados_globales = []

        for archivo in archivos:
            if archivo.name == archivo_maestro_nombre:
                continue

            excel = pd.ExcelFile(archivo)
            hoja = excel.sheet_names[0]
            df = pd.read_excel(excel, sheet_name=hoja)

            col = df.columns[0]
            base = df[col].dropna().astype(str).apply(limpiar_texto)

            emb = obtener_embeddings(base.tolist())

            df_resultado = emparejar_embeddings(
                base_master.tolist(),
                emb_master,
                base.tolist(),
                emb,
                threshold
            )

            df_resultado["Archivo"] = archivo.name
            resultados_globales.append(df_resultado)

        df_final = pd.concat(resultados_globales, ignore_index=True)

        st.success("✅ Matching completado")
        st.dataframe(df_final)

# =========================
# 🔹 MODO 2: MISMO ARCHIVO
# =========================

elif modo == "📄 Dentro del mismo archivo":

    archivo = st.file_uploader("Sube un archivo Excel", type=["xlsx"])

    if archivo:

        excel = pd.ExcelFile(archivo)
        hoja = st.selectbox("Selecciona hoja", excel.sheet_names)

        df = pd.read_excel(excel, sheet_name=hoja)

        col1 = st.selectbox("Columna 1", df.columns)
        col2 = st.selectbox("Columna 2", df.columns)

        base1 = df[col1].dropna().astype(str).apply(limpiar_texto)
        base2 = df[col2].dropna().astype(str).apply(limpiar_texto)

        st.info("⏳ Generando embeddings...")

        emb1 = obtener_embeddings(base1.tolist())
        emb2 = obtener_embeddings(base2.tolist())

        df_resultado = emparejar_embeddings(
            base1.tolist(),
            emb1,
            base2.tolist(),
            emb2,
            threshold
        )

        st.success("✅ Matching completado")
        st.dataframe(df_resultado)

        # Descargar
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False)

        st.download_button(
            "📥 Descargar resultado",
            output.getvalue(),
            "matching_mismo_archivo.xlsx"
        )
