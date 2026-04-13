import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO
import re
import unicodedata

st.set_page_config(page_title="Matching ULTRA PRO", layout="wide")

st.title("🧠 Matching + Calidad de Datos (ULTRA PRO)")

# -------------------------
# LIMPIEZA
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
# SCORE
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
# MATCHING
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

        if mejor_score >= 85:
            estado = "✅ MATCH"
        elif mejor_score >= 65:
            estado = "⚠️ REVISAR"
        else:
            estado = "❌ NO MATCH"
            mejor = None

        resultados.append({
            "Base 1": item,
            "Match": mejor,
            "Score": round(mejor_score, 2),
            "Estado": estado
        })

    return pd.DataFrame(resultados)

# -------------------------
# DUPLICADOS
# -------------------------

def analizar_duplicados(serie, nombre):
    conteo = serie.value_counts()
    duplicados = conteo[conteo > 1]

    df_dup = duplicados.reset_index()
    df_dup.columns = [nombre, "Cantidad"]

    total = len(serie)
    total_dup = duplicados.sum()
    pct = (total_dup / total * 100) if total > 0 else 0

    return df_dup, total_dup, pct

# -------------------------
# SIMILARES
# -------------------------

def duplicados_similares(base, threshold=90):
    similares = []
    lista = base.tolist()

    for i in range(len(lista)):
        for j in range(i + 1, len(lista)):
            a, b = lista[i], lista[j]
            score = fuzz.ratio(a, b)

            if score >= threshold and a != b:
                similares.append((a, b, score))

    return pd.DataFrame(similares, columns=["Valor 1", "Valor 2", "Similitud"])

# -------------------------
# UI
# -------------------------

modo = st.radio("Modo", ["📄 Mismo archivo", "📂 Multi archivo"])

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

        # 🔍 CALIDAD DE DATOS
        st.subheader("🔍 Calidad de datos")

        dup1, total1, pct1 = analizar_duplicados(base1, col1)
        dup2, total2, pct2 = analizar_duplicados(base2, col2)

        colA, colB = st.columns(2)

        with colA:
            st.write(f"**{col1}**")
            if total1 > 0:
                st.error(f"🚨 {total1} duplicados ({pct1:.2f}%)")
                st.dataframe(dup1)
            else:
                st.success("✅ Sin duplicados")

        with colB:
            st.write(f"**{col2}**")
            if total2 > 0:
                st.error(f"🚨 {total2} duplicados ({pct2:.2f}%)")
                st.dataframe(dup2)
            else:
                st.success("✅ Sin duplicados")

        # ⚠️ SIMILARES
        st.subheader("⚠️ Posibles errores de digitación")
        similares = duplicados_similares(base1)

        if not similares.empty:
            st.warning("Posibles duplicados similares")
            st.dataframe(similares)
        else:
            st.success("Sin errores similares")

        # 🔗 MATCHING
        st.subheader("🔗 Matching")

        df_res = matching_ultra(base1, base2)

        filtro = st.multiselect(
            "Filtrar",
            ["✅ MATCH", "⚠️ REVISAR", "❌ NO MATCH"],
            default=["✅ MATCH", "⚠️ REVISAR", "❌ NO MATCH"]
        )

        df_res = df_res[df_res["Estado"].isin(filtro)]

        st.dataframe(df_res)

        # 📥 EXPORT
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name="Matching", index=False)
            dup1.to_excel(writer, sheet_name="Duplicados Col1", index=False)
            dup2.to_excel(writer, sheet_name="Duplicados Col2", index=False)
            similares.to_excel(writer, sheet_name="Similares", index=False)

        st.download_button(
            "📥 Descargar Excel",
            output.getvalue(),
            "reporte_ultra.xlsx"
        )

# =========================
# 📂 MULTI ARCHIVO
# =========================

if modo == "📂 Multi archivo":

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
            "matching_multi_ultra.xlsx"
        )
