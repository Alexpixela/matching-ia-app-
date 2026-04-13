import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO
import re
import unicodedata

st.set_page_config(page_title="Matching ULTRA PRO", layout="wide")

st.title("🧠 Matching Inteligente + Calidad de Datos (FINAL PRO)")

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
# MATCHING GLOBAL (SOLUCIÓN REAL)
# -------------------------

def matching_ultra(base1, base2):
    resultados = []
    usados_a = set()
    usados_b = set()

    # 🔥 1. MATCH EXACTO
    for a in base1:
        if a in base2 and a not in usados_b:
            resultados.append({
                "Base 1": a,
                "Base 2": a,
                "Score": 100,
                "Estado": "✅ MATCH"
            })
            usados_a.add(a)
            usados_b.add(a)

    # 🔥 2. TODOS LOS POSIBLES MATCHES
    posibles = []

    for a in base1:
        if a in usados_a:
            continue

        for b in base2:
            if b in usados_b:
                continue

            score = calcular_score(a, b)
            posibles.append((a, b, score))

    # 🔥 3. ORDENAR POR SCORE
    posibles = sorted(posibles, key=lambda x: x[2], reverse=True)

    # 🔥 4. ASIGNAR SIN CONFLICTOS
    for a, b, score in posibles:
        if a in usados_a or b in usados_b:
            continue

        if score >= 85:
            estado = "✅ MATCH"
        elif score >= 65:
            estado = "⚠️ REVISAR"
        else:
            continue

        resultados.append({
            "Base 1": a,
            "Base 2": b,
            "Score": round(score, 2),
            "Estado": estado
        })

        usados_a.add(a)
        usados_b.add(b)

    # 🔥 5. NO MATCH BASE 1
    for a in base1:
        if a not in usados_a:
            resultados.append({
                "Base 1": a,
                "Base 2": None,
                "Score": 0,
                "Estado": "❌ NO MATCH"
            })

    # 🔥 6. SOBRANTES BASE 2
    for b in base2:
        if b not in usados_b:
            resultados.append({
                "Base 1": None,
                "Base 2": b,
                "Score": 0,
                "Estado": "❌ SOBRANTE B"
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
# MISMO ARCHIVO
# =========================

if modo == "📄 Mismo archivo":

    archivo = st.file_uploader("Sube Excel", type=["xlsx"])

    if archivo:
        excel = pd.ExcelFile(archivo)
        hoja = st.selectbox("Selecciona hoja", excel.sheet_names)

        df = pd.read_excel(excel, sheet_name=hoja)

        cols = df.select_dtypes(include="object").columns

        col1 = st.selectbox("Columna 1", cols)
        col2 = st.selectbox("Columna 2", cols)

        base1 = df[col1].dropna().apply(limpiar_texto)
        base2 = df[col2].dropna().apply(limpiar_texto)

        # 🔍 CALIDAD
        st.subheader("🔍 Calidad de datos")

        dup1, total1, pct1 = analizar_duplicados(base1, col1)
        dup2, total2, pct2 = analizar_duplicados(base2, col2)

        c1, c2 = st.columns(2)

        with c1:
            if total1 > 0:
                st.error(f"{col1}: {total1} duplicados ({pct1:.2f}%)")
                st.dataframe(dup1)
            else:
                st.success(f"{col1}: sin duplicados")

        with c2:
            if total2 > 0:
                st.error(f"{col2}: {total2} duplicados ({pct2:.2f}%)")
                st.dataframe(dup2)
            else:
                st.success(f"{col2}: sin duplicados")

        # ⚠️ SIMILARES
        st.subheader("⚠️ Posibles errores")
        similares = duplicados_similares(base1)

        if not similares.empty:
            st.dataframe(similares)
        else:
            st.success("Sin errores similares")

        # 🔗 MATCHING
        st.subheader("🔗 Matching")

        df_res = matching_ultra(base1, base2)

        filtro = st.multiselect(
            "Filtrar",
            ["✅ MATCH", "⚠️ REVISAR", "❌ NO MATCH", "❌ SOBRANTE B"],
            default=["✅ MATCH", "⚠️ REVISAR", "❌ NO MATCH", "❌ SOBRANTE B"]
        )

        df_res = df_res[df_res["Estado"].isin(filtro)]

        st.dataframe(df_res, use_container_width=True)

        # 📥 EXPORT
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, sheet_name="Matching", index=False)
            dup1.to_excel(writer, sheet_name="Duplicados Col1", index=False)
            dup2.to_excel(writer, sheet_name="Duplicados Col2", index=False)
            similares.to_excel(writer, sheet_name="Similares", index=False)

        st.download_button("📥 Descargar Excel", output.getvalue(), "reporte_final.xlsx")

# =========================
# MULTI ARCHIVO
# =========================

if modo == "📂 Multi archivo":

    archivos = st.file_uploader("Sube múltiples Excel", type=["xlsx"], accept_multiple_files=True)

    if archivos and len(archivos) >= 2:

        nombres = [f.name for f in archivos]
        maestro = st.selectbox("Archivo maestro", nombres)

        hojas = {}
        columnas = {}

        for archivo in archivos:
            excel = pd.ExcelFile(archivo)
            hoja = st.selectbox(f"Hoja - {archivo.name}", excel.sheet_names)

            df_temp = pd.read_excel(excel, sheet_name=hoja)
            cols = df_temp.select_dtypes(include="object").columns

            col = st.selectbox(f"Columna - {archivo.name}", cols)

            hojas[archivo.name] = hoja
            columnas[archivo.name] = col

        archivo_maestro = next(f for f in archivos if f.name == maestro)

        df_master = pd.read_excel(archivo_maestro, sheet_name=hojas[maestro])
        base1 = df_master[columnas[maestro]].dropna().apply(limpiar_texto)

        resultados = []

        for archivo in archivos:
            if archivo.name == maestro:
                continue

            df = pd.read_excel(archivo, sheet_name=hojas[archivo.name])
            base2 = df[columnas[archivo.name]].dropna().apply(limpiar_texto)

            df_res = matching_ultra(base1, base2)
            df_res["Archivo"] = archivo.name

            resultados.append(df_res)

        df_final = pd.concat(resultados)

        st.dataframe(df_final, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)

        st.download_button("📥 Descargar Excel", output.getvalue(), "matching_multi.xlsx")
