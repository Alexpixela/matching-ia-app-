import streamlit as st

st.set_page_config(layout="wide")

# -------------------------
# HERO
# -------------------------
st.title("🧠 Limpia tu base de datos en segundos")

st.markdown("""
### Detecta duplicados, errores y nombres mal escritos automáticamente

✔️ Encuentra clientes duplicados  
✔️ Detecta errores de digitación  
✔️ Mide la calidad de tu base  
✔️ Funciona con cualquier Excel  

""")

col1, col2 = st.columns(2)

with col1:
    st.success("💡 Ahorra horas de trabajo manual")
    st.warning("⚠️ Evita errores en tu CRM")
    st.info("📊 Obtén un score de calidad de datos")

with col2:
    st.image("https://images.unsplash.com/photo-1551288049-bebda4e38f71")

# CTA PRINCIPAL
if st.button("🚀 Probar ahora"):
    st.switch_page("app.py")

# -------------------------
# PROBLEMA
# -------------------------
st.markdown("---")
st.header("😩 ¿Te pasa esto?")

st.markdown("""
- Tienes clientes duplicados  
- Nombres mal escritos  
- Bases desorganizadas  
- Reportes que no cuadran  
""")

st.error("Esto cuesta dinero y genera errores en tu negocio")

# -------------------------
# SOLUCIÓN
# -------------------------
st.markdown("---")
st.header("🚀 La solución")

st.markdown("""
Sube tu Excel y en segundos obtienes:

✔️ Matching inteligente  
✔️ Duplicados detectados  
✔️ Errores humanos  
✔️ Score de calidad de datos  
""")

# -------------------------
# DEMO
# -------------------------
st.markdown("---")
st.header("📊 Resultado real")

st.image("https://images.unsplash.com/photo-1551288049-bebda4e38f71")

st.success("👉 De caos a orden en segundos")

# -------------------------
# PRECIOS
# -------------------------
st.markdown("---")
st.header("💰 Planes simples")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Free")
    st.write("$0")
    st.write("1 archivo")
    st.button("Empezar gratis")

with col2:
    st.subheader("Pro")
    st.write("$10")
    st.write("50 archivos")
    st.button("Comprar Pro")

with col3:
    st.subheader("Empresa")
    st.write("$25")
    st.write("Ilimitado")
    st.button("Comprar Empresa")

# -------------------------
# CTA FINAL
# -------------------------
st.markdown("---")
st.header("🚀 Empieza ahora")

st.markdown("Limpia tu base de datos en minutos, no en horas")

if st.button("🔥 Ir a la app"):
    st.switch_page("app.py")
