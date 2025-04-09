import streamlit as st
import pandas as pd
from io import BytesIO
import importlib.util

st.set_page_config(page_title="Conciliación Bancaria", layout="centered")

st.title("💼 Conciliación Bancaria")

# ---- Cargar archivo conciliacionPY.py como módulo externo ----
spec = importlib.util.spec_from_file_location("conciliacionPY", "conciliacionGPT.py")
conciliacion = importlib.util.module_from_spec(spec)
spec.loader.exec_module(conciliacion)

# ---- Paso 1: Cargar archivo Excel ----
uploaded_file = st.file_uploader("📂 Cargar archivo Excel", type="xlsx")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    st.write("📄 Hojas disponibles:", sheets)

    # ---- Paso 2: Selección de hojas ----
    rb_sheet = st.selectbox("🧾 Seleccioná hoja para RB", sheets)
    lb_sheet = st.selectbox("🧾 Seleccioná hoja para LB", sheets)

    if rb_sheet and lb_sheet:
        df_RB = pd.read_excel(uploaded_file, sheet_name=rb_sheet)
        df_LB = pd.read_excel(uploaded_file, sheet_name=lb_sheet)

        # ---- Paso 3: Selección de columnas ----
        st.subheader("🧩 Selección de columnas")
        col_rb = df_RB.columns.tolist()
        col_lb = df_LB.columns.tolist()

        debe_rb = st.selectbox("💸 Columna DEBE - RB", col_rb)
        haber_rb = st.selectbox("💰 Columna HABER - RB", col_rb)
        debe_lb = st.selectbox("💸 Columna DEBE - LB", col_lb)
        haber_lb = st.selectbox("💰 Columna HABER - LB", col_lb)

        resultado = None
        nombre_archivo = None
        output = None

        # ---- Paso 4: Ejecutar conciliación ----
        if st.button("⚙️ Ejecutar conciliación"):
            try:
                dfpaso1, dfpaso2, dfpaso3, dfpaso4, df3 = conciliacion.run_conciliacion(
                    df_RB, df_LB, debe_rb, haber_rb, debe_lb, haber_lb
                )
                st.success("✅ Conciliación ejecutada correctamente")

                with st.expander("📊 Ver resultados"):
                    st.subheader("Resumen")
                    st.dataframe(df3)

                # Generamos archivo en memoria
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    dfpaso1.to_excel(writer, sheet_name='Paso 1', index=False)
                    dfpaso2.to_excel(writer, sheet_name='Paso 2', index=False)
                    dfpaso3.to_excel(writer, sheet_name='Paso 3', index=False)
                    dfpaso4.to_excel(writer, sheet_name='Paso 4', index=False)
                    df3.to_excel(writer, sheet_name='Resumen', index=False)
                output.seek(0)

                nombre_archivo = st.text_input("📄 Nombre del archivo Excel (sin .xlsx)", "ConciliacionPY_Julio")

                st.download_button(
                    label="⬇️ Descargar Excel",
                    data=output,
                    file_name=f"{nombre_archivo}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error al ejecutar conciliación: {e}")

        # ---- Paso 6: Botón de reinicio ----
        if st.button("🔄 Reiniciar todo"):
            st.rerun()
