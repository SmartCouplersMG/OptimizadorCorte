import streamlit as st
import pandas as pd
from optimizador_refactorizado import ejecutar_optimizacion

# --- Configuración de la página ---
st.set_page_config(
    page_title="Optimizador de Corte",
    page_icon="✂️",
    layout="wide"
)

# --- Título y descripción ---
st.title("✂️ Optimizador de Corte de Materiales")
st.markdown("""
Esta aplicación web utiliza un script de Python para resolver el **Problema de Corte de Material (Cutting Stock Problem)**.
Sigue los siguientes pasos:
1.  **Descarga la plantilla** para ver el formato de entrada requerido.
2.  **Sube tu propio archivo** de Excel (`.xlsx`) con tus datos de inventario y despieces.
3.  **Ejecuta el optimizador** y visualiza los resultados.
""")

# --- Sidebar para carga y descarga ---
with st.sidebar:
    st.header("1. Descargar Plantilla")
    # Para crear el botón de descarga, primero debemos leer nuestro archivo de plantilla
    with open("dat_entrada.xlsx", "rb") as fp:
        st.download_button(
            label="Descargar dat_entrada.xlsx",
            data=fp,
            file_name="dat_entrada.xlsx",
            mime="application/vnd.ms-excel"
        )

    st.header("2. Cargar Archivo")
    uploaded_file = st.file_uploader(
        "Elige un archivo de Excel",
        type="xlsx"
    )

# --- Lógica principal ---
if uploaded_file is not None:
    st.header("🚀 Resultados de la Optimización")
    
    with st.spinner('Ejecutando el optimizador... Esto puede tardar unos segundos.'):
        # --- CAMBIO AQUÍ para recibir los 4 valores ---
        excel_res, pdf_res, resultados, logs = ejecutar_optimizacion(uploaded_file)

    if excel_res:
        st.success("¡Optimización completada con éxito!")

        with st.expander("Ver registro del proceso"):
            st.code('\n'.join(logs))

        # --- CAMBIO AQUÍ para ofrecer ambos botones de descarga ---
        st.subheader("3. Descargar Reportes")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="Descargar Reporte en Excel",
                data=excel_res,
                file_name="resultados_optimizacion.xlsx",
                mime="application/vnd.ms-excel"
            )
        with col2:
            if pdf_res:
                st.download_button(
                    label="Descargar Resumen en PDF",
                    data=pdf_res,
                    file_name="resumen_optimizacion.pdf",
                    mime="application/pdf"
                )

        st.subheader("📊 Visualización de los Planes de Corte")
        st.markdown("""
        El genera un plan de corte visual para cada diámetro. Aquí puedes verlos.
        Estos mismos gráficos están incluidos en el archivo de Excel que puedes descargar.
        """)
        
        # Iteramos sobre los resultados para mostrar los gráficos
        for diametro, data in resultados.items():
            if data.get('grafico'):
                st.markdown(f"---")
                st.markdown(f"#### Diámetro: {diametro}")
                col_img, col_info = st.columns([2,1])
                with col_img:
                    st.image(data['grafico'], caption=f"Plan de corte para diámetro {diametro}")
                with col_info:
                    # Accedemos al diccionario anidado 'metricas' de forma segura
                    metricas = data.get('metricas', {}) 
                    
                    # Usamos .get() con un valor por defecto (0) para evitar errores
                    eficiencia_total_pct = metricas.get('eficiencia_total', 0) * 100
                    desperdicio_corte = metricas.get('desperdicio_corte', 0)
                    
                    st.metric(label="Eficiencia Total", value=f"{eficiencia_total_pct:.2f} %")
                    st.metric(label="Desperdicio del Corte", value=f"{desperdicio_corte} cm")


    else:
        st.error("Ocurrió un error durante la optimización. Revisa los logs.")
        with st.expander("Ver registro del proceso"):
            st.code('\n'.join(logs))
else:
    st.info("Por favor, sube un archivo de Excel para comenzar el proceso de optimización.")

