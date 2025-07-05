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
Visita nuestra página web seria de gran ayuda: [www.scmgsas.com](https://www.scmgsas.com).
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
                mime="application/vnd.ms-excel",
                key="excel_download"  # <-- Se añade una clave única
            )

        with col2:
            if pdf_res:
                st.download_button(
                    label="Descargar Resumen en PDF",
                    data=pdf_res,
                    file_name="resumen_optimizacion.pdf",
                    mime="application/pdf",
                    key="pdf_download"  # <-- Se añade otra clave única
                )
        if st.session_state.get('excel_download') or st.session_state.get('pdf_download'):
            st.success("¡Descarga iniciada! Visita nuestra página web seria de gran ayuda: [www.scmgsas.com](https://www.scmgsas.com)")

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
                    if data.get('grafico_union'):
                        st.image(data['grafico_union'], caption=f"Plan de Unión para Diámetro {diametro}")
                with col_info:
                    # Accedemos al diccionario anidado 'metricas' de forma segura
                    metricas = data.get('metricas', {}) 
                    st.markdown("##### Resumen de Optimización") # Título para la sección

                    # --- Métricas de Corte ---
                    eficiencia_corte_pct = metricas.get('eficiencia_corte', 0) * 100
                    desperdicio_corte = metricas.get('desperdicio_corte', 0)
                    st.metric(label="Eficiencia de Corte", value=f"{eficiencia_corte_pct:.2f} %")
                    st.metric(label="Desperdicio por Corte", value=f"{desperdicio_corte} cm")
                    # --- Métricas de Unión (se muestran solo si existen) ---
                    if metricas.get('eficiencia_union') is not None:
                        eficiencia_union_pct = metricas.get('eficiencia_union', 0) * 100
                        exceso_union = metricas.get('exceso_union', 0)
                        st.metric(label="Eficiencia de Unión", value=f"{eficiencia_union_pct:.2f} %")
                        st.metric(label="Exceso por Unión", value=f"{exceso_union} cm")
                    
                    st.divider() # Un separador visual

                    # --- Métrica Total ---
                    eficiencia_total_pct = metricas.get('eficiencia_total', 0) * 100
                    st.metric(label="Eficiencia Total del Material", value=f"{eficiencia_total_pct:.2f} %")
    
        st.divider() # Un separador visual
        st.subheader("📦 Inventario Final Consolidado")

        inventario_final = resultados.get('inventario_final_consolidado')

        if inventario_final:
            lista_inventario = []
            for (diam, long), cant in inventario_final.items():
                # Lógica para determinar el tipo de material
                # Busca si la longitud existe en el inventario original de ese diámetro
                inventario_original_diam = resultados.get(str(diam), {}).get('inventario_original', {})
                if long in inventario_original_diam:
                    tipo = "Inventario Original"
                else:
                    tipo = "Sobrante / Exceso Generado"
                
                lista_inventario.append({
                    'Tipo de Material': tipo,
                    'Diámetro': diam,
                    'Longitud': long,
                    'Cantidad': cant
                })

            # Creamos un DataFrame de pandas con la nueva columna
            df_inventario = pd.DataFrame(lista_inventario)
            
            # Reordenamos las columnas para que 'Tipo' aparezca primero
            df_inventario = df_inventario[['Tipo de Material', 'Diámetro', 'Longitud', 'Cantidad']]
            
            st.dataframe(df_inventario, use_container_width=True)
        else:
            st.info("No hay inventario final restante.")

    else:
        st.error("Ocurrió un error durante la optimización. Revisa los logs.")
        with st.expander("Ver registro del proceso"):
            st.code('\n'.join(logs))
else:
    st.info("Por favor, sube un archivo de Excel para comenzar el proceso de optimización.")



