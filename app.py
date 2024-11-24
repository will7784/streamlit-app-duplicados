import os
import sys
import streamlit as st
import pandas as pd
from io import BytesIO

def export_excel(df, filename):
    """Exporta un DataFrame a Excel con formato y descarga el archivo."""
    # Ordenar por la columna PK_RA antes de exportar
    df = df.sort_values(by="PK_RA")

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Escribir DataFrame al archivo Excel
        df.to_excel(writer, index=False, sheet_name="Duplicados")
        
        # Obtener el workbook y worksheet de xlsxwriter
        workbook = writer.book
        worksheet = writer.sheets["Duplicados"]
        
        # Formato para el encabezado (negrita)
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        
        # Formato para las celdas del rango de datos (con bordes)
        cell_format = workbook.add_format({'border': 1})
        
        # Ajustar el ancho de las columnas y aplicar formato
        for col_num, column_name in enumerate(df.columns):
            # Establecer ancho de columna
            max_width = max(len(str(column_name)), df[column_name].astype(str).map(len).max())
            worksheet.set_column(col_num, col_num, max_width + 2)  # Agregar espacio extra
            
            # Aplicar formato al encabezado
            worksheet.write(0, col_num, column_name, header_format)
        
        # Aplicar formato a las celdas
        for row_num in range(1, len(df) + 1):
            for col_num in range(len(df.columns)):
                worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], cell_format)
    
    output.seek(0)
    st.download_button(
        label="Descargar Excel con duplicados",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def main():
    if getattr(sys, 'frozen', False):
        # Si el programa se ejecuta como un ejecutable
        os.system(f"streamlit run {sys.executable}")
    else:
        # Configurar el ancho completo para la página de Streamlit
        st.set_page_config(layout="wide")

        st.title("Carga y Visualización de Excel Control Conservador de Granero")

        # Cargar el archivo Excel
        uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])
        
        if uploaded_file:
            # Leer el archivo Excel
            try:
                df = pd.read_excel(uploaded_file, dtype=str)  # Leer todo como texto (strings)
                
                # Seleccionar columnas específicas
                columns_to_keep = [
                    'FOJAS_RA', 'NÚMERO_RA', 'AÑO_RA', 
                    'FOJAS_GRA', 'NÚMERO_GRA', 'AÑO_GRA'
                ]
                df = df[columns_to_keep]
                
                # Reemplazar NaN por cadenas vacías para evitar problemas al concatenar
                df = df.fillna("")
                
                # Crear claves primarias (PK) siguiendo el orden FOJAS + NÚMERO + AÑO
                df['PK_RA'] = df['FOJAS_RA'] + df['NÚMERO_RA'] + df['AÑO_RA']
                df['PK_GRA'] = df['FOJAS_GRA'] + df['NÚMERO_GRA'] + df['AÑO_GRA']
                
                # Mostrar el DataFrame con título personalizado
                st.subheader("Dataframe Control Excel")
                st.dataframe(df, use_container_width=True)
                
                # Identificar duplicados en la columna PK_RA
                duplicados = df[df.duplicated(subset='PK_RA', keep=False)]  # Todas las filas duplicadas
                
                # Botón para exportar duplicados
                st.markdown("---")
                if not duplicados.empty:
                    st.subheader("Duplicados Encontrados")
                    st.dataframe(duplicados, use_container_width=True)
                    if st.button("Exportar duplicados a Excel"):
                        export_excel(duplicados, "duplicados_PK_RA.xlsx")
                else:
                    st.info("No se encontraron duplicados en la columna PK_RA.")
                    st.button("No hay duplicados para exportar")  # Botón desactivado visualmente para mantener consistencia
                
            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")
        else:
            st.info("Por favor, sube un archivo Excel.")

if __name__ == "__main__":
    main()
