import os
import streamlit as st
import subprocess
import requests
import pandas as pd
from streamlit_extras.add_vertical_space import add_vertical_space
from streamlit_option_menu import option_menu
from streamlit_lottie import st_lottie
from streamlit.components.v1 import html
from time import sleep
from funcs_async import verificar_rut, verificar_folio_en_erp
import tempfile

st.set_page_config(page_title="App Template", layout="wide")

def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

loti1 = 'https://lottie.host/cf12dbd5-1b73-4f28-9b21-c7cbe4f458c0/6BbHPnunDG.json'
lot1 = load_lottieurl(loti1)
loti2= "https://lottie.host/eaa313d1-01a3-4e07-b3ee-93a67caa556f/GOxMTnclAj.json"
lot2= load_lottieurl(loti2)

def run_playwright_script(task_name):
    venv_python_path = os.path.join(os.getcwd(), "venv/Scripts/python.exe")
    script_path = os.path.join(os.getcwd(), "app_async.py")
    
    result = subprocess.run([venv_python_path, script_path, task_name], capture_output=True, text=True)
    return result.stdout, result.stderr

def main():
    col1 = st.sidebar
    col2, col3 = st.columns((4, 1))

    with col1:
        with st.sidebar:
            logo_path = os.path.join(os.getcwd(), "img/logo_1.png")
            st.image(logo_path, width=300)
            selected = option_menu("Main Menu", ["Home", 'Preparar datos', 'Validar proveedores', 'Facturar'],
                                   icons=['house', 'bi bi-upload', 'bi bi-download'],
                                   menu_icon="cast", default_index=0)

        add_vertical_space(3)
        st.write('Made with ❤️ by [Criss](https://github.com/cristian-data-science)')

    if selected == "Home":
        show_home(col1, col2)
    elif selected == "Preparar datos":
        preparar_datos(col2, "Preparar datos")
    elif selected == "Facturar":
        facturar(col2, "Facturar")
    elif selected == "Validar proveedores":
        validar_datos(col2, "Validar proveedores")

def show_home(col1, col2):
    with col2:
        col2.title("Auto facturador de pedidos de compra")
        st_lottie(lot1, key="loti1", height=700, width=780)

def preparar_datos(col2, page_name):
    with col2:
        col2.title("Auto facturador de pedidos de compra")
        st.write("Subir archivo de folios de compra")
        uploaded_file = st.file_uploader("Choose a file", type=['xlsx'])

        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            folios = df["Folio"].nunique()
            st.write(f"Se han encontrado {folios} folios únicos en el archivo subido.")
            folios_list = df["Folio"].unique()
            folios_list = [int(folio) for folio in folios_list]
            folios_list_str = ', '.join(str(folio) for folio in folios_list)
            st.write(f"Folios únicos en el archivo subido: {folios_list_str}")

            if st.button("Buscar líneas de pedido en blueline"):
                with st.spinner('🚀 Buscando líneas de pedido en blueline...'):
                    run_playwright_script("download_report")
                    run_playwright_script("process_report")

                    tabular_path = os.path.join(temp_dir, "formato_tabular.xlsx")
                    df_tabular = pd.read_excel(tabular_path)
                    matched_df = df_tabular[df_tabular["FOLIO"].isin(folios_list)]
                    matched_df = matched_df.drop(columns=['TRACKID', 'ESTADO PORTAL', 'ESTADO SII'])
                    st.write(matched_df)
                    temp_dir = tempfile.gettempdir()
                    output_path = os.path.join(temp_dir, "lineas_blueline.xlsx")
                    matched_df.to_excel(output_path, index=False)
                    st.write("Archivo 'lineas_blueline.xlsx' creado en la carpeta temporal")
                    matched_folios_count = matched_df["FOLIO"].nunique()
                    st.info(f"Se han encontrado {matched_folios_count} folios únicos en las líneas de pedido coincidentes en blueline.")

def validar_datos(col2, page_name):
    with col2:
        col2.title("Auto facturador de pedidos de compra")
        add_vertical_space(6)
        
        # Obtener la ruta de la carpeta temporal
        temp_dir = tempfile.gettempdir()
        blueline_path = os.path.join(temp_dir, "lineas_blueline.xlsx")
        
        # Leer el archivo desde la carpeta temporal
        if os.path.exists(blueline_path):
            lineas = pd.read_excel(blueline_path)
            
        else:
            st.error("El archivo 'lineas_blueline.xlsx' no se encontró en la carpeta temporal.")


        rut_receptor = lineas["RUT RECEPTOR"].nunique()
        st.info(f"Se han encontrado {rut_receptor} RUTs únicos en las lineas.")

        if st.button("Validar proveedores en ERP"):
            with st.spinner('🚀 Validando proveedores...'):
                ruts = lineas["RUT RECEPTOR"].unique().tolist()
                verificar_rut(ruts)

def facturar(col2, page_name):
    with col2:
        col2.title("Auto facturador de pedidos de compra")
        st.info("Antes de facturar revisaremos que los folios no estén facturados en el ERP")

        temp_dir = tempfile.gettempdir()
        blueline_path = os.path.join(temp_dir, "lineas_blueline.xlsx")
        folios_no_creados_path = os.path.join(temp_dir, 'folios_nocreados.xlsx')

        if st.button("Validar folios en ERP"):
            lineas = pd.read_excel(blueline_path)

            with st.spinner('🚀 Revisando si las facturas ya están registradas...'):
                folios = lineas["FOLIO"].unique().tolist()

                # Borrar archivo antiguo de folios no creados en la carpeta temporal
                if os.path.exists(folios_no_creados_path):
                    os.remove(folios_no_creados_path)

                folios_no_creados = []
                for folio in folios:
                    mensaje = verificar_folio_en_erp(folio)
                    if "no está creado" in mensaje.lower():
                        folios_no_creados.append(folio)

                if folios_no_creados:
                    df_folios_no_creados = pd.DataFrame(folios_no_creados, columns=["Folio"])
                    df_folios_no_creados.to_excel(folios_no_creados_path, index=False)
                    st.info(f"Se encontraron {len(folios_no_creados)} folios no creados.")
                    st.session_state.folios_no_creados_path = folios_no_creados_path

        if 'folios_no_creados_path' in st.session_state:
            if st.button("Crear pedidos no facturados!"):
                with st.spinner('🚀 Iniciando la secuencia de recopilación de datos. Lanzamiento hacia Blueline en 3... 2... 1...'):
                    # Paso 0: Borrar el archivo antiguo lineas_a_crear.xlsx si existe
                    temp_dir = tempfile.gettempdir()
                    lineas_a_crear_path = os.path.join(temp_dir, "lineas_a_crear.xlsx")
                    if os.path.exists(lineas_a_crear_path):
                        os.remove(lineas_a_crear_path)

                    # Paso 1 y 2: Crear el archivo lineas_a_crear.xlsx con las líneas que coincidan con los folios no creados
                    try:
                        folios_no_creados_path = st.session_state.folios_no_creados_path
                        df_folios_no_creados = pd.read_excel(folios_no_creados_path)
                        folios_no_creados = df_folios_no_creados["Folio"].tolist()

                        blueline_path = os.path.join(temp_dir, "lineas_blueline.xlsx")
                        lineas = pd.read_excel(blueline_path)
                        lineas_a_crear = lineas[lineas["FOLIO"].isin(folios_no_creados)]
                        lineas_a_crear.to_excel(lineas_a_crear_path, index=False)

                        #st.write(f"Archivo 'lineas_a_crear.xlsx' creado en la carpeta temporal: {lineas_a_crear_path}")
                        #st.info(f"Se han encontrado {len(lineas_a_crear)} líneas para los folios no creados en blueline.")
                    except FileNotFoundError:
                        st.error("No hay nada para facturar. Los pedidos ya estan creados en el ERP.")
                        return
                    # Ejecutar el script de Playwright después de crear el archivo
                    stdout, stderr = run_playwright_script("login_d365")
                    if stdout:
                        st.text("Salida:")
                        st.write(stdout)
                        st.success('Se han factuados los siguientes pedidos de venta!')
                        
                            # leer desde temp_dir pat_folios_creados.xlsx
                        pat_folios_creados = os.path.join(temp_dir, "pat_folios_creados.xlsx")
                        folios_creados = pd.read_excel(pat_folios_creados)
                        # mostrar df
                        st.write(folios_creados)
                        

                    if stderr:
                        pat_folios_creados = os.path.join(temp_dir, "pat_folios_creados.xlsx")
                        folios_creados = pd.read_excel(pat_folios_creados)
                        # mostrar df
                        st.write("##Folios parcialmente creados##")   
                        st.write(folios_creados)
                       
                        with st.expander("Mostrar errores"):
                            st.write("Errores")
                            st.error(stderr)

if __name__ == "__main__":
    main()