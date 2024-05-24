import streamlit as st
import pandas as pd
import tempfile
import os
import subprocess
from dotenv import load_dotenv
from funcs_async import verificar_folio_en_erp

def run_playwright_script(task_name, environment_url=None):
    venv_python_path = os.path.join(os.getcwd(), "venv/Scripts/python.exe" if os.name == 'nt' else "venv/bin/python")
    script_path = os.path.join(os.getcwd(), "app_async.py")

    if environment_url:
        result = subprocess.run([venv_python_path, script_path, task_name, '--environment_url', environment_url], capture_output=True, text=True)
    else:
        result = subprocess.run([venv_python_path, script_path, task_name], capture_output=True, text=True)

    return result.stdout, result.stderr

def facturar(col2, page_name):
    with col2:
        col2.title("Auto facturador de pedidos de compra")
        st.info("Antes de facturar revisaremos que los folios no estén facturados en el ERP")

        environment = st.radio(
            "Selecciona el ambiente",
            ("Ambiente de test", "Ambiente de producción")
        )

        environment_url = "url_erp_test" if environment == "Ambiente de test" else "URL_PROD"

        temp_dir = tempfile.gettempdir()
        blueline_path = os.path.join(temp_dir, "lineas_blueline.xlsx")
        folios_no_creados_path = os.path.join(temp_dir, 'folios_nocreados.xlsx')

        if st.button("Validar folios en ERP"):
            lineas = pd.read_excel(blueline_path)

            with st.spinner('🚀 Revisando si las facturas ya están registradas...'):
                folios = lineas["FOLIO"].unique().tolist()

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
                with st.spinner('🚀 Creando los pedidos de venta para los folios...'):
                    temp_dir = tempfile.gettempdir()
                    lineas_a_crear_path = os.path.join(temp_dir, "lineas_a_crear.xlsx")
                    if os.path.exists(lineas_a_crear_path):
                        os.remove(lineas_a_crear_path)

                    folios_no_creados_path = st.session_state.folios_no_creados_path
                    df_folios_no_creados = pd.read_excel(folios_no_creados_path)
                    folios_no_creados = df_folios_no_creados["Folio"].tolist()
                    blueline_path = os.path.join(temp_dir, "lineas_blueline.xlsx")
                    lineas = pd.read_excel(blueline_path)
                    lineas_a_crear = lineas[lineas["FOLIO"].isin(folios_no_creados)]
                    lineas_a_crear.to_excel(lineas_a_crear_path, index=False)

                    stdout, stderr = run_playwright_script("login_d365", environment_url)
                    if stdout:
                        st.text("Salida:")
                        st.write(stdout)
                        st.success('Se han factuados los siguientes pedidos de venta!')
                        
                        pat_folios_creados = os.path.join(temp_dir, "pat_folios_creados.csv")
                        folios_creados = pd.read_csv(pat_folios_creados)
                        st.write(folios_creados)
                        
                    if stderr:
                        pat_folios_creados = os.path.join(temp_dir, "pat_folios_creados.csv")
                        folios_creados = pd.read_csv(pat_folios_creados)
                        st.write("##Folios parcialmente creados##")
                        st.write(folios_creados)
                        
                        with st.expander("Mostrar errores"):
                            st.write("Errores")
                            st.error(stderr)

if __name__ == "__main__":
    col1, col2 = st.columns((1, 3))
    facturar(col2, "Facturar")
