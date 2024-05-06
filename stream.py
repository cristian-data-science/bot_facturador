import streamlit as st
import subprocess
import requests

#from streamlit_extras.add_vertical_space import add_vertical_space
from streamlit_option_menu import option_menu
from streamlit_lottie import st_lottie
from streamlit.components.v1 import html
from time import sleep

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



def run_playwright_script():
    # Ruta al ejecutable de Python dentro del entorno virtual
    venv_python_path = "C:/Users/cguti/OneDrive - Patagonia/Cristian/git/playwright_proyects/facturador_compras/venv/Scripts/python.exe"
    # Ruta al script app.py que deseas ejecutar
    script_path = "C:/Users/cguti/OneDrive - Patagonia/Cristian/git/playwright_proyects/facturador_compras/app_async.py"
    
    # Ejecutar el script y capturar la salida
    result = subprocess.run([venv_python_path, script_path], capture_output=True, text=True)
    
    # Devolver la salida estándar y la salida de error estándar
    return result.stdout, result.stderr


def main():

    """st.title("Mi aplicación con Streamlit")
    
    if st.button("Ejecutar Playwright"):
        stdout, stderr = run_playwright_script()
        if stdout:
            st.text("Salida:")
            st.write(stdout)
        if stderr:
            st.text("Errores:")
            st.error(stderr)"""

    col1 = st.sidebar
    col2, col3 = st.columns((4, 1))

    with col1:      
        #st_lottie(lot2, key="lol",height=180, width=280)      
        with st.sidebar:
            st.image("./img/logo_1.png", width=300)
            selected = option_menu("Main Menu", ["Home", 'Preparar datos','Validar pedidos', 'Facturar'],
                               icons=['house', 'bi bi-upload', 'bi bi-download'],
                               menu_icon="cast", default_index=0)

        #add_vertical_space(3)
        st.write('Made with ❤️ by [Criss](https://github.com/cristian-data-science)')

    if selected == "Home":
        show_home(col1, col2)
    elif selected == "Preparar datos":
        preparar_datos(col2, "Preparar datos")
    elif selected == "Facturar":
        facturar(col2, "Facturar")
    elif selected == "Validar pedidos":
        facturar(col2, "Validar pedidos")


def show_home(col1, col2):
    with col2:  
        #st.image("banner-preview.png",width=180)

        col2.title("Auto facturador de pedidos de compra")
        st_lottie(lot1, key="loti1", height=700, width=780) 

def preparar_datos(col2, page_name):
    with col2:
 
        col2.title("Auto facturador de pedidos de compra")
        #agregar boton para leer excel y posteriormente que lea el campo folio y lo guarde con pandas en una lista llamada folio        

        
    
    

        

def facturar(col2, page_name):
    with col2:
      
        col2.title("Auto facturador de pedidos de compra")

        with st.spinner('🚀 Iniciando la secuencia de recopilación de datos. Lanzamiento hacia Blueline en 3... 2... 1...'):
            # Tu código para obtener datos va aquí
            if st.button("Comandar a los Bots!"):
                stdout, stderr = run_playwright_script()
                if stdout:
                    st.text("Salida:")
                    st.write(stdout)
                    st.success('¡Aterrizaje exitoso! Los datos han sido capturados.')
                if stderr:
                    st.text("Errores:")
                    st.error(stderr)

            





if __name__ == "__main__":
    main()

