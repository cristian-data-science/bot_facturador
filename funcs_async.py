from datetime import date, timedelta
import pandas as pd
import requests
import asyncio
from dotenv import load_dotenv
import os
import csv
import streamlit as st
import time
import tempfile
from openpyxl import load_workbook
import openpyxl

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TOKEN_URL = os.getenv("TOKEN_URL")
SCOPE_URL = os.getenv("SCOPE_URL")

async def login_d365(playwright, url, user, passw, browser):
    context = await browser.new_context()
    await context.tracing.start(screenshots=True, snapshots=True, sources=False)

    
    page = await context.new_page()
    # Ajustar el tamaño de la ventana a 1920x1080 (Full HD)
    await page.set_viewport_size({"width": 1400, "height": 850})
    try:
        await page.goto(url)
        await asyncio.sleep(2)

        await page.wait_for_selector('input[name="loginfmt"]', state="visible")
        await page.fill('input[name="loginfmt"]', user)
        await page.press('input[name="loginfmt"]', 'Enter')
        await page.wait_for_selector('input[name="passwd"]', state="visible")
        await asyncio.sleep(2)
        await page.fill('input[name="passwd"]', passw)
        await page.click('input[type="submit"]')
        await page.wait_for_selector('input[name="DontShowAgain"]', state="visible")
        await page.click('input[name="DontShowAgain"]')
        await page.click('input[type="submit"]')

        

    except Exception as e:
        print(f"Error: {e}")
    
    
    

    #leer folios no creados en una sola linea
    temp_dir = tempfile.gettempdir()
    folios_no_creados = pd.read_excel(os.path.join(temp_dir, 'folios_nocreados.xlsx'))
    folios_no_creados = folios_no_creados['Folio'].tolist()
    # order to minor to major
    folios_no_creados.sort()



    # Obtener una ruta en la carpeta temporal
    temp_dir = tempfile.gettempdir()
    pat_folios_creados = os.path.join(temp_dir, 'pat_folios_creados.csv')
    #borrar archivo si existe
    if os.path.exists(pat_folios_creados):
        os.remove(pat_folios_creados)

    # para gaurdar en el for
    datos = []


    for folio in folios_no_creados:
        #await asyncio.sleep(2)
        await asyncio.sleep(2)
        
        await page.mouse.move(0, 50)
        await page.get_by_label("Buscar", exact=True).click()
        await page.wait_for_selector('role=textbox[name="Buscar una página"]', state='visible')
        await page.get_by_role("textbox", name="Buscar una página").fill("todos los pedidos de com")

        await page.wait_for_selector('role=option[name="Todos los pedidos de compra Proveedores > Pedidos de compra"]', state='visible')
        await page.get_by_role("option", name="Todos los pedidos de compra Proveedores > Pedidos de compra").click()
        await asyncio.sleep(1)
        await page.keyboard.press("Enter")

        await page.wait_for_selector('role=button[name=" Nuevo"]', state='visible')
        await page.mouse.move(0, 50)
        await page.get_by_role("button", name=" Nuevo").click()
        await page.get_by_label("Crear pedido de compra").get_by_label("Cuenta de proveedor").click()

        # leer lineas a crear
        lineas_a_crear = pd.read_excel(os.path.join(temp_dir, 'lineas_a_crear.xlsx'))
        # filtrar lineas a crear por folio
        lineas_folio = lineas_a_crear[lineas_a_crear['FOLIO'] == folio].copy()

        # obtener el rut receptor de lineas_folio
        rut_receptor = lineas_folio['RUT RECEPTOR'].iloc[0]
        

        await page.get_by_label("Crear pedido de compra").get_by_label("Cuenta de proveedor").fill(rut_receptor)

        await page.get_by_label("Sitio").click()
        await page.get_by_label("Sitio").fill("01")

        await page.locator('input[name="PurchTable_InventLocationId"][role="combobox"][type="text"]').click()



        await page.locator('input[name="PurchTable_InventLocationId"][role="combobox"][type="text"]').press("CapsLock")
        await page.locator('input[name="PurchTable_InventLocationId"][role="combobox"][type="text"]').fill("OF.CENTRAL")


        # obtener fecha de folio actual en las lineas

        lineas_folio.loc[:, 'FECHA EMISION'] = pd.to_datetime(lineas_folio['FECHA EMISION'], errors='coerce')

        fecha_folio = lineas_folio['FECHA EMISION'].iloc[0].strftime('%d/%m/%Y')







        await page.get_by_role("combobox", name="Fecha contable").nth(0).click()
        await page.get_by_role("combobox", name="Fecha contable").nth(0).fill("")
        await page.get_by_role("combobox", name="Fecha contable").nth(0).fill(fecha_folio)
        #await asyncio.sleep(5)

        # Hacer scroll hacia abajo
        await page.evaluate("window.scrollBy(0, 100)")

        # Interactuar con el primer combobox "Fecha de recepción solicitada"
        await page.get_by_role("combobox", name="Fecha de recepción solicitada").nth(1).click()
        await page.get_by_role("combobox", name="Fecha de recepción solicitada").nth(1).fill(fecha_folio)
        #await asyncio.sleep(5)
        # press alt + enter
        await page.keyboard.press("Alt+Enter")

        # esperar 2 segundos
        await asyncio.sleep(2)

        await page.get_by_role("button", name=" Quitar").click()


        # hacer un for para cada linea del folio
        contador = 0

        for index, row in lineas_folio.iterrows():
            # Haz clic en el botón para agregar línea
            await page.get_by_role("button", name=" Agregar línea").click()
            #await asyncio.sleep(2)

            try:
                # Selecciona y llena el campo de código de artículo usando el contador
                await page.get_by_label("Código de artículo").nth(contador).click()
            except Exception as e:
                pass

            await page.keyboard.type(row['CODIGO'])

            #await page.get_by_label("Código de artículo").nth(contador).fill(row['CODIGO'])
            await asyncio.sleep(1)
            await page.keyboard.press("Tab")



            #await asyncio.sleep(1)
            """await page.get_by_role("gridcell", name="Ubicación Abrir").get_by_role("button").click()
            await asyncio.sleep(1)
            await page.get_by_role("gridcell", name="Ubicación Abrir").get_by_role("button").click()


            await asyncio.sleep(1)
            await page.get_by_label("Formularios de búsqueda").locator('label:has-text("Ubicación")').first.click()
            await asyncio.sleep(1)
            # llenar con "GENERICA" """
            await page.keyboard.type("GENERICA")
            await page.keyboard.press("Enter")


            
            await page.locator(f'xpath=//*[@id="GridCell-{contador}-PurchLine_PurchQtyGrid"]').click()
            #await asyncio.sleep(1)
            await page.locator(f'xpath=//*[@id="GridCell-{contador}-PurchLine_PurchQtyGrid"]').click()
            
            # borrar la cantidad actual con suprimir
            keys = ["Delete"] * 4 + ["Backspace"] * 4
            for key in keys:
                await page.keyboard.press(key)
                await asyncio.sleep(0.2)
            await page.keyboard.press("Enter")
            
            # ingresar cantidad str(row['CANTIDAD']) por teclado
            await assyncio.sleep(1)
            await page.keyboard.type(str(row['CANTIDAD']))
            await assyncio.sleep(1)
            await page.keyboard.press("Enter")
            #await asyncio.sleep(1)

            await page.locator(f'xpath=//*[@id="GridCell-{contador}-PurchLine_PurchPriceGrid"]').click()

            #await asyncio.sleep(1)
            await page.locator(f'xpath=//*[@id="GridCell-{contador}-PurchLine_PurchPriceGrid"]').click()

            keys = ["Delete"] * 10 + ["Backspace"] * 10
            for key in keys:
                await page.keyboard.press(key)
                await asyncio.sleep(0.1)

            await page.keyboard.type(str(row['PRECIO']))
            #await asyncio.sleep(1)
  
            await page.keyboard.press("Enter")

            


            # Incrementa el contador
            contador += 1

            await asyncio.sleep(2)
        
        await page.get_by_role("button", name="Compra", exact=True).click()
        await asyncio.sleep(1)
        await page.get_by_role("button", name="Confirmar").click()
        await page.wait_for_selector('text="Operación completada"', timeout=60000)

        await page.get_by_role("button", name="Factura").click()
        #await asyncio.sleep(1)  
        await page.get_by_text("Factura", exact=True).nth(1).click()




        #await asyncio.sleep(2)  
        # Localizador con los atributos especificados
        await page.get_by_role("combobox", name="Tipo de transacción").click()

        #await asyncio.sleep(1)
        await page.get_by_role("option", name="Local").click()


        await page.get_by_label("Talonario").click()
        await page.get_by_label("Talonario").fill("46")
        #await asyncio.sleep(1)
        await page.get_by_label("Número", exact=True).click()
        # fill con folio en string en formato 46-FOLIO
        await page.keyboard.type(f"46-{folio}")
        await page.get_by_label("Descripción de factura").click()
        ## llenar con "COMPRA WW BOT"
        await page.keyboard.type("COMPRA WW BOT")
        await page.keyboard.press("Enter")

        await page.get_by_role("combobox", name="Fecha de recepción de la").nth(0).click()
        await page.get_by_role("combobox", name="Fecha de recepción de la").nth(0).fill(fecha_folio)

        await page.get_by_role("combobox", name="Fecha de la factura").click()
        await page.get_by_role("combobox", name="Fecha de la factura").fill(fecha_folio)

        await page.get_by_role("combobox", name="Fecha de registro de IVA de").click()
        await page.get_by_role("combobox", name="Fecha de registro de IVA de").fill(fecha_folio)

        await page.get_by_role("combobox", name="Fecha de registro", exact=True).click()
        await page.get_by_role("combobox", name="Fecha de registro", exact=True).fill(fecha_folio)

        await page.get_by_role("combobox", name="Fecha del registro del IVA").nth(0).click()
        await page.get_by_role("combobox", name="Fecha del registro del IVA").nth(0).fill(fecha_folio)

        await page.get_by_role("combobox", name="Fecha de vencimiento").click()
        await page.get_by_role("combobox", name="Fecha de vencimiento").fill(fecha_folio)

        
        await asyncio.sleep(1)
        # send tab
        await page.keyboard.press("Tab")
        #send fecha_folio
        await page.keyboard.type(fecha_folio)




        await page.get_by_role("button", name="Actualizar estado de").click()

         
        await asyncio.sleep(1)
        await page.get_by_role("button", name="Registrar").click()

        #await page.get_by_role("button", name="Factura").nth(1).click()
  
        await page.get_by_text("Factura", exact=True).nth(2).click()

        
        await asyncio.sleep(1)

        pat = await page.locator('xpath=/html/body/div[2]/div/div[6]/div/form[3]/div[5]/div/div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/div[3]/div/div[2]/div/div/div/div[1]/div[3]/div/div/div[1]/div[2]/div/div[1]/div/div/div/div/div/div/input').get_attribute('value')
        texto_factura = await page.locator('xpath=/html/body/div[2]/div/div[6]/div/form[3]/div[5]/div/div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/div[3]/div/div[2]/div/div/div/div[1]/div[3]/div/div/div[1]/div[2]/div/div[5]/div/div/div/div/div/div/input').get_attribute('value')


        

        datos.append({'pat': pat, 'folio': texto_factura, 'fecha_factura': fecha_folio})

        # Verificar si el archivo existe para determinar si se deben escribir los encabezados
        file_exists = os.path.exists(pat_folios_creados)

        # Escribir datos en el archivo CSV
        with open(pat_folios_creados, mode='a', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=['pat', 'folio', 'fecha_factura'])
            
            if not file_exists:
                # Escribir encabezados si el archivo no existe
                writer.writeheader()
                
            # Escribir los datos actuales
            writer.writerow({'pat': pat, 'folio': texto_factura, 'fecha_factura': fecha_folio})


    print("### Proceso finalizado ###")
    
    # Cerrar el navegador
    await browser.close()






def rev_proveedor():
    pass

async def download_report(playwright, url, user, password, browser):
    browser = await playwright.chromium.launch(headless=False)
    context = await browser.new_context(accept_downloads=True)
    page = await context.new_page()
    print("### Iniciando sesión ###")
    await page.goto(url)
    await page.fill('input[name="sUser"]', user)
    await page.fill('input[type="password"]', password)
    await page.click('input[type="submit"]')
    await page.wait_for_load_state('networkidle')

    print("### Filtrando fechas y descargando reporte ###")
    await page.goto(url)

    start_date = (date.today() - timedelta(90)).strftime('%Y-%m-%d')
    end_date = date.today().strftime('%Y-%m-%d')

    await page.fill('input#FCHDESDE', start_date)
    await page.fill('input[name="FCHHASTA"]', end_date)
    await page.press('input[name="FCHHASTA"]', 'Enter')

    await page.locator("#DOCUMENTO").select_option("46")
    await page.locator("button[name=\"BUSCAR\"]").click()

    await page.wait_for_load_state('networkidle')

    async with page.expect_download() as download_info:
        await page.click('xpath=/html/body/fieldset/form/table[3]/tbody/tr[3]/td[2]/a')
        download = await download_info.value
        download_path = await download.path()

        temp_dir = tempfile.gettempdir()
        await download.save_as(os.path.join(temp_dir, "ReporteEmitidos_Det.xls"))

def process_report():
    temp_dir = tempfile.gettempdir()
    archivo_original = os.path.join(temp_dir, 'ReporteEmitidos_Det.xls')
    df_original = pd.read_excel(archivo_original, header=None)

    facturas_index = df_original.index[df_original[0].notna() & df_original[0].astype(str).str.isnumeric()].tolist()
    facturas_index.append(len(df_original))

    columnas_deseadas = ['FOLIO', 'TRACKID', 'ESTADO PORTAL', 'ESTADO SII', 'DOCUMENTO', 'SUCURSAL', 'FECHA EMISION', 'FECHA CARGA', 'RUT RECEPTOR', 'DESCRIPCION', 'CODIGO', 'CANTIDAD', 'PRECIO', 'TOTAL']
    df_transformado = pd.DataFrame(columns=columnas_deseadas)

    for i in range(len(facturas_index) - 1):
        datos_factura = df_original.iloc[facturas_index[i]]
        rut_receptor = datos_factura[8]
        estado_sii = datos_factura[3]
        folio = datos_factura[0]

        for j in range(facturas_index[i] + 1, facturas_index[i + 1]):
            fila_producto = df_original.iloc[j, 9:14]
            if pd.notna(fila_producto[9]) and not fila_producto[9].startswith('DESCRIPCION'):
                nueva_fila = pd.Series([folio, datos_factura[1], datos_factura[2], estado_sii, datos_factura[4], datos_factura[5], datos_factura[6], datos_factura[7], rut_receptor, fila_producto[9], fila_producto[10], fila_producto[11], fila_producto[12], fila_producto[13]], index=columnas_deseadas)
                df_transformado = pd.concat([df_transformado, pd.DataFrame([nueva_fila])], ignore_index=True)

    df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']] = df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']].apply(pd.to_numeric, errors='coerce')
    df_transformado['CANTIDAD'] = df_transformado['PRECIO']
    df_transformado['PRECIO'] = df_transformado['TOTAL']
    df_transformado['TOTAL'] = df_transformado['PRECIO'] * df_transformado['CANTIDAD']
    df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']] = df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']].apply(pd.to_numeric, errors='coerce')

    archivo_resultado_final = os.path.join(temp_dir, 'formato_tabular.xlsx')
    df_transformado.to_excel(archivo_resultado_final, index=False)

token_cache = {"token": None, "expiry": time.time()}

def obtener_token():
    current_time = time.time()
    if token_cache["token"] and token_cache["expiry"] > current_time:
        return token_cache["token"]
    
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": SCOPE_URL
    }
    response = requests.post(TOKEN_URL, headers=headers, data=data)
    if response.status_code == 200:
        token_duration = 3600
        token_cache["token"] = response.json()["access_token"]
        token_cache["expiry"] = current_time + token_duration
        return token_cache["token"]
    else:
        return None

def verificar_rut(ruts):
    token = obtener_token()
    if not token:
        st.error("Error al obtener el token")
        return

    for rut in ruts:
        base_url = "https://patagonia-prod.operations.dynamics.com/data/VendorsV3"
        filtro = f"$filter=VendorAccountNumber eq '{rut}'"
        url_completa = f"{base_url}?{filtro}&$top=1"

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        response = requests.get(url_completa, headers=headers)
        if response.status_code == 200:
            datos = response.json()
            if datos['value']:
                st.success(f"El RUT {rut} está registrado como proveedor.")
            else:
                st.error(f"El RUT {rut} no existe y fue eliminado.")
                # eliminar la linea con el rut no existente del archivo de lineas de la carpeta tempral
                temp_dir = tempfile.gettempdir()
                lineas_path = os.path.join(temp_dir, "lineas_blueline.xlsx")
                lineas = pd.read_excel(lineas_path)
                lineas = lineas[lineas["RUT RECEPTOR"] != rut]
                lineas.to_excel(lineas_path, index=False)

        else:
            st.error(f"Error en la solicitud: {response.status_code}")

def verificar_folio_en_erp(folio):
    token = obtener_token()
    if not token:
        return f"Error al obtener el token para el folio {folio}"

    temp_dir = tempfile.gettempdir()
    nombre_archivo = os.path.join(temp_dir, 'folios_nocreados.xlsx')

    # Inicializar un DataFrame vacío si no existe el archivo
    if os.path.exists(nombre_archivo):
        df_folios_no_creados = pd.read_excel(nombre_archivo)
    else:
        df_folios_no_creados = pd.DataFrame(columns=['Folio'])

    base_url = "https://patagonia-prod.operations.dynamics.com/data/VendInvoiceInfoTableBiEntities"
    filtro = f"$filter=Num eq '46-{folio}'"
    url_completa = f"{base_url}?{filtro}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url_completa, headers=headers)
    if response.status_code == 200:
        datos = response.json()
        if datos['value']:
            st.success(f"El folio 46-{folio} ya está creado y facturado en el diario de compras.")
            return f"El folio 46-{folio} ya está creado y facturado en el diario de compras."
        else:
            # Crear un nuevo DataFrame con el folio no creado y concatenarlo al DataFrame existente
            nuevo_folio_df = pd.DataFrame({'Folio': [folio]})
            df_folios_no_creados = pd.concat([df_folios_no_creados, nuevo_folio_df], ignore_index=True)
            df_folios_no_creados.to_excel(nombre_archivo, index=False)
            st.error(f"El folio 46-{folio} no está creado en el ERP.")
            return f"El folio 46-{folio} no está creado en el ERP."
    else:
        st.error(f"Error en la consulta del folio {folio}: {response.text}")
        return f"Error en la consulta del folio {folio}: {response.text}"