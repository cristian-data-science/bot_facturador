from datetime import date, timedelta
import pandas as pd
import requests
import asyncio
from datetime import date, timedelta
import pandas as pd
from dotenv import load_dotenv
import os
import streamlit as st
import time

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TOKEN_URL = os.getenv("TOKEN_URL")
SCOPE_URL = os.getenv("SCOPE_URL")

async def login_d365(playwright, url, user, passw, browser):
    context = await browser.new_context()
    await context.tracing.start(screenshots=True, snapshots=True, sources=False)
    
    page = await context.new_page()
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
    await context.tracing.stop(path=os.path.join(os.getcwd(), 'trace.zip'))

    await asyncio.sleep(2)
    await page.get_by_label("Buscar").click()
    await page.wait_for_selector('role=textbox[name="Buscar una página"]', state='visible')
    await page.get_by_role("textbox", name="Buscar una página").fill("todos los pedidos de com")

    await page.wait_for_selector('role=option[name="Todos los pedidos de compra Proveedores > Pedidos de compra"]', state='visible')
    await page.get_by_role("option", name="Todos los pedidos de compra Proveedores > Pedidos de compra").click()
    await asyncio.sleep(2)

    await page.wait_for_selector('role=button[name=" Nuevo"]', state='visible')
    await page.get_by_role("button", name=" Nuevo").click()    
    await asyncio.sleep(10)

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

        await download.save_as(os.path.join(os.getcwd(), "ReporteEmitidos_Det.xls"))

def process_report():
    archivo_original = os.path.join(os.getcwd(), 'ReporteEmitidos_Det.xls')
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

    archivo_resultado_final = os.path.join(os.getcwd(), 'formato_tabular.xlsx')
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
        else:
            st.error(f"Error en la solicitud: {response.status_code}")

def verificar_folio_en_erp(folio):
    token = obtener_token()
    if not token:
        return f"Error al obtener el token para el folio {folio}"

    nombre_archivo = os.path.join(os.getcwd(), 'folios_nocreados.xlsx')
    if os.path.exists(nombre_archivo):
        os.remove(nombre_archivo)
        st.success(f"Archivo anterior {nombre_archivo} borrado.")

    folios_no_creados = []

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
            return st.success(f"El folio 46-{folio} ya está creado y facturado en el diario de compras.")
        else:
            return st.error(f"El folio 46-{folio} no está creado en el ERP.")
            folios_no_creados.append(folio)
    else:
        return st.error(f"Error en la consulta del folio {folio}: {response.text}")

    if folios_no_creados:
        df_folios_no_creados = pd.DataFrame(folios_no_creados, columns=['Folios No Creados'])
        df_folios_no_creados.to_excel(nombre_archivo, index=False)
        st.success(f"Folios no creados guardados en '{nombre_archivo}'.")

    else:
        st.info("No se creo ningún PAT porque los folios ya se encuentran facturados en el ERP.")
