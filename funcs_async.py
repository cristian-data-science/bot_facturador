from datetime import date, timedelta
import pandas as pd

import asyncio
from datetime import date, timedelta
import pandas as pd

async def login_d365(playwright, url, user, password, browser):

    browser = await playwright.chromium.launch(headless=False)
    context = await browser.new_context()
    await context.tracing.start(screenshots=True, snapshots=True, sources=True)
    
    page = await context.new_page()
    # add try except block to handle errors with 20 seconds to timeout  

    try:
        await page.goto(url)
        await asyncio.sleep(2)
        await page.wait_for_selector('input[name="loginfmt"]', state="visible")
        await page.fill('input[name="loginfmt"]', user)
        await page.press('input[name="loginfmt"]', 'Enter')
        await page.wait_for_selector('input[name="passwd"]', state="visible")
        await asyncio.sleep(2)
        await page.fill('input[name="passwd"]', password)
        await page.click('input[type="submit"]')
        await page.wait_for_selector('input[name="DontShowAgain"]', state="visible")
        await page.click('input[name="DontShowAgain"]')
        await page.click('input[type="submit"]')

    except Exception as e:
        print(f"Error: {e}")
    await context.tracing.stop(path='trace.zip')

    # Asumiendo que `page` ya está definido y que estás utilizando Playwright
    # Haz clic en "Buscar" y espera hasta que el cuadro de texto esté visible
    await asyncio.sleep(2)
    await page.get_by_label("Buscar").click()
    await page.wait_for_selector('role=textbox[name="Buscar una página"]', state='visible')
    await page.get_by_role("textbox", name="Buscar una página").fill("todos los pedidos de com")

    # Espera explícita por la opción antes de hacer clic
    await page.wait_for_selector('role=option[name="Todos los pedidos de compra Proveedores > Pedidos de compra"]', state='visible')
    await page.get_by_role("option", name="Todos los pedidos de compra Proveedores > Pedidos de compra").click()
    await asyncio.sleep(2)
    # Espera a que la navegación a la nueva URL se complete antes de proceder
    await page.goto("https://patagonia-test.sandbox.operations.dynamics.com/?cmp=PAT&mi=PurchTableListPage", wait_until='networkidle')

    # Espera al botón "Nuevo" antes de hacer clic
    await page.wait_for_selector('role=button[name=" Nuevo"]', state='visible')
    await page.get_by_role("button", name=" Nuevo").click()

    # Rellenar el combobox y esperar a que la opción esté lista para ser seleccionada
    await page.wait_for_selector('role=combobox[name="Cuenta de proveedor"]', state='visible')
    await page.get_by_role("combobox", name="Cuenta de proveedor").fill("18161379-3")

    # Es crucial esperar a que la interfaz de usuario reaccione a las interacciones anteriores
    await page.wait_for_selector('text="Empresas proveedoras"', state='visible')
    await page.get_by_label("Empresas proveedoras").get_by_label("Cuenta de proveedor").click()

    await page.wait_for_selector('role=button[name="Aceptar"]', state='visible')
    await page.get_by_role("button", name="Aceptar").click()

    # Esperar a que la navegación se complete después de redirigir
    await page.goto("https://patagonia-test.sandbox.operations.dynamics.com/?cmp=PAT&mi=PurchTableListPage", wait_until='networkidle')

    # Esperar al botón "Agregar línea" antes de hacer clic
    await page.wait_for_selector('role=button[name=" Agregar línea"]', state='visible')
    await page.get_by_role("button", name=" Agregar línea").click()

    # Rellenar el combobox de "Código de artículo" y seleccionar la opción deseada
    await page.wait_for_selector('role=combobox[name="Código de artículo"]', state='visible')
    await page.get_by_role("combobox", name="Código de artículo").fill("WW")
    await page.locator("#InventTableExpanded_ItemId_9727_0_2_input").click()

    # La última interacción debe esperar también a que el elemento esté listo
    await page.wait_for_selector('.ScrollbarLayout_face', state='visible')
    await page.locator(".ScrollbarLayout_face").first.click()


def rev_proveedor():
    pass

async def download_report(playwright, url, user, password, browser):

    browser = await playwright.chromium.launch(headless=True)
    context = await browser.new_context(accept_downloads=True)
    #context = await browser.new_context(accept_downloads=True)
    page = await context.new_page()
    print("### Iniciando sesión ###")
    await page.goto(url)
    await page.fill('input[name="sUser"]', user)
    await page.fill('input[type="password"]', password)
    await page.click('input[type="submit"]')
    await page.wait_for_load_state('networkidle')

    print("### Filtrando fechas y descargando reporte ###")
    await page.goto(url)

    start_date = (date.today() - timedelta(30)).strftime('%Y-%m-%d')
    end_date = date.today().strftime('%Y-%m-%d')

    await page.fill('input#FCHDESDE', start_date)
    await page.fill('input[name="FCHHASTA"]', end_date)
    await page.press('input[name="FCHHASTA"]', 'Enter')

    await page.locator("#DOCUMENTO").select_option("46")
    await page.locator("button[name=\"BUSCAR\"]").click()

    await page.wait_for_load_state('networkidle')

    # Uso correcto de async with para manejar eventos asíncronos
    async with page.expect_download() as download_info:
        await page.click('xpath=/html/body/fieldset/form/table[3]/tbody/tr[3]/td[2]/a')
        download = await download_info.value
        download_path = await download.path()

        await download.save_as("./ReporteEmitidos_Det.xls")


def process_report():
    archivo_original = 'ReporteEmitidos_Det.xls'
    df_original = pd.read_excel(archivo_original, header=None)

    # Identificar las filas que corresponden al inicio de las facturas
    facturas_index = df_original.index[df_original[0].notna() & df_original[0].astype(str).str.isnumeric()].tolist()
    facturas_index.append(len(df_original))  # Agregar el final del archivo como índice

    # Crear un nuevo DataFrame para los resultados
    columnas_deseadas = ['FOLIO', 'TRACKID', 'ESTADO PORTAL', 'ESTADO SII', 'DOCUMENTO', 'SUCURSAL', 'FECHA EMISION', 'FECHA CARGA', 'RUT RECEPTOR', 'DESCRIPCION', 'CODIGO', 'CANTIDAD', 'PRECIO', 'TOTAL']
    df_transformado = pd.DataFrame(columns=columnas_deseadas)

    # Procesar cada factura y sus productos
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

    # Conversión de tipos para 'CANTIDAD', 'PRECIO' y 'TOTAL'
    df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']] = df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']].apply(pd.to_numeric, errors='coerce')

    # Mover los valores de las columnas como se requiere
    df_transformado['CANTIDAD'] = df_transformado['PRECIO']
    df_transformado['PRECIO'] = df_transformado['TOTAL']

    # Calcular el nuevo valor de TOTAL como PRECIO * CANTIDAD
    df_transformado['TOTAL'] = df_transformado['PRECIO'] * df_transformado['CANTIDAD']

    # Asegurarse de que los tipos de datos sean correctos
    df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']] = df_transformado[['CANTIDAD', 'PRECIO', 'TOTAL']].apply(pd.to_numeric, errors='coerce')

    # Guardar el DataFrame transformado en un nuevo archivo Excel
    archivo_resultado_final = 'formato_tabular.xlsx'
    df_transformado.to_excel(archivo_resultado_final, index=False)