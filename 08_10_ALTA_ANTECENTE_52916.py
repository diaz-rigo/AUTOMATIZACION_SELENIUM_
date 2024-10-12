import requests
from io import BytesIO
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.firefox import GeckoDriverManager
import time
ENTRADA_CONSTANTE = '' #jrodriguez_scontino@ifreh.gob.mx
RANGO_CARGA_INICIAL =    0         # 757
RANGO_CARGA_FINALIZAR =  0


CORREO_CONSTANTE = ''  # Initialize CORREO_CONSTANTE
YEAR = ''
TOMO = ''
LIBRO = ''
VOLUMEN = ''
INSCRIPCION = ''


# Paso 1: Descargar el archivo de Google Sheets en formato Excel
sheet_url = "https://docs.google.com/spreadsheets/d/1m2wOgUVjvaYiCRomQxD1WBEElKCHC8zu/export?format=xlsx"
response = requests.get(sheet_url)

if response.status_code == 200:
    excel_data = response.content
    # Leer el archivo Excel usando Pandas
    df = pd.read_excel(BytesIO(excel_data), sheet_name='Activos', engine='openpyxl')
    print("Archivo Excel cargado con éxito")

    # Asegurarnos de que la columna 'ENTRADA' sea de tipo string
    df['ENTRADA'] = df['ENTRADA'].astype(str).str.strip()

    # Filtrar las filas donde la columna 'ENTRADA' sea igual a '52867'
    filtered_df = df[df['ENTRADA'] == ENTRADA_CONSTANTE]
    print(filtered_df.columns)
    print("Filtrado de filas donde 'ENTRADA' es:", ENTRADA_CONSTANTE)
    print(filtered_df)

    if not filtered_df.empty:
        # Almacenar el primer valor de la columna de correos
        CORREO_CONSTANTE = str(filtered_df['FOLIO REAL.1'].values[0])

        print(f"El primer correo es: {CORREO_CONSTANTE}")
    else:
        print(f"No se encontraron filas con 'ENTRADA' igual a {ENTRADA_CONSTANTE}")
else:
    print(f"Error al descargar el archivo: {response.status_code}")

try:
    # Inicia el navegador y accede a la página de login
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
    driver.get("https://ifreh-s.hidalgo.gob.mx:8443/erpp/#/login")
    print("Página abierta con éxito")

    # Esperar hasta que el campo de Username sea visible (máx 30 segundos)
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, 'username')))

    # Ingresar correo y contraseña
    driver.find_element(By.ID, 'username').send_keys(CORREO_CONSTANTE)
    driver.find_element(By.ID, 'password').send_keys('admin')

    # Iniciar sesión
    driver.find_element(By.XPATH, '/html/body/jhi-main/div/div/login-form/div/div/div[3]/form/button').click()
    print("Credenciales ingresadas y botón de inicio de sesión presionado")
    time.sleep(3) 
        # Esperar un máximo de 10 segundos para verificar si aparece el enlace de desbloqueo
    try:
        unlock_link = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.XPATH, '/html/body/jhi-main/div/div/login-form/div/div/div[2]/div/a'))
        )
        # Si aparece el enlace, hacer clic
        unlock_link.click()
        print("Enlace de desbloqueo encontrado y clickeado")

        # Esperar a que aparezca el campo de la contraseña
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.XPATH, '//*[@id="password"]'))
        )

        # Ingresar la contraseña "admin"
        driver.find_element(
            By.XPATH, '//*[@id="password"]').send_keys('admin')
        print("Contraseña 'admin' ingresada")

        # Esperar hasta que el botón de desbloqueo esté visible y luego hacer clic
        unlock_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-user-unlock/div/div/div[2]/form/button'))
        )
        unlock_button.click()
        print("Botón de desbloqueo presionado")

        # Esperar hasta que el enlace esté visible y luego hacer clic
        link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-user-unlock/div/div/div/div/a'))
        )
        link.click()
        print("Enlace clickeado")

        # Volver a llenar las credenciales e intentar iniciar sesión de nuevo
        WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, 'username'))
        )

        driver.find_element(By.ID, 'username').send_keys(CORREO_CONSTANTE)
        driver.find_element(By.ID, 'password').send_keys('admin')
        print("Credenciales ingresadas nuevamente")

        driver.find_element(
            By.XPATH, '/html/body/jhi-main/div/div/login-form/div/div/div[3]/form/button').click()
        print("Botón de inicio de sesión presionado nuevamente")

    except Exception as e:
        print("No se encontró el enlace de desbloqueo o hubo otro error:", str(e))



    # Esperar a que la página cargue y hacer clic en elementos
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab/div/div/div/div/div[1]'))).click()
    print("Primer elemento encontrado y clic realizado")

    time.sleep(2) 
    # Esperar hasta que el segundo elemento sea visible y hacer clic
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[2]/div/menu/div/div[2]/div/div/div[2]'))).click()
    print("Segundo elemento encontrado y clic realizado")
    time.sleep(2) 
    # Esperar a que la tabla se cargue y buscar ENTRADA
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/bandeja/div/div/p-datatable')))
    search_input = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/bandeja/div/div/p-datatable/div/div[2]/table/thead/tr/th[3]/input')))
    search_input.send_keys(ENTRADA_CONSTANTE)
    print(f"Texto '{ENTRADA_CONSTANTE}' ingresado en el campo de búsqueda")

    time.sleep(2) 
    # Esperar que se actualicen los resultados y hacer clic en el botón
    action_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/bandeja/div/div/p-datatable/div/div[2]/table/tbody/tr/td[8]/span/button')))
    driver.execute_script("arguments[0].click();", action_button)

    # Procesar los datos filtrados
    start_processing = False
    for index, row in filtered_df.iterrows():
        CARGA = row['CARGA']
        print("carga ingresada ",CARGA)
        print("*********************************")
        print("*********************************")
        print("**    DEBE DE INICALIZAR EN : ", CARGA+1, "    **")
        print("*********************************")
        print("*********************************")
        if CARGA == RANGO_CARGA_INICIAL:
            start_processing = True
          # Detener el procesamiento si CARGA alcanza RANGO_CARGA_FINALIZAR
        if CARGA > RANGO_CARGA_FINALIZAR:
            print(f"Proceso finalizado al alcanzar CARGA: {CARGA}")
            boton_cerrar_sesion = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/jhi-navbar/div/div/div/div[3]/div/div[2]/a/span'))
                    )
            boton_cerrar_sesion.click()
            break
        if start_processing:
            CORREO_CONSTANTE = str(row['FOLIO REAL.1'])
            YEAR = str(row['AÑO'])
            TOMO = str(row['TOMO'])
            LIBRO = str(row['LIBRO'])
            VOLUMEN = str(row['VOLUMEN'])
            INSCRIPCION = str(row['INSCRIPCION'])
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/cyvf/div/div[2]/div/div[1]/button[7]')))
                
                # Usar JavaScript para hacer clic en el botón
            new_button = driver.find_element(By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/cyvf/div/div[2]/div/div[1]/button[7]')
            driver.execute_script("arguments[0].click();", new_button)

                # ÇÇÇÇÇÇÇÇÇÇÇÇÇ AQUI ABREB EL MODAL PARA EL A ANTECEDENTE 
                # Esperar a que el modal se vuelva visible (máx 30 segundos)
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente')))

                # Interactuar con el primer dropdown
            WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[1]/select')))
            time.sleep(2) 
            dropdown1 = Select(driver.find_element(By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[1]/select'))
            dropdown1.select_by_index(1)  # Selecciona la segunda opción (índice 1)
            print("Primera opción seleccionada en el modal")

                # Interactuar con el segundo dropdown
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[2]/select')))
            time.sleep(2) 
            dropdown2 = Select(driver.find_element(By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[2]/select'))
            dropdown2.select_by_index(1)  # Selecciona la segunda opción (índice 1)
            print("Segunda opción seleccionada en el modal")

                # Interactuar con el tercer dropdown
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[3]/select')))
            time.sleep(2) 
            dropdown3 = Select(driver.find_element(By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[3]/select'))
            dropdown3.select_by_index(2)  # Selecciona la tercera opción (índice 2)
                
                
                
                
                # &&&---------------------------------------------------------------------------------------------DATA DEL EXCEL 
            def wait_and_send_keys(driver, xpath, value, description):
                try:
                    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, xpath)))
                    input_field = driver.find_element(By.XPATH, xpath)
                    input_field.send_keys(value)
                    print(f"Valor {description} '{value}' ingresado correctamente")
                except Exception as e:
                    print(f"Error al ingresar el valor '{description}': {e}")

                # Ingresar el año
            wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[4]/input', YEAR, "AÑO")

                # Ingresar '666' en el segundo campo de entrada
            wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[1]/input', TOMO, "TOMO")

                # Ingresar '777' en el tercer campo de entrada
            wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[2]/input', LIBRO, "LIBRO")

                # Ingresar '888' en el cuarto campo de entrada
            wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[3]/input', VOLUMEN, "VOLUMEN")

                # Ingresar '999' en el quinto campo de entrada
            wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[4]/input', INSCRIPCION, "INSCRIPCION")

                
                
                
                # Hacer clic en el botón
            try:
                time.sleep(3) 
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[3]/button[1]'))).click()

                try:
                    # Si el modal está presente, hacer clic en el segundo botón
                    modal_present = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                        (By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/p-dialog/div/div[2]')))

                    if modal_present:
                        print("Modal detectado, presionando el segundo botón.")
                        # WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                        #     (By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[3]/button[2]'))).click()  ESTA ES EL CANCELAR DEL MODAL  DONDE DA DE ALTA ANTENCEDEMTE
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                            (By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/p-dialog/div/div[2]/div[2]/button[1]'))).click()
                        print(
                            "Segundo botón presionado CANCELAR EN AGREGAR ANTECEDENTE.")
                except:
                    print("Modal no encontrado o no apareció.")

                print("Botón presionado correctamente.")

                
            except Exception as e:
                    print(f"Error al hacer clic en el botón: {e}")
                  # Cerrar sesión
                    boton_cerrar_sesion = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/jhi-navbar/div/div/div/div[3]/div/div[2]/a/span'))
                    )
                    boton_cerrar_sesion.click()
                    print("Se hizo clic en el botón de cerrar sesión")
            time.sleep(3) 
            print(f"Correo: {CORREO_CONSTANTE}, Año: {YEAR}, Tomo: {TOMO}, Libro: {LIBRO}, Volumen: {VOLUMEN}, Inscripción: {INSCRIPCION}")

except NoSuchElementException as e:
    print(f"Error: Elemento no encontrado - {e}")
    boton_cerrar_sesion = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/jhi-navbar/div/div/div/div[3]/div/div[2]/a/span'))
                    )
    boton_cerrar_sesion.click()
except TimeoutException as e:
    print(f"Error: Tiempo de espera excedido - {e}")
finally:
    driver.quit()
