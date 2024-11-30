import requests
from io import BytesIO
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
# from selenium.webdriver.chrome.service import Service


from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException,StaleElementReferenceException

from webdriver_manager.firefox import GeckoDriverManager
# from webdriver_manager.chrome import ChromeDriverManager

import time




ENTRADA_CONSTANTE = '52927'  # 52962	vreyes_scontino@ifreh.gob.mx

# C:\Users\rigoberto diaz\OneDrive\Documentos\SCRIP_\CRIPS__\firmas\MOPJ7808311Y2.cer
RUTA_CERTIFICADO_CER = r'C:\Users\rigoberto diaz\OneDrive\Documentos\SCRIP_\CRIPS__\firmas\MOPJ7808311Y2.cer'
RUTA_CERTIFICADO_KEY = r'C:\Users\rigoberto diaz\OneDrive\Documentos\SCRIP_\CRIPS__\firmas\MOPJ7808311Y2.key'

CORREO_CONSTANTE = ''  # Initialize CORREO_CONSTANTE
YEAR = ''
TOMO = ''
LIBRO = ''
VOLUMEN = ''
INSCRIPCION = ''
RANGO_CARGA_INICIAL = 48 #- 3269-3283
RANGO_CARGA_FINALIZAR =65
MAX_INTENTOS = 4

# Función para realizar acciones posteriores al inicio de sesión
def realizar_acciones(driver):
    try:

        # Localizar la tabla y las filas dentro de ella
        tabla = driver.find_element(By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/cyvf/div/div[2]/div/div[2]/antecedente-prelacion/div[1]/table')
        filas = tabla.find_elements(By.XPATH, './tbody/tr')
        cantidad_filas = len(filas)
        print(f"Cantidad de filas en la tabla: {cantidad_filas}")

        # Iterar sobre las filas para encontrar el checkbox y el botón 'Procesar'
        for index in range(1, cantidad_filas + 1):
            intentos = 0
            while intentos < MAX_INTENTOS:
                try:
                    print(f"Procesando fila {index}...")
                    tabla = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/cyvf/div/div[2]/div/div[2]/antecedente-prelacion/div[1]/table'))
        )
                    # Reobtener las filas para evitar referencias obsoletas
                    filas = tabla.find_elements(By.XPATH, './tbody/tr')
                    fila = filas[index - 1]  # Usar index-1 porque enumerate() empieza en 1

                    # Buscar el checkbox en la fila actual
                    checkbox = fila.find_element(By.XPATH, './td[9]/input')
                # Si el checkbox NO está seleccionado
                    if not checkbox.is_selected():
                        print(f"Checkbox en la fila {index} no está seleccionado, se hará clic en 'Procesar'.")

                        # Localizamos el botón 'Procesar' en la fila actual
                        boton_procesar = fila.find_element(By.XPATH, './td[10]/button[1]')
                        driver.execute_script("arguments[0].scrollIntoView();", boton_procesar)

                        # Esperamos 5 segundos antes de hacer clic en el botón
                        time.sleep(3)
                        boton_procesar.click()
                        print(f"Botón 'Procesar' en la fila {index} clicado con éxito.")

                        # Seleccionar opciones de un select en la siguiente página
                        select_opcion1 = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[2]/tbody/tr[2]/td[1]/select/option[1]'))
                        )
                        select_opcion1.click()
                        print("Opción 1 seleccionada")
                        time.sleep(3)

                        select_opcion29 = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[3]/tbody/tr[2]/td[1]/select/option[29]'))
                        )
                        select_opcion29.click()
                        print("Opción 29 seleccionada")
                        # time.sleep(3)

                        # Ingresar 'NA' en el primer campo de texto
                        input_na1 = WebDriverWait(driver, 30).until(
                            EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[3]/tbody/tr[2]/td[2]/input'))
                        )
                        input_na1.send_keys('NA')
                        print("Campo de texto 1 con valor 'NA'")
                        time.sleep(3)

                        # Seleccionar la opción 42 en el tercer select
                        select_opcion42 = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[4]/tbody/tr[2]/td[1]/select/option[42]'))
                        )
                        select_opcion42.click()
                        print("Opción 42 seleccionada")
                        # time.sleep(5)

                        # Ingresar 'NA' en el segundo campo de texto
                        input_na2 = WebDriverWait(driver, 30).until(
                            EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[4]/tbody/tr[2]/td[2]/input'))
                        )
                        input_na2.send_keys('NA')
                        print("Campo de texto 2 con valor 'NA'")
                        # time.sleep(5)

                        #* //$$$!!!!///////////*********************   ···$$$$$$$$$$$$$$$$$$$
                

                        # Ingresar valor 'NA' en el tercer campo de texto
                        input_na3 = WebDriverWait(driver, 30).until(
                            EC.visibility_of_element_located(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[5]/tbody/tr[3]/td[1]/input'))
                        )
                        input_na3.send_keys('NA')
                        print("Campo de texto 3 con valor 'NA'")
                        time.sleep(3)

                        # Seleccionar la opción 48 del cuarto select
                        select_opcion48 = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[6]/tbody/tr[2]/td[1]/select/option[48]'))
                        )
                        select_opcion48.click()
                        print("Opción 48 seleccionada")
                        # time.sleep(5)

                        # Ingresar valor '0' en el cuarto campo de texto
                        input_0 = WebDriverWait(driver, 30).until(
                            EC.visibility_of_element_located(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[6]/tbody/tr[2]/td[3]/input'))
                        )
                        input_0.send_keys('0')
                        print("Campo de texto 4 con valor '0'")
                        time.sleep(3)

                        # Seleccionar la opción 1 del quinto select
                        select_opcion1_final = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[6]/tbody/tr[2]/td[4]/select/option[1]'))
                        )
                        select_opcion1_final.click()
                        print("Opción 1 del select final seleccionada")
                        time.sleep(3)

                        # Seleccionar la opción 10 del último select
                        select_opcion10 = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/div[1]/table[6]/tbody/tr[2]/td[6]/select/option[10]'))
                        )
                        select_opcion10.click()
                        print("Opción 10 del select final seleccionada")
                        # time.sleep(5)

                        # Clic PARA AGREGAR PERSONA
                        # Clic para agregar persona
                        boton_agregar_persona = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/titulares-iniciales-content/div/div/form/div/div[2]/div/table/tfoot/tr/td/button'))
                        )
                        boton_agregar_persona.click()
                        print("Se hizo clic en el botón para agregar persona")

                        # Espera para permitir que el formulario cargue
                        # time.sleep(5)

                        # Llenar el campo th[3] con 'NA'
                        campo_th3 = driver.find_element(
                            By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/titulares-iniciales-content/div/div/form/div/div[2]/div/table/tbody/tr/th[3]/input')
                        campo_th3.send_keys('NA')
                        print("Campo th[3] llenado con NA")

                        # Llenar el campo th[4] con 'NA'
                        campo_th4 = driver.find_element(
                            By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/titulares-iniciales-content/div/div/form/div/div[2]/div/table/tbody/tr/th[4]/input')
                        campo_th4.send_keys('NA')
                        print("Campo th[4] llenado con NA")

                        # Llenar el campo th[5] con 'NA'
                        campo_th5 = driver.find_element(
                            By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/titulares-iniciales-content/div/div/form/div/div[2]/div/table/tbody/tr/th[5]/input')
                        campo_th5.send_keys('NA')
                        print("Campo th[5] llenado con NA")

                        # Llenar el campo th[6] con '100'
                        campo_th6 = driver.find_element(
                            By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/titulares-iniciales-content/div/div/form/div/div[2]/div/table/tbody/tr/th[6]/input')
                        campo_th6.send_keys('100')
                        print("Campo th[6] llenado con 100")

                        # Llenar el campo th[7] con '100'
                        campo_th7 = driver.find_element(
                            By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/table/tbody/tr/td/predio-detalle-form-content/div/div/form/titulares-iniciales-content/div/div/form/div/div[2]/div/table/tbody/tr/th[7]/input')
                        campo_th7.send_keys('100')
                        print("Campo th[7] llenado con 100")
                        time.sleep(3)
                    # Dar clic en el botón final de enviar formulario
                        boton_final = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/folio-form/div/div[2]/div/div/div[2]/button'))
                        )
                        boton_final.click()
                        print("Se hizo clic en el botón final")
                        # Espera después de hacer clic para que se procese la acción
                        time.sleep(3)
                        # Esperar un poco antes de cerrar sesión

                        #!! TODO   TODO VISTA
                        #!! TODO   TODO VISTA
                        # 2. Seleccionar la primera opción del primer select
                        primera_opcion_select = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[2]/div[1]/div/div[2]/div/div[1]/select/option[1]'))
                        )
                        primera_opcion_select.click()

                        # 3. Seleccionar la quinta opción del segundo select
                        segunda_opcion_select = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[2]/div[1]/div/div[2]/div/div[2]/select/option[5]'))
                        )
                        segunda_opcion_select.click()

                        # 4. Seleccionar la quinta opción del tercer select
                        tercera_opcion_select = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[2]/div[1]/div/div[2]/div/div[3]/select/option[5]'))
                        )
                        tercera_opcion_select.click()

                        # 5. Clic en el botón dentro del formulario
                        boton_formulario = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[2]/div[1]/div/div[3]/button'))
                        )
                        boton_formulario.click()
                #!! TODO   TODO BIEN HASTA AQUI  A DAR CLICK EN ICPNO DE EDITAR
                        time.sleep(3)
                        # 6. Clic en un ícono dentro de una tabla
                        boton_icono = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[2]/div[2]/div/div[2]/table/tbody/tr/td[5]/div/a[1]/i'))
                        )
                        boton_icono.click()
            #* //$$$!!!!///////////*********************  LLENAR OTROS CAMPOS ···$$$$$$$$$$$$$$$$$$$
            # 7. Interacción con los campos de input para ingresar valores
                        input1 = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[3]/acto/div/form/div[2]/div[4]/modulo/div/div[2]/div/table/tbody/tr/th[2]/campo/div/div/div/div/input'))
                        )
                        input1.send_keys('NA')

    # !$$$ !!!!/////// INPUT NUMERICOS
                        input2N = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[3]/acto/div/form/div[2]/div[4]/modulo/div/div[2]/div/table/tbody/tr/th[12]/campo/div/div/div/div/div[1]/input'))
                        )
                        input2N.send_keys('0')
                        input3N = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[3]/acto/div/form/div[2]/div[4]/modulo/div/div[2]/div/table/tbody/tr/th[13]/campo/div/div/div/div/div[1]/input'))
                        )
                        input3N.send_keys('0')
                        
                        # !$$$ !!!!/////// PARTE DE MONTO OPERACION
                        input3 = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[3]/acto/div/form/div[2]/div[8]/modulo/div/div[2]/div/table/tbody/tr/th[1]/campo/div/div/div/div/input'))
                        )
                        input3.send_keys('0')
                        time.sleep(3)
                            # 8. Seleccionar la opción 13 en el siguiente dropdown
                        dropdown_select = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[3]/acto/div/form/div[2]/div[8]/modulo/div/div[2]/div/table/tbody/tr/th[2]/campo/div/div/div/div/div[1]/lista/div/select/option[13]'))
                        )
                        dropdown_select.click()

            # 9. Clic en el botón de guardar al final
                        boton_guardar = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '//*[@id="boton-guardar"]'))
                        )
                        boton_guardar.click()
                        
                        time.sleep(3)
                        
                        # Esperar a que aparezca la ventana de confirmación y aceptarla
                        alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
                        alert = driver.switch_to.alert
                        alert.accept()  # Confirmar la ventana emergente          
                        # !$$$ !!!!/////// ICONO DE DISCO GUARDAR 
                        
                        # Espera un poco después de guardar
                        # time.sleep(5)

                        # !$$$ !!!!/////// Clic en el icono de disco guardar
                        icono_disco_guardar = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[2]/div[2]/div/div[2]/table/tbody/tr/td[5]/div/a[2]/i')
                            )
                        )
                        icono_disco_guardar.click()
                        time.sleep(3)

                        # Espera a que el modal aparezca
                        modal_confirmacion = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/p-confirmdialog/div/div[3]/p-footer/button[2]')
                            )
                        )
                        modal_confirmacion.click()
                        time.sleep(3)
                        # !$$$ !!!!/////// Clic en el SOBRESITO 
                        # Clic en el sobresito a firmar
                        sobresito_firmar = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/jhi-ingreso-acto-predio/div[2]/div[2]/div/div[2]/table/tbody/tr/td[5]/div/a[5]/i')
                            )
                        )
                        sobresito_firmar.click()
                        # !$$$ !!!!/////// INSERTA ARCHIVOS 
                        
                        # Subir el archivo .cer
                        certificado_input = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="certificate_file"]'))
                        )
                         #  certificado_input.send_keys(r'C:\MyScrips\firmas\MOPJ7808311Y2.cer')
                         #  certificado_input.send_keys(r'C:\MyScrips\firmas\MOPJ7808311Y2.cer')
                        certificado_input.send_keys(RUTA_CERTIFICADO_CER)

                        # Subir el archivo .key
                        llave_input = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="privkey_file"]'))
                        )
                          # llave_input.send_keys(r'C:\MyScrips\firmas\MOPJ7808311.key')
                        llave_input.send_keys(RUTA_CERTIFICADO_KEY)

                        # Ingresar la contraseña
                        password_input = WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="password"]'))
                        )
                        password_input.send_keys('@SContino79')  # Ingresar la contraseña
                        time.sleep(6)
                        # Clic en el botón para firmar
                        boton_firmar = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/ngb-modal-window/div/div/firma-content/div[3]/button[1]')
                            )
                        )
                        time.sleep(5)
                        # !$$$ !!!!/////// FIRMANDO -...........
                        boton_firmar.click()

                        # Esperar a que el siguiente botón esté disponible y clic
                        nuevo_boton_FINALIZAR = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/captura-actos/div/div[2]/div/div/div/div/button')
                            )
                        )
                        time.sleep(3)  # Agregar un pequeño retraso adicional si es necesario
                        nuevo_boton_FINALIZAR.click()
                        time.sleep(5)  # Agregar un pequeño retraso adicional si es necesario


                    else:
                        print(f"Checkbox en la fila {index} ya está seleccionado, saltando esta fila.")
                    break

                except NoSuchElementException:
                    print(f"No se encontró el checkbox o botón en la fila {index}, pasando a la siguiente fila.")
                    break  # Salir del while e ir a la siguiente fila

                except StaleElementReferenceException:
                      # Reobtener las filas para evitar referencias obsoletas
                    filas = tabla.find_elements(By.XPATH, './tbody/tr')
                    fila = filas[index - 1]  # Usar index-1 porque enumerate() empieza en 1

                    intentos += 1
                    print(f"Elemento en la fila {index} no es válido, reintentando ({intentos}/{MAX_INTENTOS})...")
                    time.sleep(3)  # Pequeño retraso antes de reintentar

            if intentos == MAX_INTENTOS:
                print(f"No se pudo procesar la fila {index} después de {MAX_INTENTOS} intentos, pasando a la siguiente fila.")

        print("Proceso completado.")
    

    except TimeoutException as e:
        print(f"Error: {e}")
# Paso 1: Descargar el archivo de Google Sheets en formato Excel

def download_file():
    # sheet_url = "https://docs.google.com/spreadsheets/d/1m2wOgUVjvaYiCRomQxD1WBEElKCHC8zu/export?format=xlsx"
        # URL del archivo de Google Sheets exportado como XLSX
    sheet_url = "https://docs.google.com/spreadsheets/d/1MImX4BOLJ0qOvJWOopi8tE2Q4KcIsI2_O9cDFyfG8tg/export?format=xlsx"

    response = requests.get(sheet_url)

    if response.status_code == 200:
        excel_data = response.content
        # Leer el archivo Excel usando Pandas
        # df = pd.read_excel(BytesIO(excel_data), sheet_name='Activos', engine='openpyxl')
        df = pd.read_excel(BytesIO(excel_data), sheet_name=ENTRADA_CONSTANTE, engine='openpyxl')

        print(df)
        # df = pd.read_excel(BytesIO(excel_data), sheet_name=ENTRADA_CONSTANTE, engine='openpyxl')
        print("Archivo Excel cargado con éxito")
        print(f"Filtrando por 'ENTRADA' == {ENTRADA_CONSTANTE} y 'cantidad_duplicados' == 0")
        # filtered_df = df[(df['ENTRADA'] == ENTRADA_CONSTANTE) & (df['cantidad_duplicados'] == 0)]
        filtered_df = df[df['cantidad_duplicados'] == 0]
        print(df['ENTRADA'].head())
        print(df[df['ENTRADA'] == ENTRADA_CONSTANTE])  # Asegúrate de que 52933 esté entre comillas si es una cadena
        print(df[df['ENTRADA'].isna()])

        # Verificar si se encontraron filas
        print(f"Filas filtradas: {len(filtered_df)}")

        # Asegurarnos de que la columna 'ENTRADA' sea de tipo string
        df['ENTRADA'] = df['ENTRADA'].astype(str).str.strip()
        

        # Filtrar las filas donde la columna 'ENTRADA' sea igual a '52867'
        # filtered_df = df[df['ENTRADA'] == ENTRADA_CONSTANTE]
        # filtered_df = df[(df['ENTRADA'] == ENTRADA_CONSTANTE) & (df['cantidad_duplicados'] == 0)]
        filtered_df = df[df['cantidad_duplicados'] == 0]

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
        # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

        driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
        driver.get("https://ifreh-s.hidalgo.gob.mx:8443/erpp/#/login")
        print("Página abierta con éxito")

        # Esperar hasta que el campo de Username sea visible (máx 30 segundos)
        WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.ID, 'username')))

        # Ingresar correo y contraseña
        driver.find_element(By.ID, 'username').send_keys(CORREO_CONSTANTE)
        driver.find_element(By.ID, 'password').send_keys('admin')

        # Iniciar sesión
        driver.find_element(By.XPATH, '/html/body/jhi-main/div/div/login-form/div/div/div[3]/form/button').click()
        print("Credenciales ingresadas y botón de inicio de sesión presionado")
        time.sleep(3) 
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
                EC.element_to_be_clickable(
                    (By.XPATH, '/html/body/jhi-main/div/div/jhi-user-unlock/div/div/div[2]/form/button'))
            )
            unlock_button.click()
            print("Botón de desbloqueo presionado")

            # Esperar hasta que el enlace esté visible y luego hacer clic
            link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '/html/body/jhi-main/div/div/jhi-user-unlock/div/div/div/div/a'))
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

        time.sleep(3) 
        # Esperar hasta que el segundo elemento sea visible y hacer clic
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[2]/div/menu/div/div[2]/div/div/div[2]'))).click()
        print("Segundo elemento encontrado y clic realizado")
        time.sleep(3) 
        # Esperar a que la tabla se cargue y buscar ENTRADA
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/bandeja/div/div/p-datatable')))
        search_input = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/bandeja/div/div/p-datatable/div/div[2]/table/thead/tr/th[3]/input')))
        search_input.send_keys(ENTRADA_CONSTANTE)
        print(f"Texto '{ENTRADA_CONSTANTE}' ingresado en el campo de búsqueda")

        time.sleep(3) 
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
                # YEAR = str(row['AÑO']).astype(int)
                # TOMO = str(row['TOMO']).astype(int)
                # LIBRO = str(row['LIBRO']).astype(int)
                # VOLUMEN = str(row['VOLUMEN']).astype(int)
                # INSCRIPCION = str(row['INSCRIPCION']).astype(int)
                YEAR = str(int(row['AÑO']))
                TOMO = row['TOMO']
                LIBRO = str(int(row['LIBRO']))
                VOLUMEN = str(int(row['VOLUMEN']))
                # Procesar INSCRIPCION con validación de tipo
                INSCRIPCION = row['INSCRIPCION']

                if isinstance(TOMO, float) and TOMO.is_integer():
                    TOMO = str(int(TOMO))  # Convertir a entero si es float con valor entero
                else:
                    TOMO = str(TOMO).strip()  # Mantener tal cual si es cadena o no cumple lo anterior

                if isinstance(INSCRIPCION, float) and INSCRIPCION.is_integer():
                    INSCRIPCION = str(int(INSCRIPCION))  # Convertir a entero si es float con valor entero
                else:
                    INSCRIPCION = str(INSCRIPCION).strip()  # Mantener tal cual si es cadena o no cumple lo anterior

                print(f"INSCRIPCION procesado: {INSCRIPCION}")

                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/cyvf/div/div[2]/div/div[1]/button[7]')))
                    
                    # Usar JavaScript para hacer clic en el botón
                new_button = driver.find_element(By.XPATH, '/html/body/jhi-main/div/div/jhi-home/erpp-tabs/erpp-tab[3]/div/cyvf/div/div[2]/div/div[1]/button[7]')
                driver.execute_script("arguments[0].click();", new_button)

                    # ÇÇÇÇÇÇÇÇÇÇÇÇÇ AQUI ABREB EL MODAL PARA EL A ANTECEDENTE 
                    # Esperar a que el modal se vuelva visible (máx 30 segundos)
                WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente')))

                    # Interactuar con el primer dropdown
                WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[1]/select')))
                time.sleep(3) 
                dropdown1 = Select(driver.find_element(By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[1]/select'))
                dropdown1.select_by_index(1)  # Selecciona la segunda opción (índice 1)
                print("Primera opción seleccionada en el modal")

                    # Interactuar con el segundo dropdown
                WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[2]/select')))
                time.sleep(3) 
                dropdown2 = Select(driver.find_element(By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[2]/select'))
                dropdown2.select_by_index(1)  # Selecciona la segunda opción (índice 1)
                print("Segunda opción seleccionada en el modal")

                    # Interactuar con el tercer dropdown
                WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[1]/div[3]/select')))
                time.sleep(6) 
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

                # Asegúrate de que TOMO no sea '01A' antes de ingresar el valor
                if TOMO != '01A':
                    # Ingresar 'TOMO' en el segundo campo de entrada
                    wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[1]/input', TOMO, "TOMO")
                else:
                    TOMO='1A'
                    wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[1]/input', TOMO, "TOMO")
                    print(f"El valor de TOMO ({TOMO}) no es válido. No puede ser '01A'.")


                    # Ingresar '777' en el tercer campo de entrada
                wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[2]/input', LIBRO, "LIBRO")

                    # Ingresar '888' en el cuarto campo de entrada
                wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[3]/input', VOLUMEN, "VOLUMEN")

                    # Ingresar '999' en el quinto campo de entrada
                wait_and_send_keys(driver, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[2]/div/form/div[2]/div[4]/input', INSCRIPCION, "INSCRIPCION")

                    
                    
                    
                    # Hacer clic en el botón
                try:
                    time.sleep(5) 
                    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/ngb-modal-window/div/div/jhi-cyvf-antecedente/div[3]/button[1]'))).click()

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
                        
                time.sleep(5) 
                print(f"Correo: {CORREO_CONSTANTE}, Año: {YEAR}, Tomo: {TOMO}, Libro: {LIBRO}, Volumen: {VOLUMEN}, Inscripción: {INSCRIPCION}")
                realizar_acciones(driver)

    except NoSuchElementException as e:
        print(f"Error: Elemento no encontrado - {e}")
        boton_cerrar_sesion = WebDriverWait(driver, 30).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/jhi-main/jhi-navbar/div/div/div/div[3]/div/div[2]/a/span'))
                        )
        boton_cerrar_sesion.click()
    except TimeoutException as e:
        print(f"Error: Tiempo de espera excedido - {e}")
    # finally:
    #     driver.quit()


download_file()