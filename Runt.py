import logging
import os
import time
import cv2
import numpy as np
import pytesseract
from datetime import datetime
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from typing import Optional
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import gspread
from google.oauth2.service_account import Credentials
import io

# --- Configuración ---
BASE_PATH = "C:\\Users\\svega\\Downloads\\Motos_Runt"
pytesseract.pytesseract.tesseract_cmd = 'C:\\Users\\svega\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe'
CAPTCHA_FOLDER = os.path.join(BASE_PATH, "captchas")
os.makedirs(CAPTCHA_FOLDER, exist_ok=True)

CAPTCHA_TO_TRAIN_FOLDER = os.path.join(BASE_PATH, "captchas")
os.makedirs(CAPTCHA_TO_TRAIN_FOLDER, exist_ok=True)

# --- Configuración de logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("automatizacion.log", encoding='utf-8')
    ]
)

def obtener_datos_unicos():
    # Configuración de credenciales
    SCOPES = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file("C:\\Users\\svega\\Downloads\\Motos_Runt\\sheets.json", scopes=SCOPES)
    client = gspread.authorize(creds)

    # Abre el archivo y la hoja
    sheet = client.open_by_key("1EcCiF1nRMfqKgHv7z4wPxy-QRvW1z2-sLQXlZSvgDgw")
    worksheet = sheet.worksheet("Reporte Comparendos")
    # Lee columnas A (cédula) y E (placa)
    cedulas = worksheet.col_values(1)  # Columna A
    placas = worksheet.col_values(5)   # Columna E

    # Saltamos la fila de encabezados y creamos lista de tuplas
    datos = [(str(cedulas[i]).strip(), str(placas[i]).strip()) for i in range(1, min(len(cedulas), len(placas)))]

    # Filtramos duplicados (cédula + placa) y placas vacías
    vistos = set()
    datos_unicos = []
    for cedula, placa in datos:
        if placa and (cedula, placa) not in vistos:
            vistos.add((cedula, placa))
            datos_unicos.append((cedula, placa))

    return datos_unicos

def guardar_en_sheets(resultados):
    try:
        SCOPES = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(
            "C:\\Users\\svega\\Downloads\\Motos_Runt\\sheets.json",
            scopes=SCOPES
        )
        client = gspread.authorize(creds)

        # Abrimos el archivo y la hoja
        sheet = client.open_by_key("1n-JCFHPjzh9VzUzi1fWlP-OVGCu_pKuQE4Mh4wMRoDA")
        worksheet = sheet.worksheet("Datos Runt")

        filas = []
        for r in resultados:
            # Si no hay datos, llenamos con "No datos"
            soat_data = r["datos_soat"] if isinstance(r["datos_soat"], list) else ["No datos"] * 7
            rtm_data = r["datos_técnicos"] if isinstance(r["datos_técnicos"], list) else ["No datos"] * 8

            # Quitamos la última columna de datos_técnicos si existe
            if len(rtm_data) > 0:
                rtm_data = rtm_data[:-1]

            # Construimos la fila final
            fila = [r["Tiempo ejecucion"], r["cedula"], r["placa"]] + soat_data + rtm_data
            filas.append(fila)

        # Guardamos en la hoja
        worksheet.append_rows(filas, value_input_option="RAW") # type: ignore
        logging.info(f"Se guardaron {len(filas)} filas en Google Sheets.")

    except Exception as e:
        logging.error(f"Error al guardar en Google Sheets: {e}")

# --- Driver ---
def iniciar_driver(max_attempts=3):
    for attempt in range(max_attempts):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--ignore-certificate-errors")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-blink-features=AutomationControlled")
            service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            return driver
        except WebDriverException as e:
            logging.error(f"Intento {attempt + 1} fallido al iniciar driver: {e}")
            time.sleep(2)
    logging.error("No se pudo iniciar el driver tras varios intentos.")
    return None

def cerrar_driver(driver: webdriver.Chrome) -> None:
    try:
        if driver:
            driver.quit()
            logging.info("Driver cerrado correctamente.")
    except WebDriverException as e:
        logging.error("Error al cerrar el driver: %s", e)

# --- Capturar Captcha ---
def capturar_captcha(driver, carpeta_temp=CAPTCHA_FOLDER):
    try:
        captcha_div = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/host-runt-root/app-layout/app-theme-runt2/mat-sidenav-container/mat-sidenav-content/div/ng-component/div/div/form/div[2]/div[2]/mat-card/mat-card-content/div[2]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", captcha_div)

        time.sleep(1)

        captcha_img = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "body > host-runt-root > app-layout > app-theme-runt2 > mat-sidenav-container > mat-sidenav-content > div > ng-component > div > div > form > div:nth-child(2) > div.col-sm-9 > mat-card > mat-card-content > div.divCaptcha.ng-star-inserted > img"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", captcha_img)

        captcha_div_png = captcha_div.screenshot_as_png
        div_img = Image.open(io.BytesIO(captcha_div_png))

        location = captcha_img.location
        size = captcha_img.size
        logging.info(f"Coordenadas de imgCaptcha: x={location['x']}, y={location['y']}, w={size['width']}, h={size['height']}")
        logging.info(f"Coordenadas de captcha_div: x={captcha_div.location['x']}, y={captcha_div.location['y']}")

        x = int(location['x']) - captcha_div.location['x'] + 0
        y = int(location['y']) - captcha_div.location['y'] + 0
        w = int(size['width']) - 0
        h = int(size['height']) - 0

        x = max(x, 0)
        y = max(y, 0)
        w = min(w, div_img.width - x)
        h = min(h, div_img.height - y)

        captcha_cropped = div_img.crop((x, y, x + w, y + h))

        img_pil = captcha_cropped.convert('L')
        pix = np.array(img_pil)

        # Ajuste de preprocesamiento para mejorar la lectura
        # Aplicar umbral adaptativo para manejar mejor el ruido
        pix = cv2.adaptiveThreshold(pix, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 9, 2)
        pix = cv2.medianBlur(pix, 7)  # Suavizado para reducir ruido

        # Dilatación para conectar caracteres
        kernel = np.ones((1, 1), np.uint8)
        pix = cv2.dilate(pix, kernel, iterations=1)

        img_pil = Image.fromarray(pix)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        temp_filename = os.path.join(carpeta_temp, f"captcha_temp_.png")
        
        img_pil.save(temp_filename)

        logging.info(f"Captcha inicial capturado y guardado en: {temp_filename}")
        return temp_filename, img_pil

    except Exception as e:
        logging.error(f"Error al capturar o procesar el captcha del navegador: {e}")
        return None, None

def resolver_captcha(path: str, img_pil: Image.Image) -> Optional[str]:
    logging.info(f"Intentando resolver captcha con Tesseract para: {os.path.basename(path)}")
    custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789' # En caso de caracteres adicionarlos 
    try:
        text_tesseract = pytesseract.image_to_string(img_pil, config=custom_config).strip()

        if text_tesseract and all(c.isalnum() for c in text_tesseract):
            logging.info(f"Captcha resuelto por Tesseract: '{text_tesseract}'")
            return text_tesseract
        else:
            logging.warning(f"Tesseract no pudo resolver el captcha de forma fiable: '{text_tesseract}'. Guardando para entrenamiento manual.")
            manual_train_filename = os.path.join(CAPTCHA_TO_TRAIN_FOLDER, os.path.basename(path))
            img_pil.save(manual_train_filename)
            logging.info(f"Captcha guardado para entrenamiento manual: {manual_train_filename}")
            return None

    except Exception as e:
        logging.error(f"Error al resolver captcha con Tesseract: {e}")
        return None

# --- Función para aceptar alerta de CAPTCHA inválido ---
def aceptar_alerta(driver):
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
        logging.info("Alerta de CAPTCHA inválido aceptada.")
        return True
    except:
        try:
            # Buscar el popup SweetAlert2
            aceptar_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".swal2-confirm"))
            )
            aceptar_btn.click()
            logging.info("Botón 'Aceptar' en popup de CAPTCHA inválido clicado.")
            return True
        except:
            logging.warning("No se encontró alerta ni popup de CAPTCHA inválido.")
            return False

def verificar_mensaje_error(driver, cedula, placa):
    try:
        # Espera a que aparezca el contenedor del mensaje de error
        mensaje_error = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".swal2-html-container"))
        )
        texto_error = mensaje_error.text.strip()

        if "Los datos registrados no corresponden" in texto_error:
            logging.info(f"Los datos registrados no corresponden con los propietarios activos para el vehículo consultado.: {cedula}, Placa: {placa}")
            
            # Clic en el botón Aceptar para cerrar el popup
            aceptar_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".swal2-confirm"))
            )
            aceptar_btn.click()
            return True  # Hubo error y ya se manejó
    except:
        return False  # No apareció el mensaje
    return False

def extraer_datos_tabla(driver, titulo_panel):
    """
    Expande el panel con el título dado y extrae todos los datos.
    Si no hay datos, devuelve un mensaje indicando que no se encontraron registros.
    """
    try:
        # 1. Localizamos el panel por el título
        header = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//mat-expansion-panel-header[.//mat-panel-title[contains(text(), '{titulo_panel}')]]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", header)

        # 2. Si está cerrado, lo abrimos
        if header.get_attribute("aria-expanded") == "false":
            driver.execute_script("arguments[0].click();", header)

        time.sleep(3.5)

        # 3. Esperamos el contenido del panel
        panel_contenido = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, f"//mat-expansion-panel[.//mat-panel-title[contains(text(), '{titulo_panel}')]]"))
        )

        # 4. Verificamos si hay mensaje de "No se encontró información"
        try:
            mensaje_vacio = panel_contenido.find_element(By.XPATH, ".//*[contains(text(), 'No se encontró información')]")
            if mensaje_vacio:
                return f"No se encontraron datos para {titulo_panel}"
        except:
            pass  # Si no hay mensaje, continuamos normalmente

        # 5. Buscamos todas las filas dentro del panel
        filas = panel_contenido.find_elements(By.XPATH, ".//mat-row")
        datos = []

        for fila in filas:
            celdas = fila.find_elements(By.XPATH, ".//mat-cell")
            fila_datos = [celda.text.strip() for celda in celdas]  # Solo valores, sin "Columna X"     primer_valor_t = list(tecnico_mecanica[0].values())[0]

            if fila_datos:
                datos.append(fila_datos)

        # 6. Si no hay filas, también devolvemos mensaje vacío
        if not datos:
            return f"No se encontraron datos para {titulo_panel}"

        return datos

    except Exception as e:
        return f"⚠️ Error extrayendo datos del panel '{titulo_panel}': {e}"

def limpiar_estado(estado):
    """
    Limpia el texto del estado del SOAT.
    Reemplaza 'check_circle' por '✅ VIGENTE' y 'cancel' por '❌ NO VIGENTE'.
    Si no coincide, devuelve el texto original sin saltos de línea.
    """
    estado = estado.replace("\n", "")  # Quitamos saltos de línea
    if "check_circle" in estado:
        return "✅ VIGENTE"
    elif "cancel" in estado:
        return "❌ NO VIGENTE"
    return estado  # Si no coincide, lo dejamos igual

# --- Función para procesar una consulta ---
def procesar_consulta(driver, cedula: str, placa: str):
    max_intentos = 1
    intentos = 0

    while intentos < max_intentos:
        try:
            # Navegar a la página
            driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
            time.sleep(3)
            
            try:
                WebDriverWait(driver, 5).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.swal2-container"))
                )
            except:
                pass  # Si no existe, seguimos normal

            try:
                placa_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="placa"]'))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", placa_input)
                placa_input.click()
                placa_input.send_keys(Keys.CONTROL, "a", Keys.DELETE)
                placa_input.send_keys(placa)
            except TimeoutException:
                logging.warning("No se encontró el campo de placa.")
            
            # Llenar campos
            try:
                # Campo Número de Documento
                cedula_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="documento"]'))
                )
                cedula_input.click()
                cedula_input.send_keys(Keys.CONTROL, "a", Keys.DELETE)
                cedula_input.send_keys(cedula)
            except TimeoutException:
                logging.warning("No se encontró el campo de cédula.")

            # Capturar y resolver captcha
            captcha_path, captcha_img = capturar_captcha(driver)
            if not captcha_path:
                raise Exception("No se pudo capturar el captcha.")

            captcha_text = resolver_captcha(captcha_path, captcha_img)

            if captcha_text is None:
                logging.warning(f"Captcha no resuelto para {cedula}, {placa}. Reiniciando navegador...")
                cerrar_driver(driver)
                
                # Reintentamos la consulta con los mismos datos
                return procesar_consulta(driver, cedula, placa)

            try:
                # Campo Captcha (si existe con formcontrolname)
                captcha_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="captcha"]'))
                )
                captcha_input.click()
                captcha_input.send_keys(Keys.CONTROL, "a", Keys.DELETE)
                captcha_input.send_keys(captcha_text)
            except NoSuchElementException:
                logging.warning("No se encontró el campo de captcha.")

            # Click en botón consultar
            consultar_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/host-runt-root/app-layout/app-theme-runt2/mat-sidenav-container/mat-sidenav-content/div/ng-component/div/div/form/div[2]/div[2]/mat-card/mat-card-content/div[3]/button'))
            )
            consultar_btn.click()

            if verificar_mensaje_error(driver, cedula, placa):
                return None  # Salta a la siguiente consulta

            # Verificar y aceptar alerta/popup de CAPTCHA inválido
            if aceptar_alerta(driver):
                intentos += 0
                time.sleep(1)  # Esperar antes de reintentar
                continue

            time.sleep(1)

            # Esperar la página de resultados
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//label[normalize-space(text())='PLACA DEL VEHÍCULO:']"))
            )

            # # Extraer datos
            # def get_field_value(label_text):
            #     try:
            #         label = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
            #             (By.XPATH, f"//label[normalize-space(text())='{label_text}']")
            #         ))
            #         value = label.find_element(By.XPATH, "./ancestor::div[contains(@class,'row')]//b")
            #         return value.text.strip() if value.text.strip() else "N/A"
            #     except:
            #         return "N/A"

            # licencia = get_field_value("Nro. de licencia de tránsito:")
            # estado = get_field_value("Estado del vehículo:")
            # tipo_servicio = get_field_value("Tipo de servicio:")
            # clase_vehiculo = get_field_value("Clase de vehículo:")
            # marca = get_field_value("Marca:")
            # linea = get_field_value("Línea:")
            # modelo = get_field_value("Modelo:")
            # color = get_field_value("Color:")
            # num_motor = get_field_value("Número de motor:")
            # num_chasis = get_field_value("Número de chasis:")
            # num_vin = get_field_value("Número de VIN:")
            # cilindraje = get_field_value("Cilindraje:")
            # combustible = get_field_value("Tipo Combustible:")
            # fecha_matricula = get_field_value("Fecha de Matricula Inicial(dd/mm/aaaa):")

            # Para el panel de Póliza SOAT YUJ27D
            primer_valor_s = "No hay datos"
            primer_valor_t = "No hay datos"

            datos_soat = extraer_datos_tabla(driver, "Póliza SOAT")
            print("Datos Póliza SOAT:", datos_soat)

            if isinstance(datos_soat, str):
                print(datos_soat)
            elif datos_soat and isinstance(datos_soat[0], list): # diccionario es distinto
                # Si hay datos, tomamos el primer valor
                primer_valor_s = list(datos_soat[0])
                if len(primer_valor_s) > 6:
                    primer_valor_s[6] = limpiar_estado(primer_valor_s[6])
            else:
                primer_valor_s = "No se encontraron datos para Póliza SOAT."

            # if isinstance(tecnico_mecanica, str):
            #     print(tecnico_mecanica)
            # elif tecnico_mecanica and isinstance(tecnico_mecanica[0], dict):

            tecnico_mecanica = extraer_datos_tabla(driver, "Certificado de revisión técnico mecánica y de emisiones contaminantes (RTM)")
            print("Datos Tecnico Mecanica:", tecnico_mecanica)

            if isinstance(tecnico_mecanica, str):
                primer_valor_t = tecnico_mecanica  # mensaje de "No se encontraron datos"
            elif tecnico_mecanica and isinstance(tecnico_mecanica[0], list):
                primer_valor_t = tecnico_mecanica[0]  # toma la primera fila completa
            else:
                primer_valor_t = "No se encontraron datos del Certificado de revisión técnico mecánica y de emisiones contaminantes (RTM)."

            resultado = {
                "Tiempo ejecucion": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "cedula": cedula,
                "placa": placa,
                "datos_soat": primer_valor_s if 'primer_valor_s' in locals() else "No se encontraron datos para Póliza SOAT",
                "datos_técnicos": primer_valor_t if 'primer_valor_t' in locals() else "No se encontraron datos del Certificado de revisión técnico mecánica y de emisiones contaminantes (RTM)"
            }

            return resultado

        except TimeoutException as e:
            logging.error(f"Timeout al procesar consulta para cedula {cedula} y placa {placa}: {e}")
            if aceptar_alerta(driver):
                intentos += 0
                time.sleep(2)
                continue
            return None
        except Exception as e:
            logging.error(f"Error al procesar consulta para cedula {cedula} y placa {placa}: {e}")
            if aceptar_alerta(driver):
                intentos += 0
                time.sleep(2)
                continue
            return None

# --- Main ---
def main():
    driver = iniciar_driver()
    if not driver:
        return
    driver.maximize_window()

    try:
        datos_unicos = obtener_datos_unicos()
        logging.info(f"Se encontraron {len(datos_unicos)} combinaciones únicas para procesar.")

        resultados = []
        for i, (cedula, placa) in enumerate(datos_unicos):
            if not cedula or not placa:
                logging.warning(f"Datos incompletos para cedula {cedula} y placa {placa}. Saltando...")
                continue

            resultado = procesar_consulta(driver, cedula, placa)
            if resultado:
                resultados.append(resultado)
                logging.info(f"Resultado para {cedula}, {placa}: {resultado}")

                # Guardamos cada 5 resultados para ver el avance en tiempo real
                if len(resultados) == 10:
                    guardar_en_sheets(resultados)
                    logging.info("✅ Datos guardados en Sheets.")
                    resultados.clear()  # Limpiamos la lista para el siguiente bloque

                # Clic en "Realizar otra consulta" si no es la última
                if i < len(datos_unicos) - 1:
                    try:
                        # Esperamos a que el overlay desaparezca
                        WebDriverWait(driver, 10).until(
                            EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.backdrop"))
                        )

                        otra_consulta_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-card-actions.btn-return button.mat-raised-button"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", otra_consulta_btn)
                        driver.execute_script("arguments[0].scrollIntoView(true);", otra_consulta_btn)
                        driver.execute_script("arguments[0].click();", otra_consulta_btn)

                        time.sleep(2)
                    except TimeoutException:
                        logging.error("No se encontró el botón 'Realizar otra consulta'.")
            time.sleep(5)

        # Guardamos si queda algún bloque pendiente al final
        if resultados != []:
            guardar_en_sheets(resultados)

        logging.info("Proceso completado. Todos los datos fueron guardados.")

    finally:
        cerrar_driver(driver)

if __name__ == "__main__":
    main()
