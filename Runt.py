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
import json
from pathlib import Path

# ═════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE RUTAS MODERNA
# ═════════════════════════════════════════════════════════════

BASE_PATH = Path(r"C:\Users\cmarroquin\Music\RuntPro")
CAPTCHA_FOLDER = BASE_PATH / "captchas"
CAPTCHA_FOLDER.mkdir(parents=True, exist_ok=True)

CAPTCHA_LEIDOS_FOLDER = BASE_PATH / "captchas_leidos"
CAPTCHA_LEIDOS_FOLDER.mkdir(parents=True, exist_ok=True)

TEMPLATE_FOLDER = BASE_PATH / "templates"
TEMPLATE_FOLDER.mkdir(parents=True, exist_ok=True)

TESSERACT_PATH = r"C:\Users\cmarroquin\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

GOOGLE_CREDS = BASE_PATH / "prueba-de-gmail-486215-345473339c47.json"
ESTADO_FILE = "estado_runt.json"

# ═════════════════════════════════════════════════════════════
# LOGGING
# ═════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("automatizacion.log", encoding='utf-8')
    ]
)

captcha_logger = logging.getLogger('captcha_logger')
captcha_logger.setLevel(logging.INFO)
captcha_logger.propagate = False
file_handler_captcha = logging.FileHandler("captcha_retroalimentacion.log", encoding='utf-8')
formatter_captcha = logging.Formatter('%(message)s')
file_handler_captcha.setFormatter(formatter_captcha)
captcha_logger.addHandler(file_handler_captcha)

# ═══════════════════════════��═════════════════════════════════
# MANEJO DE ESTADO (REANUDACIÓN)
# ═════════════════════════════════════════════════════════════

def estructura_estado_inicial():
    """Crea la estructura inicial del archivo de estado"""
    return {
        "last_execution": None,
        "last_cedula": None,
        "last_placa": None,
        "last_status": None,
        "processed_records": [],
        "current_index": 0,
        "total_records": 0
    }

def cargar_estado():
    """Carga el estado del archivo JSON"""
    if os.path.exists(ESTADO_FILE):
        try:
            with open(ESTADO_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            logging.warning("⚠️ Archivo de estado corrupto, creando nuevo...")
            return estructura_estado_inicial()
    return estructura_estado_inicial()

def guardar_estado(cedula, placa, status, index, total, datos_vehiculo=None, datos_soat=None, datos_tecnica=None):
    """Guarda el estado actual con información detallada"""
    estado_anterior = cargar_estado()
    
    registro_actual = {
        "timestamp": datetime.now().isoformat(),
        "cedula": cedula,
        "placa": placa,
        "status": status,
        "datos_vehiculo": datos_vehiculo or {},
        "datos_soat": datos_soat or [],
        "datos_tecnica": datos_tecnica or []
    }
    
    historial = estado_anterior.get("processed_records", [])
    historial.append(registro_actual)
    
    estado = {
        "last_execution": datetime.now().isoformat(),
        "last_cedula": cedula,
        "last_placa": placa,
        "last_status": status,
        "processed_records": historial,
        "current_index": index,
        "total_records": total
    }
    
    with open(ESTADO_FILE, "w", encoding="utf-8") as f:
        json.dump(estado, f, indent=2, ensure_ascii=False)
    
    logging.info(f"💾 Estado guardado: {placa} - {status} (Total historial: {len(historial)} registros)")

def agregar_registro_procesado(cedula, placa, status, datos_vehiculo=None, datos_soat=None, datos_tecnica=None):
    """Agrega un registro a la lista de procesados"""
    estado = cargar_estado()
    
    registro = {
        "timestamp": datetime.now().isoformat(),
        "cedula": cedula,
        "placa": placa,
        "status": status,
        "datos_vehiculo": datos_vehiculo or {},
        "datos_soat": datos_soat or [],
        "datos_tecnica": datos_tecnica or []
    }
    
    if "processed_records" not in estado:
        estado["processed_records"] = []
    
    estado["processed_records"].append(registro)
    
    with open(ESTADO_FILE, "w", encoding="utf-8") as f:
        json.dump(estado, f, indent=2, ensure_ascii=False)
# ═════════════════════════════════════════════════════════════
# DICCIONARIO DE TEMPLATES (CAPTCHA)
# ═════════════════════════════════════════════════════════════

templates = {}
diccionario_caracteres = {}

if os.path.exists(TEMPLATE_FOLDER):
    for file in os.listdir(TEMPLATE_FOLDER):
        if file.endswith(".png"):
            nombre_sin_extension = file.split(".")[0]
            
            if "_" in nombre_sin_extension:
                char_real = nombre_sin_extension.split("_")[0]
            else:
                char_real = nombre_sin_extension

            path = TEMPLATE_FOLDER / file
            template_img = cv2.imread(str(path), 0)

            if template_img is not None:
                templates[file] = (char_real, template_img)
                
                if char_real not in diccionario_caracteres:
                    diccionario_caracteres[char_real] = []
                diccionario_caracteres[char_real].append(file)

    logging.info(f"✅ Diccionario cargado: {len(templates)} variaciones para {len(diccionario_caracteres)} caracteres")
    logging.info(f"   Caracteres disponibles: {sorted(diccionario_caracteres.keys())}")
else:
    logging.error(f"❌ La carpeta de templates no existe en: {TEMPLATE_FOLDER}")

# ═════════════════════════════════════════════════════════════
# GOOGLE SHEETS
# ═════════════════════════════════════════════════════════════

def obtener_datos_unicos():
    """
    Obtiene datos de MÚLTIPLES SHEETS:
    - Motos 0_5, Motos 6_10, Motos 11_15, Motos 16_25
    
    De cada sheet extrae:
    - Columna A: Cédula asociado
    - Columna C: Cédula propietario  
    - Columna E: Placa
    
    Retorna: [(cedula_asociado, cedula_propietario, placa, numero_fila, nombre_sheet), ...]
    """
    try:
        SCOPES = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = Credentials.from_service_account_file(
            str(GOOGLE_CREDS),
            scopes=SCOPES
        )

        client = gspread.authorize(creds)
        sheet = client.open_by_key("1saIDw37nd-rnzZvvKjxUQP41LhXJvSiayYgFRR78N7o")
        
        nombres_sheets = ["Motos 0_5", "Motos 6_10", "Motos 11_15", "Motos 16_25"]
        datos = []
        vistos = set()
        
        # ═════════════════════════════════════════════════════════════
        # 🆕 ESTADÍSTICAS POR SHEET
        # ═════════════════════════════════════════════════════════════
        estadisticas_sheets = {}
        total_recopilado = 0
        
        for nombre_sheet in nombres_sheets:
            try:
                worksheet = sheet.worksheet(nombre_sheet)
                logging.info(f"📊 Leyendo sheet: {nombre_sheet}")
                
                cedulas_asociado = worksheet.col_values(2)  # Columna B
                cedulas_propietario = worksheet.col_values(4)  # Columna D
                placas = worksheet.col_values(6)  # Columna F
                
                registros_sheet = 0  # Contador para este sheet
                
                for i in range(1, min(len(cedulas_asociado), len(cedulas_propietario), len(placas))):
                    cedula_asoc = str(cedulas_asociado[i]).strip()
                    cedula_prop = str(cedulas_propietario[i]).strip()
                    placa = str(placas[i]).strip()
                    
                    if placa and (cedula_asoc, placa) not in vistos:
                        vistos.add((cedula_asoc, placa))
                        datos.append((cedula_asoc, cedula_prop, placa, i + 1, nombre_sheet))
                        registros_sheet += 1  # Incrementar contador
                
                estadisticas_sheets[nombre_sheet] = registros_sheet
                total_recopilado += registros_sheet
                
                logging.info(f"   ✅ {registros_sheet} registros únicos de {nombre_sheet}")
                
            except gspread.WorksheetNotFound:
                logging.warning(f"⚠️  Sheet '{nombre_sheet}' no encontrada")
                estadisticas_sheets[nombre_sheet] = 0
                continue
        
        # ═════════════════════════════════════════════════════════════
        # 🆕 MOSTRAR ESTADÍSTICAS RESUMIDAS
        # ═════════════════════════════════════════════════════════════
        logging.info("\n" + "="*70)
        logging.info("📊 ESTADÍSTICAS DE DATOS RECOPILADOS")
        logging.info("="*70)
        
        for nombre_sheet in nombres_sheets:
            cantidad = estadisticas_sheets.get(nombre_sheet, 0)
            logging.info(f"   📋 {nombre_sheet}: {cantidad} registros")
        
        logging.info("="*70)
        
        duplicados_eliminados = total_recopilado - len(datos)
        
        logging.info(f"   ✅ TOTAL RECOPILADO: {total_recopilado} registros")
        logging.info(f"   ✅ DUPLICADOS ELIMINADOS: {duplicados_eliminados} registros")
        logging.info(f"   ✅ TOTAL PENDIENTE: {len(datos)} registros")
        logging.info("="*70 + "\n")
        
        logging.info(f"📊 Total pendientes: {len(datos)} registros")
        return datos
    
    except Exception as e:
        logging.error(f"❌ Error obteniendo datos de Sheets: {e}")
        return []


# ═════════════════════════════════════════════════════════════
# GUARDAR EN HOJA "RESULTADOS"
# ═════════════════════════════════════════════════════════════

def guardar_resultado_en_resultados(cedula_asociado, cedula_propietario, placa, cedula_usada, estado_final):
    """
    Guarda el resultado en la hoja 'Resultados':
    - A: Cédula asociado
    - B: Cédula propietario
    - C: Placa
    - D: Vacío (para separar)
    - E: "Funcionó" o "Falló"
    - F: Cédula que funcionó o falló (asociado o propietario)
    
    ⭐ ESCRIBE DESDE LA FILA 2 (sin tocar encabezados)
    """
    try:
        SCOPES = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = Credentials.from_service_account_file(
            str(GOOGLE_CREDS),
            scopes=SCOPES
        )

        client = gspread.authorize(creds)
        sheet = client.open_by_key("1vs414iH3QVeLoTcY2CExg4kD9eCkXZRRfax_WTlUXPk")
        
        # ═══ OBTENER WORKSHEET (asumiendo que ya existe con encabezados) ═══
        worksheet = sheet.worksheet("Resultados")
        logging.info("✅ Hoja 'Resultados' encontrada")
        
        # ═══ BUSCAR PRIMERA FILA VACÍA (a partir de fila 2) ═══
        todas_filas = worksheet.get_all_values()
        fila_nueva = len(todas_filas) + 1  # Próxima fila después de las existentes
        
        # Si la fila calculada es menor a 2, entonces es la primera vez
        if fila_nueva < 2:
            fila_nueva = 2
        
        logging.info(f"📝 Escribiendo en fila {fila_nueva}")
        
        # ═══ PREPARAR DATOS ═══
        estado_texto = "Funcionó" if estado_final == "Exitoso" else "Falló"
        nueva_fila = [
            cedula_asociado,
            cedula_propietario,
            placa,
            "",
            estado_texto,
            cedula_usada
        ]
        
        # ═══ ESCRIBIR DATOS (SIN TOCAR FILA 1) ═══
        rango = f"A{fila_nueva}:F{fila_nueva}"
        worksheet.update([nueva_fila], range_name=rango, value_input_option="RAW")
        
        logging.info(f"✅ Resultado guardado en 'Resultados' (fila {fila_nueva}):")
        logging.info(f"   Cédula Asociado: {cedula_asociado}")
        logging.info(f"   Cédula Propietario: {cedula_propietario}")
        logging.info(f"   Placa: {placa}")
        logging.info(f"   Estado: {estado_texto}")
        logging.info(f"   Cédula Usada: {cedula_usada}")
        
    except gspread.WorksheetNotFound:
        logging.error(f"❌ Hoja 'Resultados' NO encontrada. Debes crearla manualmente con los encabezados en la fila 1")
    except Exception as e:
        logging.error(f"❌ Error escribiendo en 'Resultados': {e}", exc_info=True)



def escribir_datos_vehiculo_sheets(datos_vehiculo, fila_destino=None):
    """
    Escribe los datos del vehículo en la hoja 'Datos Vehiculo' de Google Sheets.
    Si no se especifica fila, busca la primera fila vacía.
    """
    try:
        SCOPES = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = Credentials.from_service_account_file(
            str(GOOGLE_CREDS),
            scopes=SCOPES
        )

        client = gspread.authorize(creds)
        sheet = client.open_by_key("1oc6vcS6Y7i1IyxuEoFwsQAdPTMQCBJbixpmP8c-oOVw")
        
        # Acceder a la hoja "Datos Vehiculo"
        try:
            worksheet = sheet.worksheet("Datos Vehiculo")
        except gspread.WorksheetNotFound:
            worksheet = sheet.add_worksheet(title="Datos Vehiculo", rows=1000, cols=40)
            logging.info("📝 Hoja 'Datos Vehiculo' creada")
        
        # Definir el orden de las columnas
        encabezados = [
            "PLACA", "NRO_LICENCIA_TRANSITO", "ESTADO_VEHICULO", "TIPO_SERVICIO", 
            "CLASE_VEHICULO", "MARCA", "MODELO", "LINEA", "COLOR", "NUMERO_SERIE", 
            "NUMERO_MOTOR", "NUMERO_CHASIS", "NUMERO_VIN", "CILINDRAJE", 
            "TIPO_COMBUSTIBLE", "TIPO_CARROCERIA", "FECHA_MATRICULA_INICIAL", 
            "AUTORIDAD_TRANSITO", "CLASICO_ANTIGUO", "REGRABACION_MOTOR", 
            "REGRABACION_CHASIS", "REGRABACION_SERIE", "REGRABACION_VIN", 
            "NRO_REGRABACION_MOTOR", "NRO_REGRABACION_CHASIS", "NRO_REGRABACION_SERIE", 
            "NRO_REGRABACION_VIN", "VEHICULO_ENSENANZA", "GRAVAMENES_PROPIEDAD", 
            "REPOTENCIADO", "PUERTAS"
        ]
        
        # Verificar si la hoja tiene encabezados
        primera_fila = worksheet.row_values(1)
        if not primera_fila or primera_fila[0] != "PLACA":
            logging.info("📝 Escribiendo encabezados en 'Datos Vehiculo'")
            for col_num, encabezado in enumerate(encabezados, start=1):
                worksheet.update_cell(1, col_num, encabezado)
            fila_datos = 2
        else:
            if fila_destino:
                fila_datos = fila_destino
            else:
                todas_placas = worksheet.col_values(1)
                fila_datos = len(todas_placas) + 1
        
        # Preparar los datos
        fila_valores = []
        for encabezado in encabezados:
            valor = datos_vehiculo.get(encabezado, "No disponible")
            fila_valores.append(valor)
        
        # Escribir los datos
        for col_num, valor in enumerate(fila_valores, start=1):
            worksheet.update_cell(fila_datos, col_num, valor)
        
        logging.info(f"✅ Datos de vehículo escritos en fila {fila_datos} de 'Datos Vehiculo'")
        return fila_datos
        
    except Exception as e:
        logging.error(f"❌ Error escribiendo en 'Datos Vehiculo': {e}")
        return None

# ==================== FUNCIÓN MODIFICADA: guardar_en_sheets ====================

def guardar_en_sheets(resultados):
    """Guarda los resultados en Google Sheets - Cilindraje en columna D"""
    try:
        SCOPES = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = Credentials.from_service_account_file(
            str(GOOGLE_CREDS),
            scopes=SCOPES
        )

        client = gspread.authorize(creds)
        sheet = client.open_by_key("1vs414iH3QVeLoTcY2CExg4kD9eCkXZRRfax_WTlUXPk")
        worksheet = sheet.worksheet("Datos Runt")

        filas = []
        for r in resultados:
            # SOAT: 7 columnas
            soat_data = r.get("datos_soat")
            if not isinstance(soat_data, list) or len(soat_data) < 7:
                soat_data = ["No disponible"] * 7
            
            # RTM: 7 columnas
            rtm_data = r.get("datos_técnicos") 
            if not isinstance(rtm_data, list) or len(rtm_data) < 7:
                rtm_data = ["No disponible"] * 7

            # ESTRUCTURA CORRECTA:
            # A: Tiempo ejecucion
            # B: cedula
            # C: placa
            # D: CILINDRAJE (columna independiente)
            # E-K: SOAT (7 columnas)
            # L-R: RTM (7 columnas)
            # S: estado
            
            fila = [
                r["Tiempo ejecucion"],  # A
                r["cedula"],             # B
                r["placa"],              # C
                r.get("cilindraje", "No disponible"),  # D (NUEVO - independiente)
            ] + soat_data + rtm_data + [r["estado"]]   # E en adelante
            
            filas.append(fila)

        fila_inicio = max(len(worksheet.get_all_values()) + 1, 3)
        rango_inicio = f"A{fila_inicio}"

        worksheet.update(values=filas, range_name=rango_inicio, value_input_option="RAW")

        logging.info(f"✅ Se guardaron {len(filas)} filas en Google Sheets con cilindraje en columna D")

    except Exception as e:
        logging.error(f"❌ Error al guardar en Google Sheets: {e}")
# ==================== FIN FUNCIÓN MODIFICADA ====================


# ═════════════════════════════════════════════════════════════
# DRIVER - MANTIENE FUNCIÓN ORIGINAL
# ═════════════════════════════════════════════════════════════

def iniciar_driver(max_attempts=3):
    """Inicia el driver de Chrome"""
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

    logging.error("❌ No se pudo iniciar el driver tras varios intentos.")
    return None

def cerrar_driver(driver):
    """Cierra el driver de Chrome"""
    try:
        if driver:
            driver.quit()
            logging.info("✅ Driver cerrado correctamente.")
    except WebDriverException as e:
        logging.error(f"❌ Error al cerrar el driver: {e}")

# ═════════════════════════════════════════════════════════════
# LIMPIEZA DE CAMPOS
# ═════════════════════════════════════════════════════════════

def limpiar_campo_input(driver, element, campo_nombre=""):
    """Limpia un campo de input de forma AGRESIVA"""
    try:
        element.click()
        time.sleep(0.3)
        
        element.clear()
        time.sleep(0.3)
        
        element.send_keys(Keys.CONTROL + "a")
        time.sleep(0.2)
        element.send_keys(Keys.DELETE)
        time.sleep(0.3)
        
        for _ in range(20):
            element.send_keys(Keys.BACKSPACE)
            time.sleep(0.05)
        
        time.sleep(0.3)
        
        driver.execute_script("arguments[0].value = '';", element)
        time.sleep(0.2)
        
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
        """, element)
        time.sleep(0.3)
        
        logging.info(f"✅ Campo '{campo_nombre}' limpiado")
        
    except Exception as e:
        logging.error(f"⚠️ Error al limpiar el campo '{campo_nombre}': {e}")

def limpiar_todos_los_campos(driver):
    """Limpia TODOS los campos de input de la página"""
    try:
        logging.info("🧹 Limpiando todos los campos de la página...")
        
        todos_los_inputs = driver.find_elements(By.TAG_NAME, "input")
        
        for i, inp in enumerate(todos_los_inputs):
            try:
                if inp.is_displayed():
                    inp.click()
                    time.sleep(0.1)
                    inp.clear()
                    time.sleep(0.1)
                    inp.send_keys(Keys.CONTROL + "a")
                    time.sleep(0.05)
                    inp.send_keys(Keys.DELETE)
                    time.sleep(0.1)
                    
                    driver.execute_script("arguments[0].value = '';", inp)
                    time.sleep(0.05)
            except:
                pass
        
        time.sleep(0.5)
        logging.info(f"✅ Se limpiaron {len(todos_los_inputs)} campos")
        
    except Exception as e:
        logging.error(f"⚠️ Error limpiando campos: {e}")

def limpiar_campos_individuales_validado(driver, cedula, placa):
    """Limpia los campos uno por uno de forma granular y validada"""
    try:
        logging.info("🔧 Iniciando limpieza individual y validada de campos...")
        
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        # ═══ LIMPIAR CAMPO PLACA ═══
        try:
            logging.info("🧹 Limpiando campo PLACA...")
            placa_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="placa"]'))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", placa_input)
            time.sleep(0.5)
            limpiar_campo_input(driver, placa_input, "PLACA")
            logging.info(f"   ✅ Campo PLACA limpiado correctamente")
            
        except Exception as e:
            logging.error(f"   ⚠️ Error limpiando PLACA: {e}")
        
        time.sleep(0.5)
        
        # ═══ LIMPIAR CAMPO CÉDULA ═══
        try:
            logging.info("🧹 Limpiando campo CÉDULA...")
            cedula_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="documento"]'))
            )
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cedula_input)
            time.sleep(0.5)
            limpiar_campo_input(driver, cedula_input, "CÉDULA")
            logging.info(f"   ✅ Campo CÉDULA limpiado correctamente")
            
        except Exception as e:
            logging.error(f"   ⚠️ Error limpiando CÉDULA: {e}")
        
        time.sleep(0.5)
        
        # ═══ LIMPIAR CAMPO CAPTCHA ═══
        try:
            logging.info("🧹 Limpiando campo CAPTCHA...")
            
            driver.execute_script("window.scrollBy(0, 300);")
            time.sleep(1)
            
            captcha_inputs = driver.find_elements(By.TAG_NAME, "input")
            captcha_field = None
            
            for inp in reversed(captcha_inputs):
                if inp.is_displayed():
                    captcha_field = inp
                    break
            
            if captcha_field:
                limpiar_campo_input(driver, captcha_field, "CAPTCHA")
                logging.info(f"   ✅ Campo CAPTCHA limpiado correctamente")
            else:
                logging.warning(f"   ⚠️ No se encontró campo de CAPTCHA")
            
        except Exception as e:
            logging.error(f"   ⚠️ Error limpiando CAPTCHA: {e}")
        
        logging.info("✅ Limpieza individual completada")
        return True
        
    except Exception as e:
        logging.error(f"❌ Error en limpieza individual: {e}")
        return False

# ═════════════════════════════════════════════════════════════
# FUNCIONES ANGULAR MATERIAL - ⭐ MEJORADAS
# ════════════════════��════════════════════════════════════════

def abrir_seccion_angular(driver, texto_del_titulo):
    """Abre paneles de Angular Material buscando por el texto del título"""
    try:
        logging.info(f"🔍 Intentando abrir panel: {texto_del_titulo}")
        xpath_header = f"//mat-expansion-panel-header[.//mat-panel-title[contains(., '{texto_del_titulo}')]]"
        
        header = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath_header))
        )
        
        if header.get_attribute("aria-expanded") == "true":
            logging.info(f"✅ El panel '{texto_del_titulo}' ya está abierto.")
            return True

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", header)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", header)
        time.sleep(2)
        return True
    except Exception as e:
        logging.error(f"❌ No se pudo abrir {texto_del_titulo}: {e}")
        return False

def cerrar_todos_los_paneles(driver):
    """⭐ NUEVA FUNCIÓN: Cierra TODOS los paneles expandidos de Angular Material"""
    try:
        logging.info("🔄 Cerrando todos los paneles...")
        
        script_cerrar = """
        const headers = document.querySelectorAll('mat-expansion-panel-header[aria-expanded="true"]');
        headers.forEach(header => {
            header.click();
        });
        return headers.length;
        """
        
        cantidad = driver.execute_script(script_cerrar)
        logging.info(f"✅ Se cerraron {cantidad} paneles")
        time.sleep(1)
        return True
    except Exception as e:
        logging.error(f"⚠️ Error cerrando paneles: {e}")
        return False

# ═════════════════════════════════════════════════════════════
# CAPTCHA
# ═════════════════════════════════════════════════════════════

def capturar_captcha(driver, placa, carpeta_temp=CAPTCHA_FOLDER):
    """Captura la imagen del captcha"""
    try:
        logging.info("🔍 Buscando imagen del captcha...")
        
        xpath_imagen = "/html/body/host-runt-root/app-layout/app-theme-runt2/mat-sidenav-container/mat-sidenav-content/div/ng-component/div/div[2]/div[1]/form/div[2]/div/mat-card/mat-card-content/div[7]/div[3]/img"
        
        try:
            captcha_img_element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, xpath_imagen))
            )
        except TimeoutException:
            logging.warning("⚠️ XPath original falló, buscando alternativas...")
            all_images = driver.find_elements(By.TAG_NAME, "img")
            
            captcha_img_element = None
            for img in all_images:
                if img.is_displayed():
                    captcha_img_element = img
                    break
            
            if not captcha_img_element:
                logging.error("❌ No hay imagen visible del captcha")
                return None, None
        
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", captcha_img_element)
        time.sleep(1)

        captcha_screenshot = captcha_img_element.screenshot_as_png
        
        nparr = np.frombuffer(captcha_screenshot, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_GRAYSCALE)
        
        _, img_bin = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        img_pil = Image.fromarray(img_bin)

        timestamp = datetime.now().strftime('%H%M%S')
        nombre_archivo = f"captcha_{str(placa).strip()}_{timestamp}.png"
        temp_filename = carpeta_temp / nombre_archivo

        img_pil.save(str(temp_filename))
        
        return temp_filename, img_pil

    except Exception as e:
        logging.error(f"❌ Error al capturar captcha: {e}")
        return None, None

def verificar_caracter_en_templates(caracter):
    """Verifica si un carácter está en el diccionario de templates"""
    return caracter in diccionario_caracteres

def obtener_caracter_del_diccionario(caracter):
    """Obtiene el carácter confirmado del diccionario"""
    if caracter in diccionario_caracteres:
        logging.info(f"   ✅ '{caracter}' ENCONTRADO en templates")
        captcha_logger.info(f"      → '{caracter}' CONFIRMADO en diccionario")
        return caracter
    else:
        logging.warning(f"   ⚠️ '{caracter}' NO está en templates, usando como está")
        return caracter

def resolver_captcha(img_pil: Image.Image, placa: str) -> Optional[str]:
    """Resuelve el captcha con verificación de templates"""
    
    custom_config = r'--oem 3 --psm 8 -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'

    try:
        img_array = np.array(img_pil)
        texto_imagen_original = pytesseract.image_to_string(img_array, config=custom_config).strip()
        
        text_tesseract = pytesseract.image_to_string(img_pil, config=custom_config).strip()

        logging.info(f"📝 Tesseract extrajo: '{text_tesseract}'")
        
        captcha_logger.info(f"Placa: {placa} | TEXTO IMAGEN: {texto_imagen_original} | TESSERACT LEYÓ: {text_tesseract}")

        if text_tesseract and all(c.isalnum() for c in text_tesseract):
            
            reemplazos = {
                "O": "0",
                "l": "1",
                "I": "1",
            }
            
            texto_corregido = ""
            for c in text_tesseract:
                texto_corregido += reemplazos.get(c, c)
            
            logging.info(f"📝 Después limpieza básica: '{texto_corregido}'")
            
            logging.info(f"🔍 Verificando caracteres contra diccionario...")
            texto_final = ""
            
            for i, char in enumerate(texto_corregido):
                logging.info(f"   Carácter {i+1}: '{char}'")
                
                if verificar_caracter_en_templates(char):
                    char_confirmado = obtener_caracter_del_diccionario(char)
                    texto_final += char_confirmado
                else:
                    logging.info(f"   ⚠️ '{char}' NO en diccionario, manteniendo...")
                    captcha_logger.info(f"      → '{char}' NO en diccionario")
                    texto_final += char
            
            logging.info(f"✅ Captcha después verificación: '{texto_final}'")
            captcha_logger.info(f"         CORREGIDO: {texto_final}")
            
            return texto_final

        logging.warning(f"⚠️ Tesseract extrajo algo inválido: {text_tesseract}")
        captcha_logger.info(f"Placa: {placa} | TEXTO IMAGEN: {texto_imagen_original} | RESULTADO: ❌ NO LEGIBLE")
        return None

    except Exception as e:
        logging.error(f"❌ Error OCR: {e}")
        return None

def detectar_mensaje_error(driver):
    """Detecta el tipo de mensaje de error usando XPath específicos"""
    try:
        logging.info("🔍 Detectando tipo de mensaje de error...")
        
        # XPaths más robustos para SweetAlert2
        contenedor_modal = "//div[contains(@class, 'swal2-popup')]"
        xpath_mensaje_texto = "//div[contains(@class, 'swal2-html-container')]"
        
        try:
            modal_container = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, contenedor_modal))
            )
            
            logging.info("✅ Modal detectado")
            time.sleep(1)
            
            try:
                mensaje_element = driver.find_element(By.XPATH, xpath_mensaje_texto)
                texto_mensaje = mensaje_element.text.strip()
                
                logging.info(f"📄 Mensaje capturado: {texto_mensaje[:100]}")
                
                texto_lower = texto_mensaje.lower()
                
                if "no corresponden" in texto_lower or "propietarios activos" in texto_lower:
                    logging.warning("❌ DETECCIÓN: No hay personas asociadas al vehículo")
                    captcha_logger.info(f"         RESULTADO: ❌ NO HAY PERSONAS ASOCIADAS")
                    return "no_personas"
                
                if "captcha" in texto_lower and ("incorrecto" in texto_lower or "inválido" in texto_lower or "válido" in texto_lower):
                    logging.warning("❌ DETECCIÓN: Captcha incorrecto o inválido")
                    captcha_logger.info(f"         RESULTADO: ❌ CAPTCHA INCORRECTO")
                    return "captcha_incorrecto"
                
                logging.warning(f"⚠️ Modal detectado pero no reconocido: {texto_mensaje[:80]}")
                return "error_desconocido"
                
            except NoSuchElementException:
                logging.warning("⚠️ Modal visible pero no se encontró el texto del mensaje")
                return None
                
        except TimeoutException:
            logging.info("ℹ️ No hay modal de error visible (respuesta OK)")
            return None
            
    except Exception as e:
        logging.error(f"⚠️ Error detectando mensaje: {e}")
        return None

def leer_mensaje_no_disponible(driver):
    """Lee el mensaje cuando no hay información registrada"""
    try:
        logging.info("📸 Leyendo mensaje 'No se encontró información'...")
        
        try:
            xpath_msg = "//div[contains(text(), 'No se encontró información registrada')]"
            msg_element = driver.find_element(By.XPATH, xpath_msg)
            if msg_element.is_displayed():
                return msg_element.text.strip()
        except:
            pass
        
        try:
            elemento = driver.find_element(By.CSS_SELECTOR, "div.ng-star-inserted")
            if elemento.is_displayed():
                return elemento.text.strip() if elemento.text else "No disponible"
        except:
            pass
        
        return "No disponible"
        
    except Exception as e:
        logging.warning(f"⚠️ No se pudo leer mensaje: {e}")
        return "No disponible"


def limpiar_texto_celda(texto):
    """Limpia el texto de las celdas eliminando íconos y formatos especiales"""
    if not texto:
        return "No disponible"
    
    # Si contiene check_circle o cancel, extraer solo el estado
    if "check_circle" in texto or "cancel" in texto:
        # Separar por saltos de línea y tomar la última parte
        partes = texto.split('\n')
        if len(partes) > 1:
            return partes[-1].strip()  # Toma "VIGENTE" o "NO VIGENTE"
    
    # Si no hay íconos, devolver el texto original
    return texto.strip()

# ═════════════════════════════════════════════════════════════
# EXTRACCIÓN DE DATOS - ⭐ MEJORADA CON CIERRE DE PANELES
# ═════════════════════════════════════════════════════════════

def extraer_datos_soat(driver):
    """⭐ Extrae datos del SOAT - SOLO 7 COLUMNAS (sin cilindraje)"""
    try:
        logging.info("📋 INICIANDO EXTRACCIÓN DE SOAT...")
        logging.info("="*70)
        
        cerrar_todos_los_paneles(driver)
        time.sleep(1)
        
        if abrir_seccion_angular(driver, "Póliza SOAT"):
            time.sleep(2)
            
            try:
                xpath_primera_fila = "//mat-table//mat-row[1]"
                fila_mas_reciente = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, xpath_primera_fila))
                )
                
                celdas = fila_mas_reciente.find_elements(By.TAG_NAME, "mat-cell")
                datos_soat = [limpiar_texto_celda(celda.text) for celda in celdas]
                
                if not datos_soat or not datos_soat[0]:
                    return ["No disponible"] * 7

                # Asegurar que tenemos 7 elementos
                while len(datos_soat) < 7:
                    datos_soat.append("No disponible")
                    
                return datos_soat[:7]  # Solo 7 columnas

            except Exception as e:
                logging.warning(f"⚠️ No se encontró tabla de SOAT: {e}")
                return ["No disponible"] * 7
        else:
            return ["No disponible"] * 7
            
    except Exception as e:
        logging.error(f"❌ Error en extraer_datos_soat: {e}")
        return ["No disponible"] * 7



def extraer_datos_rtm(driver):
    """⭐ Extrae datos de RTM - Toma la PRIMERA FILA de la tabla (más reciente)"""
    try:
        logging.info("📋 INICIANDO EXTRACCIÓN DE RTM...")
        logging.info("="*70)
        
        # Cerramos paneles previos para evitar que se tapen los elementos
        cerrar_todos_los_paneles(driver)
        time.sleep(1)
        
        if abrir_seccion_angular(driver, "(RTM)"):
            time.sleep(2)
            
            try:
                # Cambiado: Selecciona el primer mat-row sin filtrar por texto "VIGENTE"
                xpath_primera_fila = "//mat-expansion-panel[.//mat-panel-title[contains(., '(RTM)')]]//mat-table//mat-row[contains(@class, 'mat-row')][1]"
                
                fila_mas_reciente = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.XPATH, xpath_primera_fila))
                )
                
                celdas = fila_mas_reciente.find_elements(By.TAG_NAME, "mat-cell")
                datos_crudos = [celda.text.strip() for celda in celdas]

                # Mapeo de columnas según la imagen:
                # [0] Tipo Revisión
                # [1] Fecha Expedición
                # [2] Fecha Vigencia
                # [3] CDA expide RTM
                # [4] Vigente ⬅️ ESTO DEBE IR EN LA POSICIÓN 6 DE GOOGLE SHEETS
                # [5] Nro. certificado
                # [6] Información consistente

                # Reordenar para que coincida con Google Sheets
                datos = [
                    datos_crudos[0],  # Tipo Revisión
                    datos_crudos[1],  # Fecha Expedición
                    datos_crudos[2],  # Fecha Vigencia
                    datos_crudos[3],  # CDA expide RTM
                    datos_crudos[5],  # Nro. certificado ⬅️ Intercambiado
                    datos_crudos[4],  # Vigente ⬅️ Intercambiado
                    datos_crudos[6]   # Información consistente
                ]
                
                if not datos or not datos[0]:
                    return ["No disponible"] * 7

                logging.info(f"✅ Datos RTM encontrados (Fila 1)")
                return datos

            except Exception:
                logging.info(f"⚠️ No se encontró tabla de datos, leyendo mensaje...")
                mensaje = leer_mensaje_no_disponible(driver)
                return [mensaje] * 7 
        else:
            return ["No disponible"] * 7
            
    except Exception as e:
        logging.error(f"❌ Error en extraer_datos_rtm: {e}")
        return ["No disponible"] * 7


def extraer_datos_vehiculo_optimizado(driver):
    """
    ⭐ ESTRATEGIA MEJORADA: Extrae labels en orden y busca valores consecutivos
    SIN SALTARSE A CAMPOS LEJANOS
    """
    try:
        logging.info("\n" + "="*70)
        logging.info("🚗 EXTRAYENDO DATOS COMPLETOS DEL VEHÍCULO...")
        logging.info("="*70)
        
        panel_content = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.panel-content"))
        )
        
        logging.info("✅ Panel de contenido encontrado")
        
        # ═══ MAPEO CORRECTO - ORDEN IMPORTA ═══
        mapeo_labels = [
            ("PLACA DEL VEHÍCULO", "PLACA"),
            ("NRO. DE LICENCIA DE TRÁNSITO", "NRO_LICENCIA_TRANSITO"),
            ("ESTADO DEL VEHÍCULO", "ESTADO_VEHICULO"),
            ("TIPO DE SERVICIO", "TIPO_SERVICIO"),
            ("CLASE DE VEHÍCULO", "CLASE_VEHICULO"),
            ("MARCA", "MARCA"),
            ("LÍNEA", "LINEA"),
            ("MODELO", "MODELO"),
            ("COLOR", "COLOR"),
            ("NÚMERO DE SERIE", "NUMERO_SERIE"),
            ("NÚMERO DE MOTOR", "NUMERO_MOTOR"),
            ("NÚMERO DE CHASIS", "NUMERO_CHASIS"),
            ("NÚMERO DE VIN", "NUMERO_VIN"),
            ("CILINDRAJE", "CILINDRAJE"),
            ("TIPO DE CARROCERÍA", "TIPO_CARROCERIA"),
            ("TIPO COMBUSTIBLE", "TIPO_COMBUSTIBLE"),
            ("FECHA DE MATRICULA INICIAL", "FECHA_MATRICULA_INICIAL"),
            ("AUTORIDAD DE TRÁNSITO", "AUTORIDAD_TRANSITO"),
            ("GRAVÁMENES A LA PROPIEDAD", "GRAVAMENES_PROPIEDAD"),
            ("CLÁSICO O ANTIGUO", "CLASICO_ANTIGUO"),
            ("REPOTENCIADO", "REPOTENCIADO"),
            ("REGRABACIÓN MOTOR (SI/NO)", "REGRABACION_MOTOR"),
            ("NRO. REGRABACIÓN MOTOR", "NRO_REGRABACION_MOTOR"),
            ("REGRABACIÓN CHASIS (SI/NO)", "REGRABACION_CHASIS"),
            ("NRO. REGRABACIÓN CHASIS", "NRO_REGRABACION_CHASIS"),
            ("REGRABACIÓN SERIE (SI/NO)", "REGRABACION_SERIE"),
            ("NRO. REGRABACIÓN SERIE", "NRO_REGRABACION_SERIE"),
            ("REGRABACIÓN VIN (SI/NO)", "REGRABACION_VIN"),
            ("NRO. REGRABACIÓN VIN", "NRO_REGRABACION_VIN"),
            ("VEHÍCULO ENSEÑANZA (SI/NO)", "VEHICULO_ENSENANZA"),
            ("PUERTAS", "PUERTAS"),
        ]
        
        datos_vehiculo = {campo: "No disponible" for etiqueta, campo in mapeo_labels}
        
        # ═══ OBTENER TODOS LOS LABELS ═══
        logging.info("🔍 Obteniendo todos los labels...")
        labels = driver.find_elements(By.CSS_SELECTOR, "div.panel-content label")
        logging.info(f"   Encontrados {len(labels)} labels")
        
        # Crear diccionario de labels para búsqueda rápida
        label_dict = {}
        for i, label in enumerate(labels):
            label_text = label.text.strip()
            if label_text:
                label_dict[label_text] = label
                logging.info(f"   [{i}] {label_text}")
        
        # 🔥 CREAR LISTA DE ETIQUETAS COMPLETAS (para rechazar correctamente)
        etiquetas_completas = [etiqueta.upper() for etiqueta, _ in mapeo_labels]
        
        # ═══ PROCESAR CADA LABEL EN ORDEN ═══
        logging.info("\n🔍 EXTRAYENDO PARES LABEL-VALOR...")
        
        for etiqueta_buscada, campo_normalizado in mapeo_labels:
            logging.info(f"\n📌 Buscando: {etiqueta_buscada}")
            
            # Buscar coincidencia flexible del label
            label_encontrado = None
            for label_text, label_elem in label_dict.items():
                if etiqueta_buscada.upper() in label_text.upper() or label_text.upper() in etiqueta_buscada.upper():
                    label_encontrado = label_elem
                    logging.info(f"   ✓ Label encontrado: '{label_text}'")
                    break
            
            if not label_encontrado:
                logging.warning(f"   ✗ Label no encontrado")
                continue
            
            # ═══ BÚSQUEDA CONFINADA AL CONTENEDOR LOCAL ═══
            valor = None
            
            try:
                padre_label = label_encontrado.find_element(By.XPATH, "..")
                
                # Estrategia 1: <b> directo en hermanos inmediatos
                try:
                    bold_hermano = padre_label.find_element(By.XPATH, ".//b")
                    valor = bold_hermano.text.strip()
                    if valor:
                        logging.info(f"   ✓ [LOCAL-b] Valor: {valor}")
                except:
                    pass
                
                # Estrategia 2: Si no, buscar en el siguiente div INMEDIATO (CONFINADO)
                if not valor:
                    try:
                        siguiente_div = padre_label.find_element(By.XPATH, "./following-sibling::div[1]")
                        try:
                            bold_en_siguiente = siguiente_div.find_element(By.XPATH, ".//b")
                            valor = bold_en_siguiente.text.strip()
                            if valor:
                                logging.info(f"   ✓ [SIGUIENTE-LOCAL-b] Valor: {valor}")
                        except:
                            texto_siguiente = siguiente_div.text.strip()
                            if texto_siguiente and len(texto_siguiente) < 100 and not any(c in texto_siguiente for c in [':', '(', ')']):
                                valor = texto_siguiente
                                logging.info(f"   ✓ [SIGUIENTE-LOCAL-texto] Valor: {valor}")
                    except:
                        pass
                
                # Estrategia 3: Buscar en el padre del padre (subir un nivel más)
                if not valor:
                    try:
                        abuelo = padre_label.find_element(By.XPATH, "..")
                        siguiente_del_abuelo = abuelo.find_element(By.XPATH, "./following-sibling::*[1]")
                        try:
                            bold_en_abuelo = siguiente_del_abuelo.find_element(By.XPATH, ".//b")
                            valor = bold_en_abuelo.text.strip()
                            if valor:
                                logging.info(f"   ✓ [ABUELO-LOCAL-b] Valor: {valor}")
                        except:
                            pass
                    except:
                        pass
                
            except Exception as e:
                logging.warning(f"   ⚠️ Error en búsqueda confinada: {e}")
            
            # ═══ VALIDAR QUE EL VALOR SEA REALMENTE VÁLIDO ═══
            if valor:
                valor_upper = valor.upper().strip()
                
                # 🔥 NUEVA VALIDACIÓN: Rechazar SOLO si es EXACTAMENTE otra etiqueta
                es_etiqueta_exacta = valor_upper in etiquetas_completas
                
                # TAMBIÉN rechazar si es SOLO números muy pequeños (como 1, 2, 3)
                es_numero = valor.isdigit()
                es_numero_valido = not (es_numero and len(valor) < 2)
                
                # RECHAZAR: Si es exactamente otra etiqueta O si es un número muy pequeño
                if es_etiqueta_exacta:
                    logging.warning(f"   ❌ Valor rechazado (es otra etiqueta): '{valor}'")
                    valor = None
                elif not es_numero_valido:
                    logging.warning(f"   ❌ Valor rechazado (número inválido): '{valor}'")
                    valor = None
                # ACEPTAR: Todo lo demás
                else:
                    logging.info(f"   ✅ Valor aceptado: {valor}")
            
            # ═══ GUARDAR RESULTADO FINAL ═══
            if valor:
                datos_vehiculo[campo_normalizado] = valor
                logging.info(f"   ✅ MAPEADO: {campo_normalizado} = {valor}")
            else:
                datos_vehiculo[campo_normalizado] = "No disponible"
                logging.warning(f"   ✗ {campo_normalizado} → No disponible")
        
        # ═══ LOG RESUMEN ═══
        campos_llenos = sum(1 for v in datos_vehiculo.values() if v != "No disponible")
        logging.info(f"\n{'='*70}")
        logging.info(f"✅ EXTRACCIÓN COMPLETADA: {campos_llenos}/{len(datos_vehiculo)} campos")
        logging.info(f"{'='*70}")
        
        for campo, valor in sorted(datos_vehiculo.items()):
            if valor != "No disponible":
                logging.info(f"   ✓ {campo}: {valor}")
            else:
                logging.info(f"   ✗ {campo}: No disponible")
        
        return datos_vehiculo
        
    except Exception as e:
        logging.error(f"❌ Error grave: {e}", exc_info=True)
        # ... (retorno de diccionario con "No disponible")
        return {campo: "No disponible" for etiqueta, campo in mapeo_labels}

def escribir_datos_vehiculo_en_sheets(datos_vehiculo, cedula, placa):
    """
    ⭐ ESCRIBE en Google Sheets hoja 'Datos Vehiculo'
    - Crea encabezados una sola vez
    - Escribe los datos en la siguiente fila
    """
    try:
        logging.info("\n" + "="*70)
        logging.info("💾 ESCRIBIENDO DATOS EN SHEETS 'Datos Vehiculo'...")
        logging.info("="*70)
        
        SCOPES = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = Credentials.from_service_account_file(
            str(GOOGLE_CREDS),
            scopes=SCOPES
        )

        client = gspread.authorize(creds)
        # MISMA ID donde están SOAT y RTM
        sheet = client.open_by_key("1vs414iH3QVeLoTcY2CExg4kD9eCkXZRRfax_WTlUXPk")
        
        # Obtener o crear worksheet "Datos Vehiculo"
        try:
            worksheet = sheet.worksheet("Datos Vehiculo")
            logging.info("✅ Hoja 'Datos Vehiculo' encontrada")
        except gspread.WorksheetNotFound:
            worksheet = sheet.add_worksheet(title="Datos Vehiculo", rows=1000, cols=40)
            logging.info("📝 Hoja 'Datos Vehiculo' creada")
        
        # Orden de columnas (31 campos)
        encabezados = [
            "PLACA", "NRO_LICENCIA_TRANSITO", "ESTADO_VEHICULO", "TIPO_SERVICIO", 
            "CLASE_VEHICULO", "MARCA", "MODELO", "LINEA", "COLOR", "NUMERO_SERIE", 
            "NUMERO_MOTOR", "NUMERO_CHASIS", "NUMERO_VIN", "CILINDRAJE", 
            "TIPO_COMBUSTIBLE", "TIPO_CARROCERIA", "FECHA_MATRICULA_INICIAL", 
            "AUTORIDAD_TRANSITO", "CLASICO_ANTIGUO", "REGRABACION_MOTOR", 
            "REGRABACION_CHASIS", "REGRABACION_SERIE", "REGRABACION_VIN", 
            "NRO_REGRABACION_MOTOR", "NRO_REGRABACION_CHASIS", "NRO_REGRABACION_SERIE", 
            "NRO_REGRABACION_VIN", "VEHICULO_ENSENANZA", "GRAVAMENES_PROPIEDAD", 
            "REPOTENCIADO", "PUERTAS"
        ]
        
        # ═══ VERIFICAR ENCABEZADOS ═══
        try:
            primera_fila = worksheet.row_values(1)
        except:
            primera_fila = []
        
        if not primera_fila or primera_fila[0] != "PLACA":
            logging.info("📝 Escribiendo encabezados...")
            worksheet.update([encabezados], range_name="A1", value_input_option="RAW")
            fila_datos = 2
            logging.info(f"✅ {len(encabezados)} encabezados escritos en fila 1")
        else:
            # Buscar primera fila vacía
            todas_placas = worksheet.col_values(1)
            fila_datos = len(todas_placas) + 1
            logging.info(f"✅ Encabezados ya existen, escribiendo en fila {fila_datos}")
        
        # ═══ PREPARAR VALORES EN ORDEN ═══
        fila_valores = []
        for encabezado in encabezados:
            valor = datos_vehiculo.get(encabezado, "No disponible")
            fila_valores.append(valor)
        
        # ═══ ESCRIBIR EN SHEETS ═══
        rango = f"A{fila_datos}:AF{fila_datos}"  # A a AF (32 columnas)
        
        try:
            worksheet.update([fila_valores], range_name=rango, value_input_option="RAW")
            logging.info(f"✅ {len(fila_valores)} datos escritos en fila {fila_datos}")
            logging.info(f"   Placa: {datos_vehiculo.get('PLACA', 'N/A')}")
            logging.info(f"   Marca: {datos_vehiculo.get('MARCA', 'N/A')}")
            logging.info(f"   Modelo: {datos_vehiculo.get('MODELO', 'N/A')}")
            return fila_datos
            
        except Exception as e:
            logging.error(f"❌ Error en update: {e}")
            logging.warning("🔄 Intentando método alternativo...")
            
            for col_num, valor in enumerate(fila_valores, start=1):
                try:
                    worksheet.update_cell(fila_datos, col_num, valor)
                    time.sleep(0.05)
                except:
                    pass
            
            logging.info(f"✅ Datos escritos (método alternativo) en fila {fila_datos}")
            return fila_datos
        
    except Exception as e:
        logging.error(f"❌ Error escribiendo en Sheets: {e}", exc_info=True)
        return None


# ═════════════════════════════════════════════════════════════
# 🔥 MANEJO UNIVERSAL DE MODALES - ⭐ NUEVO
# ═════════════════════════════════════════════════════════════

def detectar_y_cerrar_modal_universal(driver, timeout_max=5, max_intentos=4):
    """
    🔥 UNIVERSAL: Detecta y cierra CUALQUIER modal de error.
    
    - No depende del contenido del mensaje
    - Intenta múltiples estrategias de cierre
    - Si falla, recarga la página automáticamente
    - Retorna True si se cerró, False si recargó
    
    Args:
        driver: Selenium WebDriver
        timeout_max: Segundos máximos para detectar modal
        max_intentos: Intentos de click antes de reload
    
    Returns:
        True: Modal cerrado exitosamente
        False: Página recargada (modal persistía)
    """
    
    try:
        logging.info("🔍 [MODAL DETECTOR] Buscando modales abiertos...")
        
        # ═══ DETECTAR SI HAY MODAL VISIBLE ═══
        selectores_modal = [
            ("swal2", "div.swal2-container"),
            ("angular-material", "div.mat-dialog-container"),
            ("bootstrap", ".modal.show"),
            ("genérico-overlay", "div[role='dialog']"),
        ]
        
        modal_encontrado = False
        modal_element = None
        
        for tipo_modal, selector in selectores_modal:
            try:
                elementos = driver.find_elements(By.CSS_SELECTOR, selector)
                modales_visibles = [e for e in elementos if e.is_displayed()]
                
                if modales_visibles:
                    modal_element = modales_visibles[0]
                    modal_encontrado = True
                    logging.info(f"✅ Modal detectado: {tipo_modal}")
                    break
            except:
                continue
        
        if not modal_encontrado:
            logging.info("ℹ️  No hay modales visibles en pantalla")
            return True  # No hay modal, continuamos
        
        # ═══ INTENTAR CERRAR CON MÚLTIPLES ESTRATEGIAS ═══
        estrategias_cierre = [
            # Estrategia 1: Botones por texto
            ("XPATH - Botón 'Aceptar'", "//button[contains(translate(., 'ÁÉÍÓÚ', 'aeiou'), 'aceptar')]"),
            ("XPATH - Botón 'OK'", "//button[contains(., 'OK')]"),
            ("XPATH - Botón 'Cerrar'", "//button[contains(translate(., 'ÁÉÍÓÚ', 'aeiou'), 'cerrar')]"),
            ("XPATH - Botón 'Continuar'", "//button[contains(translate(., 'ÁÉÍÓÚ', 'aeiou'), 'continuar')]"),
            
            # Estrategia 2: Selectores SweetAlert
            ("CSS - SweetAlert2 Confirm", "button.swal2-confirm"),
            ("CSS - SweetAlert2 Cancel", "button.swal2-cancel"),
            ("CSS - SweetAlert Action Button", ".swal2-actions button"),
            
            # Estrategia 3: Selectores Angular Material
            ("XPATH - MAT Dialog Close", "//button[@mat-dialog-close]"),
            ("CSS - MAT Dialog Close", "button[mat-dialog-close]"),
            
            # Estrategia 4: Botones genéricos en modal
            ("CSS - Primer botón en modal", "div[role='dialog'] button, .modal button"),
        ]
        
        for intento in range(max_intentos):
            logging.info(f"\n🔄 Intento {intento + 1}/{max_intentos}")
            
            cerrado = False
            
            for nombre_estrategia, selector in estrategias_cierre:
                try:
                    if selector.startswith("//"):
                        # XPath
                        logging.info(f"   🔍 {nombre_estrategia}")
                        botones = driver.find_elements(By.XPATH, selector)
                        botones_visibles = [b for b in botones if b.is_displayed()]
                        
                        if botones_visibles:
                            boton = botones_visibles[0]
                            logging.info(f"      ✓ Elemento encontrado")
                            
                            # Scroll al botón
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton)
                            time.sleep(0.3)
                            
                            # Intento 1: Click normal
                            try:
                                boton.click()
                                logging.info(f"      ✓ Click normal ejecutado")
                                cerrado = True
                            except:
                                # Intento 2: Click con JavaScript
                                driver.execute_script("arguments[0].click();", boton)
                                logging.info(f"      ✓ Click JavaScript ejecutado")
                                cerrado = True
                        else:
                            continue
                    else:
                        # CSS Selector
                        logging.info(f"   🔍 {nombre_estrategia}")
                        botones = driver.find_elements(By.CSS_SELECTOR, selector)
                        botones_visibles = [b for b in botones if b.is_displayed()]
                        
                        if botones_visibles:
                            boton = botones_visibles[0]
                            logging.info(f"      ✓ Elemento encontrado")
                            
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton)
                            time.sleep(0.3)
                            
                            try:
                                boton.click()
                                logging.info(f"      ✓ Click normal ejecutado")
                                cerrado = True
                            except:
                                driver.execute_script("arguments[0].click();", boton)
                                logging.info(f"      ✓ Click JavaScript ejecutado")
                                cerrado = True
                        else:
                            continue
                    
                    if cerrado:
                        time.sleep(1)
                        break  # Salir del loop de estrategias si uno funcionó
                        
                except Exception as e:
                    continue
            
            # ═══ VERIFICAR SI EL MODAL DESAPARECIÓ ═══
            time.sleep(1)
            try:
                modales_actuales = []
                for tipo, selector in selectores_modal:
                    elementos = driver.find_elements(By.CSS_SELECTOR, selector)
                    modales_actuales.extend([e for e in elementos if e.is_displayed()])
                
                if not modales_actuales:
                    logging.info("✅ [MODAL DETECTOR] ¡Modal cerrado exitosamente!")
                    return True
                else:
                    logging.warning(f"⚠️  Modal sigue visible (intento {intento + 1}/{max_intentos})")
                    time.sleep(1)
                    continue
                    
            except Exception as e:
                logging.warning(f"⚠️  Error verificando modal: {e}")
                return True  # Asumimos que desapareció
        
        # ═══ SI LLEGAMOS AQUÍ, NO SE PUDO CERRAR ═══
        logging.error("❌ [MODAL DETECTOR] Modal no respondió a ninguna estrategia")
        logging.warning("🔄 FORZANDO RECARGA DE PÁGINA Y LIMPIEZA...")
        
        try:
            driver.refresh()
            time.sleep(5)
            limpiar_todos_los_campos(driver)
            time.sleep(1)
            logging.info("✅ Página recargada y campos limpiados")
        except Exception as e:
            logging.error(f"❌ Error en recarga: {e}")
        
        return False  # Retorna False indicando que se recargó
        
    except Exception as e:
        logging.error(f"❌ [MODAL DETECTOR] Error grave: {e}")
        logging.warning("🔄 Intentando recarga de seguridad...")
        try:
            driver.refresh()
            time.sleep(5)
        except:
            pass
        return False



# ═════════════════════════════════════════════════════════════
# PROCESAMIENTO DE CONSULTA - ⭐ MEJORADO
# ═════════════════════════════════════════════════════════════

def procesar_consulta(driver, cedula_asociado, cedula_propietario, placa, fila_numero):
    """
    Procesa una consulta con reintentos:
    1. INTENTO 1: Con cedula_asociado
    2. Si falla "no hay personas": INTENTO 2 con cedula_propietario
    3. Siempre guarda en "Datos Runt" con cedula_asociado
    4. Guarda resultado en "Resultados" con ambas cédulas
    """
    
    logging.info(f"\n{'='*70}")
    logging.info(f"🔍 INTENTO 1: Usando CÉDULA ASOCIADO")
    logging.info(f"{'='*70}")
    
    resultado, _ = procesar_consulta_interno(
        driver, 
        cedula_asociado,
        placa, 
        fila_numero,
        es_reintento=False
    )
    
    if resultado and resultado.get("estado") == "Exitoso":
        logging.info(f"✅ INTENTO 1 EXITOSO con cédula asociado")
        
        guardar_resultado_en_resultados(
            cedula_asociado, 
            cedula_propietario, 
            placa, 
            cedula_asociado,
            "Exitoso"
        )
        
        return resultado, fila_numero
    
    elif resultado and resultado.get("estado") == "Exitoso - Sin personas asociadas":
        logging.warning(f"⚠️ INTENTO 1: No hay personas asociadas")
        logging.info(f"🔄 INICIANDO INTENTO 2: Usando CÉDULA PROPIETARIO")
        
        logging.info(f"\n{'='*70}")
        logging.info(f"🔍 INTENTO 2: Usando CÉDULA PROPIETARIO")
        logging.info(f"{'='*70}")
        
        resultado_reintento, _ = procesar_consulta_interno(
            driver, 
            cedula_propietario,
            placa, 
            fila_numero,
            es_reintento=True
        )
        
        if resultado_reintento and resultado_reintento.get("estado") == "Exitoso":
            logging.info(f"✅ INTENTO 2 EXITOSO con cédula propietario")
            
            resultado_reintento["cedula"] = cedula_asociado
            
            guardar_resultado_en_resultados(
                cedula_asociado, 
                cedula_propietario, 
                placa, 
                cedula_propietario,
                "Exitoso"
            )
            
            return resultado_reintento, fila_numero
        else:
            logging.error(f"❌ INTENTO 2 TAMBIÉN FALLÓ")
            
            guardar_resultado_en_resultados(
                cedula_asociado, 
                cedula_propietario, 
                placa, 
                cedula_asociado,
                "Falló"
            )
            
            return None, fila_numero
    else:
        logging.error(f"❌ INTENTO 1 FALLÓ")
        
        guardar_resultado_en_resultados(
            cedula_asociado, 
            cedula_propietario, 
            placa, 
            cedula_asociado,
            "Falló"
        )
        
        return None, fila_numero


def procesar_consulta_interno(driver, cedula, placa, fila_numero, es_reintento=False):
    """
    Lógica interna de procesamiento
    """
    max_reintentos = 2
    
    for reintento_general in range(max_reintentos):
        try:
            logging.info(f"\n{'='*70}")
            logging.info(f"🔍 Procesando consulta [{reintento_general + 1}/{max_reintentos}]")
            logging.info(f"Cédula: {cedula}, Placa: {placa}")
            logging.info(f"{'='*70}")
            
            if reintento_general > 0:
                logging.info(f"🔄 REINTENTANDO CON LOS MISMOS DATOS (Intento {reintento_general + 1}/{max_reintentos})")
                driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                time.sleep(4)
                limpiar_campos_individuales_validado(driver, cedula, placa)
                time.sleep(1)
            else:
                driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                time.sleep(4)
                limpiar_todos_los_campos(driver)
                time.sleep(1)

            # ═══ LLENAR PLACA ═══
            try:
                placa_input = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[formcontrolname="placa"]'))
                )
                logging.info("📝 Llenando placa...")
                
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", placa_input)
                time.sleep(0.5)
                
                limpiar_campo_input(driver, placa_input, "PLACA")
                time.sleep(0.5)
                
                for char in str(placa).strip():
                    placa_input.send_keys(char)
                    time.sleep(0.1)
                
                logging.info(f"✅ Placa escrita: {placa}")
                time.sleep(0.5)
            except Exception as e:
                logging.error(f"❌ Error al llenar placa: {e}")
                continue

            # ═══ LLENAR CÉDULA ═══
            try:
                cedula_input = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[formcontrolname="documento"]'))
                )
                logging.info("📝 Llenando cédula...")
                
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cedula_input)
                time.sleep(0.5)
                
                limpiar_campo_input(driver, cedula_input, "CÉDULA")
                time.sleep(0.5)
                
                for char in str(cedula).strip():
                    cedula_input.send_keys(char)
                    time.sleep(0.1)
                
                logging.info(f"✅ Cédula escrita: {cedula}")
                time.sleep(1)
            except Exception as e:
                logging.error(f"❌ Error al llenar cédula: {e}")
                continue

            logging.info("📜 Scrolling para ver el captcha...")
            driver.execute_script("window.scrollBy(0, 300);")
            time.sleep(2)

            # ═══ RESOLVER Y ESCRIBIR CAPTCHA ═══
            for intento_captcha in range(2):
                logging.info(f"\n🔐 Intento {intento_captcha + 1} de resolver captcha...")
                
                captcha_path, captcha_img = capturar_captcha(driver, placa)
                if not captcha_img:
                    logging.warning("⚠️ No se pudo capturar el captcha")
                    time.sleep(2)
                    continue

                texto_final = resolver_captcha(captcha_img, placa)

                if texto_final:
                    logging.info(f"📝 Captcha final a escribir: {texto_final}")
                    
                    try:
                        captcha_inputs = driver.find_elements(By.TAG_NAME, "input")
                        captcha_field = None
                        
                        for inp in captcha_inputs:
                            if inp.is_displayed():
                                captcha_field = inp
                        
                        if not captcha_field:
                            logging.error("❌ No se encontró el campo de captcha")
                            continue
                        
                        logging.info("🧹 Limpiando campo de captcha...")
                        limpiar_campo_input(driver, captcha_field, "CAPTCHA")
                        time.sleep(0.5)
                        
                        logging.info(f"⌨️ Escribiendo: {texto_final}")
                        captcha_logger.info(f"         ESCRITO: {texto_final}")
                        
                        for char in texto_final:
                            captcha_field.send_keys(char)
                            time.sleep(0.2)
                        
                        logging.info("✅ Captcha escrito")
                        time.sleep(1)
                        
                        try:
                            boton_consultar = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Consultar')]"))
                            )
                            logging.info("🔘 Clickeando botón 'Consultar'...")
                            boton_consultar.click()

                            logging.info("⏱️ Esperando respuesta del servidor (máx. 4s)...")
                            error_detectado = None
                            for s in range(4):
                                time.sleep(1)
                                logging.info(f"   ({s+1}/4s) Verificando...")
                                error_temp = detectar_mensaje_error(driver)
                                if error_temp:
                                    error_detectado = error_temp
                                    logging.info(f"   ⚠️ Error detectado: {error_detectado}")
                                    break

                            logging.info("🏁 Fin de la espera de respuesta.")
                            
                        except Exception as e:
                            logging.error(f"❌ Error al clickear Consultar: {e}")
                            continue

                        # ═══ MANEJO DE MODAL Y RESPUESTA ═══
                        time.sleep(2)
                        
                        resultado_cierre = detectar_y_cerrar_modal_universal(driver, timeout_max=3, max_intentos=4)

                        if not resultado_cierre:
                            logging.warning("⚠️ Página fue recargada. Rompiendo ciclo de captcha...")
                            break
                        else:
                            if error_detectado == "captcha_incorrecto":
                                logging.warning("❌ Captcha rechazado, reintentando...")
                                continue
                            elif error_detectado == "no_personas":
                                logging.warning("ℹ️ No hay personas asociadas a este vehículo")
                                
                                resultado = {
                                    "Tiempo ejecucion": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                    "cedula": cedula,
                                    "placa": placa,
                                    "cilindraje": "No disponible",
                                    "estado": "Exitoso - Sin personas asociadas",
                                    "datos_soat": ["No disponible"] * 7,
                                    "datos_técnicos": ["No disponible"] * 7
                                }
                                agregar_registro_procesado(cedula, placa, "Exitoso - Sin personas", None, ["No disponible"] * 7, ["No disponible"] * 7)
                                return resultado, fila_numero

                            elif error_detectado == "error_desconocido":
                                logging.warning("⚠️ Error desconocido detectado")
                                continue
                            else:
                                logging.info("✅ ¡CAPTCHA ACEPTADO!")

                                try:
                                    timestamp = datetime.now().strftime('%H%M%S')
                                    nombre_exitoso = f"exitoso_{placa}_{texto_final}_{timestamp}.png"
                                    ruta_exitoso = CAPTCHA_LEIDOS_FOLDER / nombre_exitoso
                                    img_pil = Image.open(captcha_path)
                                    img_pil.save(str(ruta_exitoso))
                                    logging.info(f"✅ Captcha exitoso guardado")
                                except:
                                    pass
                                
                                # PASO 1: Extraer DATOS DEL VEHÍCULO
                                logging.info("\n" + "="*70)
                                logging.info("📊 PASO 1: LEYENDO DATOS DEL VEHÍCULO...")
                                logging.info("="*70)
                                datos_vehiculo = extraer_datos_vehiculo_optimizado(driver)
                                time.sleep(2)
                                
                                fila_vehiculo = escribir_datos_vehiculo_en_sheets(
                                    datos_vehiculo, cedula, placa
                                )
                                
                                if fila_vehiculo:
                                    logging.info(f"✅ Datos del vehículo guardados en fila {fila_vehiculo}")
                                
                                time.sleep(2)

                                # PASO 2: Extraer SOAT
                                logging.info("\n" + "="*70)
                                logging.info("📊 PASO 2: LEYENDO DATOS DE SOAT...")
                                logging.info("="*70)
                                soat_datos = extraer_datos_soat(driver)
                                time.sleep(2)

                                # PASO 3: Extraer RTM
                                logging.info("\n" + "="*70)
                                logging.info("📊 PASO 3: LEYENDO DATOS DE TÉCNOMECÁNICA...")
                                logging.info("="*70)
                                rtm_datos = extraer_datos_rtm(driver)
                                time.sleep(2)

                                logging.info("📜 Scrolling al inicio...")
                                driver.execute_script("window.scrollTo(0, 0);")
                                time.sleep(1)

                                resultado = {
                                    "Tiempo ejecucion": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                    "cedula": cedula,
                                    "placa": placa,
                                    "cilindraje": datos_vehiculo.get("CILINDRAJE", "No disponible"),
                                    "estado": "Exitoso",
                                    "datos_vehiculo": datos_vehiculo,
                                    "datos_soat": soat_datos,
                                    "datos_técnicos": rtm_datos
                                }

                                agregar_registro_procesado(cedula, placa, "Exitoso", datos_vehiculo, soat_datos, rtm_datos)
                                
                                return resultado, fila_numero

                    except Exception as e:
                        logging.error(f"❌ Error escribiendo captcha: {e}")
                        captcha_logger.info(f"         RESULTADO: ❌ ERROR")
                        continue
                else:
                    logging.warning(f"⚠️ No se resolvió el captcha en intento {intento_captcha + 1}")

        except Exception as e:
            logging.error(f"❌ Error en reintento {reintento_general + 1}: {e}")
            continue
    
    logging.error(f"❌ No se completó la consulta para {placa} tras {max_reintentos} intentos")
    agregar_registro_procesado(cedula, placa, "Falló - Error técnico")
    return None, fila_numero

# ═════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════

def main():
    """Función principal CON SOPORTE DE REANUDACIÓN MEJORADO"""
    driver = iniciar_driver()
    if not driver:
        logging.error("❌ No se pudo inicializar el driver")
        return

    driver.maximize_window()

    try:
        # ═══ CARGAR DATOS Y ESTADO ANTERIOR ═══
        datos_unicos = obtener_datos_unicos()
        estado_anterior = cargar_estado()
        
        last_cedula = estado_anterior.get("last_cedula")
        last_placa = estado_anterior.get("last_placa")
        last_status = estado_anterior.get("last_status")
        
        logging.info(f"\n{'='*70}")
        logging.info(f"📋 ESTADO ANTERIOR CARGADO")
        logging.info(f"   Última placa: {last_placa}")
        logging.info(f"   Última cédula: {last_cedula}")
        logging.info(f"   Último estado: {last_status}")
        logging.info(f"{'='*70}\n")
        
        if not datos_unicos:
            logging.info(f"\n{'='*70}")
            logging.info(f"✅ NO HAY REGISTROS PENDIENTES POR PROCESAR")
            logging.info(f"{'='*70}\n")
            captcha_logger.info(f"✅ NO HAY REGISTROS PENDIENTES")
            return
        
        # ═══ LÓGICA DE REANUDACIÓN ═══
        inicio_desde = 0
        
        if last_placa and last_status:
            for idx, (ced_asoc, ced_prop, plac, fil, sheet_name) in enumerate(datos_unicos):
                if plac == last_placa and ced_asoc == last_cedula:
                    if last_status == "Exitoso":
                        inicio_desde = idx + 1
                        logging.info(f"✅ Retomando desde índice {inicio_desde} (después de {last_placa})")
                    elif last_status in ["Pendiente", "Error", "Falló"]:
                        inicio_desde = idx
                        logging.info(f"🔄 Reintentando desde índice {inicio_desde} ({last_placa})")
                    else:
                        inicio_desde = idx
                    break
        
        logging.info(f"\n{'='*70}")
        logging.info(f"🚀 Iniciando proceso para {len(datos_unicos)} registros PENDIENTES")
        logging.info(f"   Comenzando desde: índice {inicio_desde}")
        logging.info(f"{'='*70}\n")
        
        captcha_logger.info(f"{'='*70}")
        captcha_logger.info(f"🚀 INICIANDO - {len(datos_unicos)} REGISTROS PENDIENTES")
        captcha_logger.info(f"   Desde índice: {inicio_desde}")
        captcha_logger.info(f"{'='*70}\n")

        # ═══ PROCESAR REGISTROS ═══
        for i in range(inicio_desde, len(datos_unicos)):
            cedula_asociado, cedula_propietario, placa, fila_numero, sheet_origen = datos_unicos[i]
            
            logging.info(f"\n{'='*70}")
            logging.info(f"📊 Procesando [{i + 1}/{len(datos_unicos)}]: Placa {placa} (desde {sheet_origen})")
            logging.info(f"{'='*70}")
            
            resultado, fila = procesar_consulta(
                driver, 
                cedula_asociado,
                cedula_propietario,
                placa, 
                fila_numero
            )

            if resultado:
                estado_resultado = resultado.get("estado", "Pendiente")
                guardar_estado(
                    cedula_asociado, 
                    placa, 
                    estado_resultado, 
                    i, 
                    len(datos_unicos),
                    resultado.get("datos_vehiculo"),  # NUEVO
                    resultado.get("datos_soat"),
                    resultado.get("datos_técnicos")
                )
                
                # Guardar en hoja "Datos Runt"
                guardar_en_sheets([resultado])
                
                # ==================== NUEVO: Guardar en hoja "Datos Vehiculo" ====================
                # INSERTA ESTAS LÍNEAS DESPUÉS DE guardar_en_sheets
                
                if "datos_vehiculo" in resultado and resultado["datos_vehiculo"]:
                    escribir_datos_vehiculo_sheets(resultado["datos_vehiculo"])
                    logging.info(f"✅ Datos completos de {placa} guardados en 'Datos Vehiculo'")
                # ==================== FIN NUEVO ====================
                
                logging.info(f"✅ Datos de {placa} guardados en Sheets")

            # ═══ PREPARAR SIGUIENTE CONSULTA ═══
            if i < len(datos_unicos) - 1:
                try:
                    logging.info("🔄 Preparando siguiente consulta...")
                    limpiar_todos_los_campos(driver)
                    time.sleep(1)
                    
                    # Intentar con diferentes selectores
                    try:
                        otra_consulta_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Otra consulta')]"))
                        )
                    except:
                        try:
                            otra_consulta_btn = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.mat-raised-button"))
                            )
                        except:
                            logging.warning("⚠️ No se encontró botón 'Otra consulta', recargando página...")
                            driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                            time.sleep(3)
                            limpiar_todos_los_campos(driver)
                            continue
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", otra_consulta_btn)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", otra_consulta_btn)
                    time.sleep(3)
                    limpiar_todos_los_campos(driver)
                    time.sleep(1)
                    
                except Exception as e:
                    logging.error(f"⚠️ Error al siguiente: {e}")
                    logging.info("🔄 Recargando página...")
                    driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                    time.sleep(3)
                    limpiar_todos_los_campos(driver)
                    time.sleep(1)

        logging.info(f"\n{'='*70}")
        logging.info(f"✅ PROCESO FINALIZADO - TODOS LOS REGISTROS PROCESADOS")
        logging.info(f"{'='*70}\n")
        
        captcha_logger.info(f"\n{'='*70}")
        captcha_logger.info(f"✅ PROCESO FINALIZADO")
        captcha_logger.info(f"{'='*70}\n")

    except Exception as e:
        logging.error(f"❌ Error principal: {e}", exc_info=True)
    finally:
        cerrar_driver(driver)

if __name__ == "__main__":
    logging.info("🚀 Iniciando Bot RUNT FINAL...")
    main()