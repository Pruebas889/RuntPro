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

LOGS_FOLDER = BASE_PATH / "Escritura_Runt_principal"
LOGS_FOLDER.mkdir(parents=True, exist_ok=True)

TESSERACT_PATH = r"C:\Users\cmarroquin\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

GOOGLE_CREDS = BASE_PATH / "prueba-de-gmail-486215-345473339c47.json"
ESTADO_FILE = str(LOGS_FOLDER / "estado_runt.json")


REGISTROS_ANTES_REINICIO = 6  # Reiniciar sesión cada 6 registros

# ═════════════════════════════════════════════════════════════
# LOGGING
# ═════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(str(LOGS_FOLDER / "automatizacion.log"), encoding='utf-8')
    ]
)

captcha_logger = logging.getLogger('captcha_logger')
captcha_logger.setLevel(logging.INFO)
captcha_logger.propagate = False
file_handler_captcha = logging.FileHandler(str(LOGS_FOLDER / "captcha_retroalimentacion.log"), encoding='utf-8')
formatter_captcha = logging.Formatter('%(message)s')
file_handler_captcha.setFormatter(formatter_captcha)
captcha_logger.addHandler(file_handler_captcha)


def reiniciar_sesion_periodico(driver, contador):
    """
    Reinicia la sesión cada 5 registros procesados
    VERSIÓN AGRESIVA: Entra a otra URL y vuelve
    """
    
    if contador > 0 and contador % 5 == 0:
        logging.info(f"\n{'='*70}")
        logging.info(f"⏸️  PAUSA PROGRAMADA: {contador} registros procesados")
        logging.info(f"{'='*70}")
        logging.info(f"⏱️  Esperando 10 segundos antes de refrescar...")
        
        for segundo in range(10, 0, -1):
            if segundo % 5 == 0 or segundo <= 3:
                logging.info(f"   ⏳ {segundo} segundos restantes...")
            time.sleep(1)
        
        logging.info(f"🔄 RECARGANDO PÁGINA (CAMBIO DE URL VISIBLE)...")
        try:
            # Paso 1: Ir a Google (visible)
            logging.info(f"   → Yendo a Google...")
            driver.get("https://www.google.com")
            time.sleep(2)
            
            # Paso 2: Volver a RUNT
            logging.info(f"   → Volviendo a RUNT...")
            driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
            time.sleep(5)
            
            # Paso 3: Limpiar
            logging.info(f"   → Limpiando campos...")
            limpiar_todos_los_campos(driver)
            time.sleep(2)
            
            logging.info(f"✅ SESIÓN REINICIADA - Listo para continuar")
            logging.info(f"{'='*70}\n")
            return True
            
        except Exception as e:
            logging.error(f"❌ Error en reinicio: {e}")
            logging.info(f"{'='*70}\n")
            return False
    
    return False

# ═══════════════════════════��═════════════════════════════════
# MANEJO DE ESTADO (REANUDACIÓN)
# ═════════════════════════════════════════════════════════════

def estructura_estado_inicial():
    """Crea la estructura inicial del archivo de estado - SIMPLIFICADO"""
    return {
        "resumen": {
            "placas_procesadas": {},  # {"PLACA001": "Exitoso", "PLACA002": "Pendiente"}
            "ultima_actualizacion": datetime.now().isoformat()
        },
        "historial_completo": [],  # Mantiene todos los detalles históricos
        "last_execution": None,
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
    """Guarda el estado actual - Actualiza RESUMEN e HISTORIAL"""
    estado_anterior = cargar_estado()
    
    # Normalizar status a Exitoso/Pendiente para el resumen
    estado_resumen = "Exitoso" if "Exitoso" in status else "Pendiente"
    
    # Actualizar RESUMEN con la placa y estado
    estado_anterior["resumen"]["placas_procesadas"][placa] = estado_resumen
    estado_anterior["resumen"]["ultima_actualizacion"] = datetime.now().isoformat()
    
    # Agregar al HISTORIAL COMPLETO
    registro_actual = {
        "timestamp": datetime.now().isoformat(),
        "cedula": cedula,
        "placa": placa,
        "status": status,
        "datos_vehiculo": datos_vehiculo or {},
        "datos_soat": datos_soat or [],
        "datos_tecnica": datos_tecnica or []
    }
    
    estado_anterior["historial_completo"].append(registro_actual)
    
    # Actualizar índices
    estado_anterior["last_execution"] = datetime.now().isoformat()
    estado_anterior["current_index"] = index
    estado_anterior["total_records"] = total
    
    # Guardar en archivo
    with open(ESTADO_FILE, "w", encoding="utf-8") as f:
        json.dump(estado_anterior, f, indent=2, ensure_ascii=False)
    
    total_historial = len(estado_anterior["historial_completo"])
    logging.info(f"💾 Estado guardado: {placa} - {estado_resumen} (Total historial: {total_historial} registros)")



def agregar_registro_procesado(cedula, placa, status, datos_vehiculo=None, datos_soat=None, datos_tecnica=None):
    """Agrega un registro a la lista de procesados - Mantiene RESUMEN + HISTORIAL"""
    estado = cargar_estado()
    
    # Normalizar status a Exitoso/Pendiente para el resumen
    estado_resumen = "Exitoso" if "Exitoso" in status else "Pendiente"
    
    # Actualizar RESUMEN
    if "resumen" not in estado:
        estado["resumen"] = {"placas_procesadas": {}, "ultima_actualizacion": None}
    
    estado["resumen"]["placas_procesadas"][placa] = estado_resumen
    estado["resumen"]["ultima_actualizacion"] = datetime.now().isoformat()
    
    # Agregar al HISTORIAL COMPLETO
    registro = {
        "timestamp": datetime.now().isoformat(),
        "cedula": cedula,
        "placa": placa,
        "status": status,
        "datos_vehiculo": datos_vehiculo or {},
        "datos_soat": datos_soat or [],
        "datos_tecnica": datos_tecnica or []
    }
    
    if "historial_completo" not in estado:
        estado["historial_completo"] = []
    
    estado["historial_completo"].append(registro)
    
    # Guardar
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
    - Columna B (índice 1): Cédula asociado
    - Columna D (índice 3): Cédula propietario  
    - Columna F (índice 5): Placa
    
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
        # vistos = set()
        
        estadisticas_sheets = {}
        total_recopilado = 0
        
        for nombre_sheet in nombres_sheets:
            try:
                worksheet = sheet.worksheet(nombre_sheet)
                logging.info(f"📊 Leyendo sheet: {nombre_sheet}")
                
                # ═══ OBTENER TODAS LAS FILAS ═══
                todas_filas = worksheet.get_all_values()
                logging.info(f"   📍 Total de filas en sheet: {len(todas_filas)}")
                
                if len(todas_filas) <= 1:
                    logging.warning(f"   ⚠️  Sheet '{nombre_sheet}' vacío o solo tiene encabezados")
                    estadisticas_sheets[nombre_sheet] = 0
                    continue
                
                # ═══ DIAGNÓSTICO: MOSTRAR PRIMERAS FILAS ═══
                logging.info(f"   🔍 DIAGNÓSTICO - Primeras 3 filas del sheet:")
                for row_idx, fila in enumerate(todas_filas[:3]):
                    logging.info(f"      Fila {row_idx}: {fila}")
                
                # ═══ LEER COLUMNAS CON VALIDACIÓN ═══
                cedulas_asociado = worksheet.col_values(2)   # Columna B (índice 1)
                cedulas_propietario = worksheet.col_values(4) # Columna D (índice 3)
                placas = worksheet.col_values(6)              # Columna F (índice 5)
                
                logging.info(f"   📋 Columna B (Cédula Asociado): {cedulas_asociado[:3]}")
                logging.info(f"   📋 Columna D (Cédula Propietario): {cedulas_propietario[:3]}")
                logging.info(f"   📋 Columna F (Placa): {placas[:3]}")
                
                registros_sheet = 0
                
                # ═══ PROCESAR FILAS (saltando encabezado en fila 0) ═══
                min_len = min(len(cedulas_asociado), len(cedulas_propietario), len(placas))
                
                for i in range(1, min_len):  # Comienza en 1 para saltar encabezado
                    cedula_asoc = str(cedulas_asociado[i]).strip()
                    cedula_prop = str(cedulas_propietario[i]).strip()
                    placa = str(placas[i]).strip()
                    
                    # ═══ VALIDACIÓN BÁSICA ═══
                    if not placa or placa.lower() == "nan":
                        continue
                    
                    if not cedula_asoc or cedula_asoc.lower() == "nan":
                        continue
                    
                    
                    # ═══ PERMITIR DUPLICADOS ENTRE HOJAS ═══
                    datos.append((cedula_asoc, cedula_prop, placa, i + 1, nombre_sheet))
                    registros_sheet += 1
                    
                    if registros_sheet <= 3:
                        logging.info(f"      ✓ Fila {i+1}: {cedula_asoc} | {cedula_prop} | {placa}")
                
                estadisticas_sheets[nombre_sheet] = registros_sheet
                total_recopilado += registros_sheet
                
                logging.info(f"   ✅ {registros_sheet} registros únicos de {nombre_sheet}")
                
            except gspread.WorksheetNotFound:
                logging.warning(f"⚠️  Sheet '{nombre_sheet}' no encontrada")
                estadisticas_sheets[nombre_sheet] = 0
                continue
            except Exception as e:
                logging.error(f"❌ Error leyendo sheet '{nombre_sheet}': {e}")
                estadisticas_sheets[nombre_sheet] = 0
                continue
        
        # ═════════════════════════════════════════════════════════════
        # MOSTRAR ESTADÍSTICAS RESUMIDAS
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
        logging.error(f"❌ Error obteniendo datos de Sheets: {e}", exc_info=True)
        return []


def validar_registro(cedula_asoc, cedula_prop, placa):
    """
    Valida que los datos NO contengan 'sin información' o similares
    Retorna: True si es válido, False si debe ignorarse
    """
    # Palabras clave que indican datos REALMENTE inválidos
    palabras_invalidas = [
        "sin información", 
        "sin info",
        "n/a",
        "na",
        "null",
    ]
    
    # Convertir a strings
    cedula_asoc_str = str(cedula_asoc).strip().lower()
    cedula_prop_str = str(cedula_prop).strip().lower()
    placa_str = str(placa).strip().upper()
    
    # ❌ RECHAZAR SOLO si está LITERALMENTE vacío
    if not cedula_asoc_str or cedula_asoc_str == "nan":
        logging.warning(f"   ❌ Rechazado: Cédula asociado vacía")
        return False
    
    if not placa_str or placa_str == "NAN":
        logging.warning(f"   ❌ Rechazado: Placa vacía")
        return False
    
    # ❌ RECHAZAR solo si contiene las palabras clave de "sin información"
    for palabra_invalida in palabras_invalidas:
        if palabra_invalida in cedula_asoc_str:
            logging.warning(f"   ❌ Rechazado: Cédula contiene '{palabra_invalida}'")
            return False
        if palabra_invalida in placa_str.lower():
            logging.warning(f"   ❌ Rechazado: Placa contiene '{palabra_invalida}'")
            return False
    
    # ✅ TODO DEMÁS ES VÁLIDO
    return True

def filtrar_datos_validos(datos_unicos):
    """
    Filtra los datos y retorna estadísticas
    """
    logging.info("\n" + "="*70)
    logging.info("🔍 VALIDANDO DATOS DESDE GOOGLE SHEETS...")
    logging.info("="*70)
    
    datos_validos = []
    datos_invalidos = []
    
    for cedula_asoc, cedula_prop, placa, fila, sheet_name in datos_unicos:
        if validar_registro(cedula_asoc, cedula_prop, placa):
            datos_validos.append((cedula_asoc, cedula_prop, placa, fila, sheet_name))
        else:
            datos_invalidos.append({
                "cedula_asoc": cedula_asoc,
                "cedula_prop": cedula_prop,
                "placa": placa,
                "sheet": sheet_name,
                "razon": "Datos incompletos o inválidos"
            })
    
    # ═══ MOSTRAR ESTADÍSTICAS ═══
    logging.info("\n" + "="*70)
    logging.info("📊 VALIDACIÓN DE DATOS COMPLETADA")
    logging.info("="*70)
    logging.info(f"✅ Datos válidos: {len(datos_validos)} registros")
    logging.info(f"❌ Datos inválidos: {len(datos_invalidos)} registros")
    logging.info("="*70 + "\n")
    
    if datos_invalidos:
        logging.warning("⚠️  Registros ignorados por datos incompletos:")
        for inv in datos_invalidos[:10]:  # Mostrar máximo 10
            logging.warning(f"   - Placa: {inv['placa']}, Cédula: {inv['cedula_asoc']}, Sheet: {inv['sheet']}")
        if len(datos_invalidos) > 10:
            logging.warning(f"   ... y {len(datos_invalidos) - 10} más")
    
    return datos_validos, datos_invalidos




# ═════════════════════════════════════════════════════════════
# ⭐ NUEVA FUNCIÓN: OBTENER ÍNDICE DESDE ÚLTIMO EXITOSO
# ═════════════════════════════════════════════════════════════

def obtener_indice_reanudacion_desde_ultimo_exitoso(datos_validos, placas_procesadas):
    """
    🎯 Encuentra el ÚLTIMO registro con "Exitoso" en el JSON
    y retorna el índice del SIGUIENTE registro por procesar
    
    Returns:
        (indice_inicio, ultima_placa_exitosa, total_exitosos_anteriores)
    """
    
    if not placas_procesadas:
        logging.info("📊 Estado vacío - Iniciando desde el primer registro (índice 0)")
        return 0, None, 0
    
    # 🔍 BUSCAR EN ORDEN EL ÚLTIMO QUE DICE "EXITOSO"
    ultima_exitosa_encontrada = None
    ultima_posicion_exitosa = -1
    total_exitosos = 0
    
    logging.info("\n🔍 Buscando el ÚLTIMO registro 'EXITOSO' en JSON...")
    logging.info(f"   Total de placas en JSON: {len(placas_procesadas)}")
    
    # Iterar datos_validos en ORDEN
    for idx, (cedula_a, cedula_p, placa, fila, sheet) in enumerate(datos_validos):
        estado = placas_procesadas.get(placa)
        
        # Si encuentra "Exitoso", lo guarda (sobrescribe cada vez que encuentra uno más nuevo)
        if estado == "Exitoso":
            ultima_exitosa_encontrada = placa
            ultima_posicion_exitosa = idx
            total_exitosos += 1
            logging.info(f"   ✓ Exitoso #{total_exitosos}: {placa} (posición {idx})")
    
    if ultima_exitosa_encontrada is None:
        logging.warning("⚠️  No se encontró ningún 'Exitoso' en el JSON")
        logging.warning("   → Iniciando desde el primer registro (índice 0)")
        return 0, None, 0
    
    # El siguiente índice es el que viene después del último exitoso
    proximo_indice = ultima_posicion_exitosa + 1
    
    logging.info("\n" + "="*70)
    logging.info(f"✅ ÚLTIMO EXITOSO ENCONTRADO: {ultima_exitosa_encontrada}")
    logging.info(f"   Posición en lista: {ultima_posicion_exitosa}")
    logging.info(f"   Total exitosos antes: {total_exitosos}")
    logging.info(f"   PRÓXIMO POR PROCESAR → Índice {proximo_indice}")
    logging.info("="*70 + "\n")
    
    return proximo_indice, ultima_exitosa_encontrada, total_exitosos


# ═════════════════════════════════════════════════════════════
# ⭐ NUEVA FUNCIÓN: REINTENTOS FINALES DE PENDIENTES
# ═════════════════════════════════════════════════════════════

def procesar_reintentos_finales_pendientes(driver, datos_validos, max_intentos=5):
    """
    🔴 AL FINAL: Reintenta SOLO los "Pendiente" hasta {max_intentos} veces
    para confirmar si es error de captcha o "sin personas"
    
    NO toca Exitoso ni Sin_Personas
    """
    
    estado = cargar_estado()
    placas_procesadas = estado.get("resumen", {}).get("placas_procesadas", {})
    
    # Obtener SOLO los Pendiente
    pendientes = {placa: est for placa, est in placas_procesadas.items() if est == "Pendiente"}
    
    if not pendientes:
        logging.info(f"\n✅ No hay registros 'Pendiente' por reintentar")
        return None
    
    logging.info(f"\n{'='*70}")
    logging.info(f"🔴 FASE FINAL: REINTENTANDO {len(pendientes)} PENDIENTES")
    logging.info(f"Max intentos por registro: {max_intentos}")
    logging.info(f"{'='*70}\n")
    
    resultados_finales = {
        "recuperados": [],
        "sin_personas_confirmado": [],
        "fallidos_permanentes": []
    }
    
    for idx_registro, placa in enumerate(sorted(pendientes.keys()), 1):
        logging.info(f"\n{'='*70}")
        logging.info(f"🔴 [{idx_registro}/{len(pendientes)}] Reintentando: {placa}")
        logging.info(f"{'='*70}")
        
        # Encontrar datos de esta placa
        datos_encontrados = None
        for cedula_a, cedula_p, plac, fila, sheet in datos_validos:
            if plac == placa:
                datos_encontrados = (cedula_a, cedula_p, plac, fila)
                break
        
        if not datos_encontrados:
            logging.warning(f"⚠️  No se encontraron datos para {placa}, saltando...")
            continue
        
        cedula_a, cedula_p, placa_found, fila = datos_encontrados
        
        # 🔄 REINTENTOS
        exitoso_en_reintento = False
        sin_personas_confirmado = False
        
        for intento_actual in range(1, max_intentos + 1):
            logging.info(f"\n🔄 Intento {intento_actual}/{max_intentos}")
            
            # Intentar con asociado
            logging.info(f"   → Probando con cédula ASOCIADO...")
            resultado, _ = procesar_consulta_interno(
                driver,
                cedula_a,
                placa,
                fila,
                es_reintento=True,
                max_intentos_internos=1
            )
            
            if resultado and resultado.get("estado") == "Exitoso":
                logging.info(f"   ✅ ¡ÉXITO CON ASOCIADO!")
                exitoso_en_reintento = True
                
                # Actualizar JSON de Pendiente → Exitoso
                guardar_estado(cedula_a, placa, "Exitoso", idx_registro, len(pendientes),
                              resultado.get("datos_vehiculo"), resultado.get("datos_soat"), 
                              resultado.get("datos_técnicos"))
                
                # Guardar en sheets (ACTUALIZAR fila existente)
                guardar_en_sheets([resultado], actualizar_existente=True)
                escribir_datos_vehiculo_en_sheets(resultado.get("datos_vehiculo", {}), cedula_a, placa)
                guardar_resultado_en_resultados(cedula_a, cedula_p, placa, cedula_a, "Exitoso")
                
                resultados_finales["recuperados"].append({
                    "placa": placa,
                    "cedula_usada": cedula_a,
                    "intento": intento_actual
                })
                break
            
            elif resultado and resultado.get("estado") == "Exitoso - Sin personas asociadas":
                logging.warning(f"   ⚠️  Sin personas con ASOCIADO, probando PROPIETARIO...")
                
                # Intentar con propietario
                resultado_prop, _ = procesar_consulta_interno(
                    driver,
                    cedula_p,
                    placa,
                    fila,
                    es_reintento=True,
                    max_intentos_internos=1
                )
                
                if resultado_prop and resultado_prop.get("estado") == "Exitoso":
                    logging.info(f"   ✅ ¡ÉXITO CON PROPIETARIO!")
                    exitoso_en_reintento = True
                    
                    guardar_estado(cedula_a, placa, "Exitoso", idx_registro, len(pendientes),
                                  resultado_prop.get("datos_vehiculo"), resultado_prop.get("datos_soat"),
                                  resultado_prop.get("datos_técnicos"))
                    
                    guardar_en_sheets([resultado_prop], actualizar_existente=True)
                    escribir_datos_vehiculo_en_sheets(resultado_prop.get("datos_vehiculo", {}), cedula_a, placa)
                    guardar_resultado_en_resultados(cedula_a, cedula_p, placa, cedula_p, "Exitoso")
                    
                    resultados_finales["recuperados"].append({
                        "placa": placa,
                        "cedula_usada": cedula_p,
                        "intento": intento_actual
                    })
                    break
                else:
                    # SIN PERSONAS EN AMBAS
                    logging.error(f"   ❌ Sin personas en AMBAS cédulas (confirmado)")
                    sin_personas_confirmado = True
                    break
            else:
                # Error técnico
                logging.warning(f"   ⚠️  Error técnico en intento {intento_actual}")
                if intento_actual < max_intentos:
                    logging.info(f"   ⏳ Esperando 5 segundos...")
                    time.sleep(5)
                    continue
        
        # ═══ GUARDAR RESULTADO FINAL ═══
        if exitoso_en_reintento:
            logging.info(f"✅ {placa}: RECUPERADO")
        
        elif sin_personas_confirmado:
            logging.error(f"❌ {placa}: Sin personas (CONFIRMADO)")
            
            # Cambiar estado de Pendiente → Sin_Personas
            estado = cargar_estado()
            estado["resumen"]["placas_procesadas"][placa] = "Sin_Personas"
            with open(ESTADO_FILE, "w", encoding="utf-8") as f:
                json.dump(estado, f, indent=2, ensure_ascii=False)
            
            guardar_resultado_en_resultados(cedula_a, cedula_p, placa, cedula_a, "Falló - Sin personas")

            # ⭐ NUEVO: Guardar en hoja "Sin asociados"
            fecha_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            guardar_sin_personas_en_sheets(fecha_actual, placa, cedula_a, cedula_p)
            
            resultados_finales["sin_personas_confirmado"].append({
                "placa": placa,
                "razon": "No hay personas en ambas cédulas"
            })
        
        else:
            logging.error(f"❌ {placa}: SIGUE SIENDO PENDIENTE tras {max_intentos} intentos")
            resultados_finales["fallidos_permanentes"].append({
                "placa": placa,
                "razon": f"Error técnico persistente tras {max_intentos} intentos"
            })
        
        # Preparar siguiente
        try:
            limpiar_todos_los_campos(driver)
            time.sleep(1)
            driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
            time.sleep(3)
        except:
            pass
    
    # ═══ REPORTE FINAL ═══
    generar_reporte_reintentos_finales(resultados_finales)
    return resultados_finales


# ═════════════════════════════════════════════════════════════
# ⭐ NUEVA FUNCIÓN: REPORTE DE REINTENTOS FINALES
# ═════════════════════════════════════════════════════════════

def generar_reporte_reintentos_finales(resultados):
    """Genera reporte de reintentos finales"""
    
    logging.info(f"\n{'='*70}")
    logging.info(f"📊 REPORTE FINAL DE REINTENTOS DE PENDIENTES")
    logging.info(f"{'='*70}")
    
    recuperados = len(resultados["recuperados"])
    sin_personas = len(resultados["sin_personas_confirmado"])
    fallidos = len(resultados["fallidos_permanentes"])
    total = recuperados + sin_personas + fallidos
    
    logging.info(f"\n✅ RECUPERADOS: {recuperados} registros")
    if resultados["recuperados"]:
        for item in resultados["recuperados"]:
            logging.info(f"   ✓ {item['placa']} (intento {item['intento']})")
    
    logging.info(f"\n❌ SIN PERSONAS (Confirmado): {sin_personas} registros")
    if resultados["sin_personas_confirmado"]:
        for item in resultados["sin_personas_confirmado"]:
            logging.info(f"   ✗ {item['placa']}")
    
    logging.info(f"\n⚠️  FALLIDOS PERMANENTES: {fallidos} registros")
    if resultados["fallidos_permanentes"]:
        for item in resultados["fallidos_permanentes"]:
            logging.info(f"   ⚠️  {item['placa']}")
    
    if total > 0:
        porcentaje = (recuperados / total * 100)
        logging.info(f"\n{'='*70}")
        logging.info(f"📈 RESUMEN: {recuperados}/{total} recuperados ({porcentaje:.1f}%)")
        logging.info(f"{'='*70}\n")
    else:
        logging.info(f"\n{'='*70}")
        logging.info(f"ℹ️  No había registros pendientes para reintentar")
        logging.info(f"{'='*70}\n")



# ═════════════════════════════════════════════════════════════
# 📊 TRACKING DE RESULTADOS PARA REINTENTOS
# ═════════════════════════════════════════════════════════════

def crear_estructura_resultados():
    """Crea la estructura para guardar resultados de cada intento"""
    return {
        "intento_1": {
            "exitosos": [],
            "fallos_tecnica": [],
            "fallos_datos": [],
            "sin_procesar": []
        },
        "intento_2": {
            "exitosos": [],
            "fallos_tecnica": [],
            "fallos_datos": []
        },
        "intento_3": {
            "exitosos": [],
            "fallos_tecnica": [],
            "fallos_datos": []
        }
    }

def agregar_resultado_tracking(tracking, intento, placa, cedula, estado, razon=""):
    """
    Agrega un resultado al tracking
    estado: "exitoso", "fallo_tecnica", "fallo_datos"
    """
    clave_intento = f"intento_{intento}"
    
    if estado == "exitoso":
        tracking[clave_intento]["exitosos"].append({
            "placa": placa,
            "cedula": cedula,
            "timestamp": datetime.now().isoformat()
        })
    elif estado == "fallo_tecnica":
        tracking[clave_intento]["fallos_tecnica"].append({
            "placa": placa,
            "cedula": cedula,
            "razon": razon,
            "timestamp": datetime.now().isoformat()
        })
    elif estado == "fallo_datos":
        tracking[clave_intento]["fallos_datos"].append({
            "placa": placa,
            "cedula": cedula,
            "razon": razon,
            "timestamp": datetime.now().isoformat()
        })

def generar_reporte_final(tracking):
    """Genera reporte ejecutivo final"""
    logging.info("\n" + "="*70)
    logging.info("📊 REPORTE FINAL EJECUTIVO")
    logging.info("="*70)
    
    total_exitosos = (
        len(tracking["intento_1"]["exitosos"]) +
        len(tracking["intento_2"]["exitosos"]) +
        len(tracking["intento_3"]["exitosos"])
    )
    
    total_fallos_tecnica = (
        len(tracking["intento_1"]["fallos_tecnica"]) +
        len(tracking["intento_2"]["fallos_tecnica"]) +
        len(tracking["intento_3"]["fallos_tecnica"])
    )
    
    total_fallos_datos = (
        len(tracking["intento_1"]["fallos_datos"]) +
        len(tracking["intento_2"]["fallos_datos"]) +
        len(tracking["intento_3"]["fallos_datos"])
    )
    
    total_procesados = total_exitosos + total_fallos_tecnica + total_fallos_datos
    
    logging.info(f"\n✅ EXITOSOS: {total_exitosos} registros")
    logging.info(f"⚠️  ERROR TÉCNICO (tras reintentos): {total_fallos_tecnica} registros")
    logging.info(f"❌ ERROR DE DATOS (sin asociados, placa inexistente): {total_fallos_datos} registros")
    logging.info(f"\n📈 TOTAL PROCESADO: {total_procesados} registros")
    logging.info("="*70 + "\n")
    
    # Detalles por intento
    logging.info("\n📋 DETALLES POR INTENTO:")
    logging.info(f"\n  🔄 INTENTO 1 (Inicial):")
    logging.info(f"     ✅ Exitosos: {len(tracking['intento_1']['exitosos'])}")
    logging.info(f"     ⚠️  Fallos técnica: {len(tracking['intento_1']['fallos_tecnica'])}")
    logging.info(f"     ❌ Fallos datos: {len(tracking['intento_1']['fallos_datos'])}")
    
    logging.info(f"\n  🔄 INTENTO 2 (Reintento técnica):")
    logging.info(f"     ✅ Exitosos: {len(tracking['intento_2']['exitosos'])}")
    logging.info(f"     ⚠️  Fallos técnica: {len(tracking['intento_2']['fallos_tecnica'])}")
    
    logging.info(f"\n  🔄 INTENTO 3 (Segundo reintento):")
    logging.info(f"     ✅ Exitosos: {len(tracking['intento_3']['exitosos'])}")
    logging.info(f"     ⚠️  Fallos técnica: {len(tracking['intento_3']['fallos_tecnica'])}")
    
    return tracking



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
    
    ⭐ BUSCA LA FILA EXISTENTE Y LA ACTUALIZA (sin duplicados)
    ⭐ SI NO EXISTE, CREA UNA NUEVA
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
        
        # ═══ OBTENER WORKSHEET ═══
        try:
            worksheet = sheet.worksheet("Resultados")
            logging.info("✅ Hoja 'Resultados' encontrada")
        except gspread.WorksheetNotFound:
            logging.error(f"❌ Hoja 'Resultados' NO encontrada. Debes crearla manualmente con los encabezados en la fila 1")
            return
        
        # ═══ OBTENER TODAS LAS FILAS ═══
        todas_filas = worksheet.get_all_values()
        logging.info(f"📊 Total de filas en Resultados: {len(todas_filas)}")
        
        # ═══ BUSCAR SI YA EXISTE ESTA PLACA + CÉDULA ASOCIADO ═══
        fila_existente = None
        numero_fila_existente = None
        
        for idx, fila in enumerate(todas_filas[1:], start=2):  # Comienza desde fila 2 (saltando encabezado)
            # Columna C (índice 2) es la placa
            # Columna A (índice 0) es la cédula asociado
            if len(fila) > 2:
                placa_en_fila = fila[2].strip() if len(fila) > 2 else ""
                cedula_en_fila = fila[0].strip() if len(fila) > 0 else ""
                
                if placa_en_fila == placa.strip() and cedula_en_fila == cedula_asociado.strip():
                    fila_existente = fila
                    numero_fila_existente = idx
                    logging.info(f"🔍 Fila existente encontrada en fila {numero_fila_existente}")
                    logging.info(f"   Contenido actual: {fila}")
                    break
        
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
        
        # ═══ ACTUALIZAR O CREAR NUEVA ═══
        if numero_fila_existente:
            # ⭐ ACTUALIZAR FILA EXISTENTE (SIN DUPLICAR)
            logging.info(f"\n🔄 ACTUALIZANDO fila {numero_fila_existente}")
            logging.info(f"   Anterior: {fila_existente}")
            logging.info(f"   Nuevo:    {nueva_fila}")
            
            rango = f"A{numero_fila_existente}:F{numero_fila_existente}"
            worksheet.update([nueva_fila], range_name=rango, value_input_option="RAW")
            
            logging.info(f"✅ Resultado ACTUALIZADO en 'Resultados' (fila {numero_fila_existente}):")
            logging.info(f"   Cédula Asociado: {cedula_asociado}")
            logging.info(f"   Cédula Propietario: {cedula_propietario}")
            logging.info(f"   Placa: {placa}")
            logging.info(f"   Estado: {estado_texto}")
            logging.info(f"   Cédula Usada: {cedula_usada}")
            
        else:
            # ⭐ CREAR NUEVA FILA (al final, sin duplicados)
            fila_nueva = len(todas_filas) + 1
            
            if fila_nueva < 2:
                fila_nueva = 2
            
            logging.info(f"\n📝 CREANDO nueva fila {fila_nueva}")
            logging.info(f"   (No existía previamente)")
            
            rango = f"A{fila_nueva}:F{fila_nueva}"
            worksheet.update([nueva_fila], range_name=rango, value_input_option="RAW")
            
            logging.info(f"✅ Resultado INSERTADO en 'Resultados' (fila {fila_nueva}):")
            logging.info(f"   Cédula Asociado: {cedula_asociado}")
            logging.info(f"   Cédula Propietario: {cedula_propietario}")
            logging.info(f"   Placa: {placa}")
            logging.info(f"   Estado: {estado_texto}")
            logging.info(f"   Cédula Usada: {cedula_usada}")
        
    except Exception as e:
        logging.error(f"❌ Error escribiendo en 'Resultados': {e}", exc_info=True)




def guardar_sin_personas_en_sheets(fecha, placa, cedula_asociado, cedula_propietario):
    """
    Guarda en la hoja 'Sin asociados' los registros sin personas
    Columnas:
    A: Fecha
    B: Placa
    C: Cédula Asociado
    D: Cédula Propietario
    E: Estado (Sin personas asociados)
    
    ⭐ BUSCA LA FILA EXISTENTE Y LA ACTUALIZA (sin duplicados)
    ⭐ SI NO EXISTE, CREA UNA NUEVA
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
        
        # ═══ OBTENER O CREAR WORKSHEET "Sin asociados" ═══
        try:
            worksheet = sheet.worksheet("Sin asociados")
            logging.info("✅ Hoja 'Sin asociados' encontrada")
        except gspread.WorksheetNotFound:
            worksheet = sheet.add_worksheet(title="Sin asociados", rows=1000, cols=10)
            logging.info("📝 Hoja 'Sin asociados' creada")
            
            # Crear encabezados
            encabezados = ["Fecha", "Placa", "Cédula Asociado", "Cédula Propietario", "Estado"]
            worksheet.update([encabezados], range_name="A1:E1", value_input_option="RAW")
            logging.info("✅ Encabezados creados en 'Sin asociados'")
        
        # ═══ OBTENER TODAS LAS FILAS ═══
        todas_filas = worksheet.get_all_values()
        logging.info(f"📊 Total de filas en Sin asociados: {len(todas_filas)}")
        
        # ═══ BUSCAR SI YA EXISTE ESTA PLACA + CÉDULA ASOCIADO ═══
        fila_existente = None
        numero_fila_existente = None
        
        for idx, fila in enumerate(todas_filas[1:], start=2):  # Comienza desde fila 2 (saltando encabezado)
            # Columna B (índice 1) es la placa
            # Columna C (índice 2) es la cédula asociado
            if len(fila) > 2:
                placa_en_fila = fila[1].strip() if len(fila) > 1 else ""
                cedula_en_fila = fila[2].strip() if len(fila) > 2 else ""
                
                if placa_en_fila == placa.strip() and cedula_en_fila == cedula_asociado.strip():
                    fila_existente = fila
                    numero_fila_existente = idx
                    logging.info(f"🔍 Fila existente encontrada en fila {numero_fila_existente}")
                    logging.info(f"   Contenido actual: {fila}")
                    break
        
        # ═══ PREPARAR DATOS ═══
        nueva_fila = [
            fecha,
            placa,
            cedula_asociado,
            cedula_propietario if cedula_propietario else "",  # Vacío si no hay
            "Sin personas asociados"
        ]
        
        # ═══ ACTUALIZAR O CREAR NUEVA ═══
        if numero_fila_existente:
            # ⭐ ACTUALIZAR FILA EXISTENTE (SIN DUPLICAR)
            logging.info(f"\n🔄 ACTUALIZANDO fila {numero_fila_existente}")
            logging.info(f"   Anterior: {fila_existente}")
            logging.info(f"   Nuevo:    {nueva_fila}")
            
            rango = f"A{numero_fila_existente}:E{numero_fila_existente}"
            worksheet.update([nueva_fila], range_name=rango, value_input_option="RAW")
            
            logging.info(f"✅ Registro ACTUALIZADO en 'Sin asociados' (fila {numero_fila_existente}):")
            logging.info(f"   Fecha: {fecha}")
            logging.info(f"   Placa: {placa}")
            logging.info(f"   Cédula Asociado: {cedula_asociado}")
            logging.info(f"   Cédula Propietario: {cedula_propietario}")
            logging.info(f"   Estado: Sin personas asociados")
            
        else:
            # ⭐ CREAR NUEVA FILA (al final, sin duplicados)
            fila_nueva = len(todas_filas) + 1
            
            if fila_nueva < 2:
                fila_nueva = 2
            
            logging.info(f"\n📝 CREANDO nueva fila {fila_nueva}")
            logging.info(f"   (No existía previamente)")
            
            rango = f"A{fila_nueva}:E{fila_nueva}"
            worksheet.update([nueva_fila], range_name=rango, value_input_option="RAW")
            
            logging.info(f"✅ Registro INSERTADO en 'Sin asociados' (fila {fila_nueva}):")
            logging.info(f"   Fecha: {fecha}")
            logging.info(f"   Placa: {placa}")
            logging.info(f"   Cédula Asociado: {cedula_asociado}")
            logging.info(f"   Cédula Propietario: {cedula_propietario}")
            logging.info(f"   Estado: Sin personas asociados")
        
    except Exception as e:
        logging.error(f"❌ Error escribiendo en 'Sin asociados': {e}", exc_info=True)





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

def guardar_en_sheets(resultados, actualizar_existente=False):
    """
    Guarda los resultados en Google Sheets - Cilindraje en columna D
    
    Si actualizar_existente=True: Busca la placa y actualiza la fila existente
    Si actualizar_existente=False: Agrega nueva fila al final
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
        worksheet = sheet.worksheet("Datos Runt")

        # ═══ PROCESAR CADA RESULTADO ═══
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
            
            fila_valores = [
                r["Tiempo ejecucion"],  # A
                r["cedula"],             # B
                r["placa"],              # C
                r.get("cilindraje", "No disponible"),  # D (NUEVO - independiente)
            ] + soat_data + rtm_data + [r["estado"]]   # E en adelante
            
            placa = r["placa"]
            
            # ═══ OBTENER TODAS LAS FILAS PARA BUSCAR PLACA ═══
            todas_filas = worksheet.get_all_values()
            logging.info(f"📊 Buscando placa {placa} en {len(todas_filas)} filas...")
            
            # ═══ BUSCAR LA FILA DONDE ESTÁ ESTA PLACA (columna C = índice 2) ═══
            numero_fila_existente = None
            
            for idx, fila in enumerate(todas_filas):
                if len(fila) > 2:  # Asegurar que tiene al menos 3 columnas
                    placa_en_fila = fila[2].strip() if fila[2] else ""
                    
                    if placa_en_fila == placa.strip():
                        numero_fila_existente = idx + 1  # +1 porque Google Sheets es 1-indexed
                        logging.info(f"🔍 ¡PLACA ENCONTRADA en fila {numero_fila_existente}!")
                        break
            
            # ═══ ACTUALIZAR O CREAR NUEVA ═══
            if numero_fila_existente and actualizar_existente:
                # ⭐ ACTUALIZAR FILA EXISTENTE (COMPLETA, SIN ROMPER ORDEN)
                logging.info(f"\n🔄 ACTUALIZANDO fila {numero_fila_existente}")
                logging.info(f"   Placa: {placa}")
                logging.info(f"   Tiempo: {r['Tiempo ejecucion']}")
                logging.info(f"   Estado: {r['estado']}")
                
                # Calcular el rango dinámicamente (A:S o más allá si hay más columnas)
                num_columnas = len(fila_valores)
                col_final = chr(64 + num_columnas)  # A=65, B=66, etc.
                
                rango = f"A{numero_fila_existente}:{col_final}{numero_fila_existente}"
                
                try:
                    worksheet.update([fila_valores], range_name=rango, value_input_option="RAW")
                    logging.info(f"✅ FILA ACTUALIZADA en {rango}")
                    logging.info(f"   ✓ Tiempo ejecucion: {r['Tiempo ejecucion']}")
                    logging.info(f"   ✓ Cedula: {r['cedula']}")
                    logging.info(f"   ✓ Placa: {placa}")
                    logging.info(f"   ✓ Cilindraje: {r.get('cilindraje', 'No disponible')}")
                    logging.info(f"   ✓ Estado: {r['estado']}")
                    
                except Exception as e:
                    logging.error(f"❌ Error en update: {e}")
                    logging.warning("🔄 Intentando método alternativo...")
                    
                    # Método alternativo: actualizar celda por celda
                    for col_num, valor in enumerate(fila_valores, start=1):
                        try:
                            worksheet.update_cell(numero_fila_existente, col_num, valor)
                            time.sleep(0.05)
                        except Exception as cell_error:
                            logging.warning(f"⚠️ Error actualizando celda ({numero_fila_existente},{col_num}): {cell_error}")
                    
                    logging.info(f"✅ Fila actualizada (método celda por celda)")
            
            elif numero_fila_existente and not actualizar_existente:
                # ⭐ SI LA PLACA EXISTE PERO NO QUEREMOS ACTUALIZAR
                logging.info(f"⏭️  Placa {placa} ya existe en fila {numero_fila_existente}")
                logging.info(f"   (actualizar_existente={actualizar_existente}, saltando...)")
            
            else:
                # ⭐ CREAR NUEVA FILA (NO EXISTE LA PLACA)
                fila_nueva = len(todas_filas) + 1
                
                if fila_nueva < 3:
                    fila_nueva = 3
                
                logging.info(f"\n📝 CREANDO NUEVA FILA {fila_nueva}")
                logging.info(f"   (Placa {placa} no existía previamente)")
                
                num_columnas = len(fila_valores)
                col_final = chr(64 + num_columnas)
                
                rango = f"A{fila_nueva}:{col_final}{fila_nueva}"
                
                try:
                    worksheet.update([fila_valores], range_name=rango, value_input_option="RAW")
                    logging.info(f"✅ NUEVA FILA INSERTADA en {rango}")
                    logging.info(f"   ✓ Placa: {placa}")
                    logging.info(f"   ✓ Estado: {r['estado']}")
                    
                except Exception as e:
                    logging.error(f"❌ Error en insert: {e}")
                    logging.warning("🔄 Intentando método alternativo...")
                    
                    for col_num, valor in enumerate(fila_valores, start=1):
                        try:
                            worksheet.update_cell(fila_nueva, col_num, valor)
                            time.sleep(0.05)
                        except Exception as cell_error:
                            logging.warning(f"⚠️ Error insertando en ({fila_nueva},{col_num}): {cell_error}")
                    
                    logging.info(f"✅ Nueva fila insertada (método celda por celda)")

        logging.info(f"\n✅ PROCESO DE GUARDADO COMPLETADO")

    except Exception as e:
        logging.error(f"❌ Error grave al guardar en Google Sheets: {e}", exc_info=True)

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


def procesar_consulta_interno(driver, cedula, placa, fila_numero, es_reintento=False, max_intentos_internos=2):
    """
    Lógica interna de procesamiento
    max_intentos_internos: Número de reintentos internos (default 2, reducir a 1 en reintentos)
    """
    max_reintentos = max_intentos_internos
    
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
                                
                                # ⭐ NUEVO: Guardar en hoja "Sin asociados"
                                # Buscar cedula_propietario correspondiente
                                cedula_prop_temp = ""
                                for ced_a, ced_p, plac, _, _ in datos_validos:
                                    if plac == placa:
                                        cedula_prop_temp = ced_p
                                        break
                                
                                fecha_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                guardar_sin_personas_en_sheets(fecha_actual, placa, cedula, cedula_prop_temp)
                                
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

def main():
    """Función principal CON VALIDACIÓN, REINTENTOS INTELIGENTES Y REANUDACIÓN"""
    driver = iniciar_driver()
    if not driver:
        logging.error("❌ No se pudo inicializar el driver")
        return

    driver.maximize_window()

    try:
        # ═══ CARGAR DATOS Y VALIDAR ═══
        datos_brutos = obtener_datos_unicos()
        datos_validos, datos_invalidos = filtrar_datos_validos(datos_brutos)
        
        if not datos_validos:
            logging.info(f"\n{'='*70}")
            logging.info(f"❌ NO HAY REGISTROS VÁLIDOS POR PROCESAR")
            logging.info(f"   Total cargado: {len(datos_brutos)}")
            logging.info(f"   Todos tienen datos incompletos")
            logging.info(f"{'='*70}\n")
            return
        
                # ═══ CARGAR ESTADO PREVIO Y DETERMINAR PUNTO DE INICIO ═══
        estado_actual = cargar_estado()
        placas_procesadas = estado_actual.get("resumen", {}).get("placas_procesadas", {})
        
        logging.info(f"\n{'='*70}")
        logging.info(f"📋 VERIFICANDO PUNTO DE REANUDACIÓN...")
        logging.info(f"{'='*70}")
        
        # 🎯 OBTENER ÍNDICE DESDE ÚLTIMO EXITOSO (NO DESDE ÚLTIMO PROCESADO)
        indice_reanudacion, ultima_exitosa, registros_exitosos = obtener_indice_reanudacion_desde_ultimo_exitoso(
            datos_validos,
            placas_procesadas
        )
        
        # ✂️ CORTAR DESDE EL PUNTO DE REANUDACIÓN
        datos_por_procesar = datos_validos[indice_reanudacion:]
        
        logging.info(f"\n{'='*70}")
        logging.info(f"📊 RESUMEN DE REANUDACIÓN")
        logging.info(f"{'='*70}")
        logging.info(f"Total en Google Sheets: {len(datos_validos)} registros")
        logging.info(f"Ya exitosos: {registros_exitosos} registros")
        logging.info(f"Pendientes por procesar: {len(datos_por_procesar)} registros")
        if ultima_exitosa:
            logging.info(f"Último exitoso: {ultima_exitosa}")
            if datos_por_procesar:
                logging.info(f"Próximo a procesar: {datos_por_procesar[0][2]}")
        logging.info(f"{'='*70}\n")
        
        if not datos_por_procesar:
            logging.info(f"\n{'='*70}")
            logging.info(f"✅ TODOS LOS REGISTROS DE SHEETS YA FUERON PROCESADOS")
            logging.info(f"   Total: {registros_exitosos} exitosos")
            logging.info(f"{'='*70}\n")
            return
        
        logging.info(f"\n{'='*70}")
        logging.info(f"🚀 INICIANDO PROCESAMIENTO - {len(datos_por_procesar)} REGISTROS PENDIENTES")
        logging.info(f"{'='*70}\n")
        
        # ═══ INICIALIZAR TRACKING ═══
        tracking_resultados = crear_estructura_resultados()
        
        # ═══ INTENTO 1: PROCESAMIENTO INICIAL ═══
        registros_procesados_contador = 0
        
        for i, (cedula_asociado, cedula_propietario, placa, fila_numero, sheet_origen) in enumerate(datos_por_procesar):
            
            # 🔄 REINICIO PERIÓDICO CADA 5 REGISTROS
            reiniciar_sesion_periodico(driver, registros_procesados_contador)
            registros_procesados_contador += 1
            
            logging.info(f"\n{'='*70}")
            logging.info(f"📊 Procesando [{i + 1}/{len(datos_por_procesar)}]: Placa {placa}")
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
                
                # Clasificar resultado
                if estado_resultado == "Exitoso":
                    agregar_resultado_tracking(tracking_resultados, 1, placa, cedula_asociado, "exitoso")
                elif estado_resultado == "Exitoso - Sin personas asociadas":
                    agregar_resultado_tracking(tracking_resultados, 1, placa, cedula_asociado, "fallo_datos", 
                                             "No hay personas asociadas al vehículo")
                else:
                    agregar_resultado_tracking(tracking_resultados, 1, placa, cedula_asociado, "fallo_tecnica",
                                             "Error en consulta RUNT")
                
                guardar_estado(cedula_asociado, placa, estado_resultado, i, len(datos_por_procesar))
                guardar_en_sheets([resultado])
                
                if "datos_vehiculo" in resultado and resultado["datos_vehiculo"]:
                    escribir_datos_vehiculo_sheets(resultado["datos_vehiculo"])
                    
                logging.info(f"✅ Datos de {placa} guardados")
            else:
                agregar_resultado_tracking(tracking_resultados, 1, placa, cedula_asociado, "fallo_tecnica",
                                         "No retornó resultado")
                guardar_estado(cedula_asociado, placa, "Pendiente", i, len(datos_por_procesar))

            # Preparar siguiente consulta
            if i < len(datos_por_procesar) - 1:
                try:
                    logging.info("🔄 Preparando siguiente consulta...")
                    limpiar_todos_los_campos(driver)
                    time.sleep(1)
                    
                    try:
                        otra_consulta_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Otra consulta')]"))
                        )
                    except:
                        logging.warning("⚠️  Recargando página...")
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
                    logging.error(f"⚠️  Error: {e}")
                    driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                    time.sleep(3)
                    limpiar_todos_los_campos(driver)
        
        # ═══ INTENTO 2: REINTENTAR SOLO FALLOS TÉCNICOS ═══
        fallos_tecnica_intento1 = tracking_resultados["intento_1"]["fallos_tecnica"]
        
        if fallos_tecnica_intento1:
            logging.info(f"\n{'='*70}")
            logging.info(f"🔄 INTENTO 2 - REINTENTANDO {len(fallos_tecnica_intento1)} FALLOS TÉCNICOS")
            logging.info(f"{'='*70}\n")
            
            contador_intento2 = 0
            for fallo in fallos_tecnica_intento1:
                
                # 🔄 REINICIO PERIÓDICO CADA 5 REGISTROS
                reiniciar_sesion_periodico(driver, contador_intento2)
                contador_intento2 += 1
                placa_reintento = fallo["placa"]
                cedula_reintento = fallo["cedula"]
                
                # Encontrar el cedula_propietario correspondiente
                cedula_prop_reintento = cedula_reintento
                for ced_a, ced_p, plac, _, _ in datos_validos:
                    if plac == placa_reintento:
                        cedula_prop_reintento = ced_p
                        break
                
                logging.info(f"\n🔄 INTENTO 2 - Reintentando: {placa_reintento}")
                
                # ⭐ NUEVO: Solo 1 intento interno en reintento
                resultado_reintento, _ = procesar_consulta_interno(
                    driver,
                    cedula_reintento,
                    placa_reintento,
                    0,
                    es_reintento=True,
                    max_intentos_internos=1
                )
                
                if resultado_reintento and resultado_reintento.get("estado") == "Exitoso":
                    agregar_resultado_tracking(tracking_resultados, 2, placa_reintento, cedula_reintento, "exitoso")
                    logging.info(f"✅ Recuperado en intento 2: {placa_reintento}")
                    guardar_en_sheets([resultado_reintento], actualizar_existente=True)
                    guardar_estado(cedula_reintento, placa_reintento, "Exitoso", 0, len(datos_por_procesar))
                
                elif resultado_reintento and resultado_reintento.get("estado") == "Exitoso - Sin personas asociadas":
                    # ⭐ NUEVO: Si sin personas con asociado, probar con propietario
                    logging.warning(f"⚠️  Sin personas con asociado, probando con propietario...")
                    
                    resultado_con_prop, _ = procesar_consulta_interno(
                        driver,
                        cedula_prop_reintento,
                        placa_reintento,
                        0,
                        es_reintento=True,
                        max_intentos_internos=1
                    )
                    
                    if resultado_con_prop and resultado_con_prop.get("estado") == "Exitoso":
                        agregar_resultado_tracking(tracking_resultados, 2, placa_reintento, cedula_reintento, "exitoso")
                        logging.info(f"✅ Recuperado en intento 2 (con propietario): {placa_reintento}")
                        guardar_en_sheets([resultado_con_prop], actualizar_existente=True)
                        guardar_estado(cedula_reintento, placa_reintento, "Exitoso", 0, len(datos_por_procesar))
                    else:
                        # ⭐ NUEVO: Confirmar sin personas y TERMINAR
                        agregar_resultado_tracking(tracking_resultados, 2, placa_reintento, cedula_reintento, 
                                                 "fallo_datos", "Confirmado: Sin personas asociadas en ambas cédulas")
                        logging.error(f"❌ Confirmado: Sin personas asociadas (asociado + propietario): {placa_reintento}")
                        
                        guardar_resultado_en_resultados(
                            cedula_reintento,
                            cedula_prop_reintento,
                            placa_reintento,
                            cedula_reintento,
                            "Falló - Sin personas asociadas"
                        )
                        guardar_estado(cedula_reintento, placa_reintento, "Pendiente", 0, len(datos_por_procesar))
                else:
                    agregar_resultado_tracking(tracking_resultados, 2, placa_reintento, cedula_reintento, 
                                             "fallo_tecnica", "Error técnico persistió")
                    logging.warning(f"⚠️  Error técnico persistió en intento 2: {placa_reintento}")
                    guardar_estado(cedula_reintento, placa_reintento, "Pendiente", 0, len(datos_por_procesar))
                
                time.sleep(2)
        
        # ═══ INTENTO 3: SEGUNDO REINTENTO (SOLO TÉCNICOS) ═══
        fallos_tecnica_intento2 = tracking_resultados["intento_2"]["fallos_tecnica"]
        
        if fallos_tecnica_intento2:
            logging.info(f"\n{'='*70}")
            logging.info(f"🔄 INTENTO 3 - SEGUNDO REINTENTO DE {len(fallos_tecnica_intento2)} REGISTROS")
            logging.info(f"{'='*70}\n")
            
            contador_intento3 = 0
            for fallo in fallos_tecnica_intento2:
                
                    # 🔄 REINICIO PERIÓDICO CADA 5 REGISTROS
                reiniciar_sesion_periodico(driver, contador_intento3)
                contador_intento3 += 1
                placa_reintento = fallo["placa"]
                cedula_reintento = fallo["cedula"]
                
                cedula_prop_reintento = cedula_reintento
                for ced_a, ced_p, plac, _, _ in datos_validos:
                    if plac == placa_reintento:
                        cedula_prop_reintento = ced_p
                        break
                
                logging.info(f"\n🔄 INTENTO 3 - Segundo reintento: {placa_reintento}")
                
                # ⭐ NUEVO: Solo 1 intento interno en tercer intento
                resultado_reintento, _ = procesar_consulta_interno(
                    driver,
                    cedula_reintento,
                    placa_reintento,
                    0,
                    es_reintento=True,
                    max_intentos_internos=1
                )
                
                if resultado_reintento and resultado_reintento.get("estado") == "Exitoso":
                    agregar_resultado_tracking(tracking_resultados, 3, placa_reintento, cedula_reintento, "exitoso")
                    logging.info(f"✅ Recuperado en intento 3: {placa_reintento}")
                    guardar_en_sheets([resultado_reintento], actualizar_existente=True)
                    guardar_estado(cedula_reintento, placa_reintento, "Exitoso", 0, len(datos_por_procesar))
                else:
                    agregar_resultado_tracking(tracking_resultados, 3, placa_reintento, cedula_reintento, 
                                             "fallo_tecnica", "Falló en todos los intentos")
                    logging.error(f"❌ Falló permanentemente tras 3 intentos: {placa_reintento}")
                    guardar_estado(cedula_reintento, placa_reintento, "Pendiente", 0, len(datos_por_procesar))
                
                time.sleep(2)
        
                # ═══ GENERAR REPORTE FINAL DE INTENTOS 1-3 ═══
        generar_reporte_final(tracking_resultados)
        
        # 🔴 FASE FINAL: REINTENTAR PENDIENTES PARA CONFIRMAR ESTADO
        logging.info(f"\n\n{'='*70}")
        logging.info(f"⭐ INICIANDO FASE FINAL: VALIDAR PENDIENTES")
        logging.info(f"{'='*70}\n")
        
        resultados_finales = procesar_reintentos_finales_pendientes(
            driver,
            datos_validos,
            max_intentos=5  # ← Máximo 5 intentos por pendiente
        )
        
        # ═════════════════════════════════════════════════════════════
        
        logging.info(f"\n{'='*70}")
        logging.info(f"✅ PROCESO COMPLETADO EXITOSAMENTE")
        logging.info(f"{'='*70}\n")

    except Exception as e:
        logging.error(f"❌ Error principal: {e}", exc_info=True)
    finally:
        cerrar_driver(driver)

if __name__ == "__main__":
    logging.info("🚀 Iniciando Bot RUNT FINAL...")
    main()