"""
Script Verificador de Fallos para el Bot RUNT
==========================================
Este script funciona de forma independiente pero reutiliza la lógica core de Runt.py.
Su objetivo es procesar EXCLUSIVAMENTE los registros que quedaron marcados como 'Falló'
en la hoja 'Resultados' de Google Sheets.

Opera en su propia carpeta 'Verificacion' con sus propios logs y estado JSON.
Implementa una lógica agresiva de reintentos para distinguir entre un error técnico/captcha
y un mensaje real de "Sin personas asociadas".

Hojas de Sheets que afecta:
- Resultados: Actualiza el estado (Funcionó / Sin personas asociadas / Falló (persistente)).
- Sin Asociados: Registra de forma única las placas que no tienen dueños activos.
- Datos Runt / Datos Vehiculo: Registra datos si logra recuperar un vehículo.
"""

import logging
import time
import json
import os
from datetime import datetime
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials

# ============================================================
# 1. CONFIGURACIÓN Y ESTRUCTURA DE CARPETAS (EXCLUSIVA)
# ============================================================

# Crear carpeta Verificacion si no existe
BASE_PATH = Path(__file__).parent / "Verificacion"
BASE_PATH.mkdir(parents=True, exist_ok=True)

# Rutas de archivos propias
ESTADO_FILE = BASE_PATH / "estado_verificacion.json"
LOG_FILE = BASE_PATH / "verificacion.log"

# IDs y Nombres de Hojas de Sheets
SHEET_ID_BASE = "1hYwf3AMdUWY6Lk6VjeGKPiDnR003ftkybv79b_pUN64" # Tu ID de Sheet Base
NAME_RESULTADOS = "Resultados"
NAME_SIN_ASOCIADOS = "Sin Asociados"
NAME_DATOS_RUNT = "Datos Runt"
NAME_DATOS_VEHICULO = "Datos Vehiculo"

# Limpiar handlers de logging previos (para que no use el de Runt.py)
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

# Configurar logging EXCLUSIVO para la verificación
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding='utf-8')
    ],
    force=True # Forzar esta configuración
)

# Importar funciones CORE de Runt.py (Asumiendo que Runt.py está en la misma carpeta)
try:
    from Runt import (
        iniciar_driver, cerrar_driver, GOOGLE_CREDS,
        limpiar_todos_los_campos, procesar_consulta_interno,
        escribir_datos_vehiculo_en_sheets, guardar_en_sheets
    )
    logging.info("✅ Funciones core de Runt.py importadas correctamente")
except ImportError as e:
    logging.error(f"❌ Error crítico: No se pudo importar Runt.py. Asegúrate de que esté en la misma carpeta. Error: {e}")
    exit()

# ============================================================
# 2. FUNCIONES DE GOOGLE SHEETS (ESPECÍFICAS)
# ============================================================

def conectar_google_sheets():
    """Establece conexión con Google Sheets API"""
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
        logging.info("✅ Conexión a Google Sheets establecida para Verificación")
        return client
    except Exception as e:
        logging.error(f"❌ Error conectando a Google Sheets: {e}")
        return None

def leer_registros_fallidos(client):
    """Lee registros de 'Resultados' que tengan estado 'Falló'"""
    try:
        sheet = client.open_by_key(SHEET_ID_BASE)
        worksheet = sheet.worksheet(NAME_RESULTADOS)
        
        todas_filas = worksheet.get_all_values()
        if len(todas_filas) < 2:
            logging.info("ℹ️ La hoja Resultados está vacía o solo tiene encabezados")
            return []
        
        registros_fallidos = []
        
        # Columnas en Resultados: A:CedulaAsoc, B:CedulaProp, C:Placa, D:(vacío), E:Estado, F:CedulaUsada
        for idx, fila in enumerate(todas_filas[1:], start=2): # Start 2 para obtener el número de fila real
            if len(fila) < 5: continue # Fila incompleta
                
            placa = fila[2].strip()
            estado = fila[4].strip()
            
            # Solo nos interesan los que dicen 'Falló'
            if estado == "Falló" and placa:
                cedula_asoc = fila[0].strip()
                cedula_prop = fila[1].strip()
                registros_fallidos.append({
                    'cedula_asociado': cedula_asoc,
                    'cedula_propietario': cedula_prop,
                    'placa': placa,
                    'numero_fila_resultados': idx # Guardamos la fila para reescribir después
                })
        
        logging.info(f"📊 Se encontraron {len(registros_fallidos)} registros marcados como 'Falló' para verificar.")
        return registros_fallidos
        
    except Exception as e:
        logging.error(f"❌ Error leyendo hoja Resultados: {e}")
        return []

def garantizar_hoja_sin_asociados(client):
    """Verifica si existe la hoja 'Sin Asociados', si no, la crea con encabezados"""
    try:
        sheet = client.open_by_key(SHEET_ID_BASE)
        
        try:
            worksheet = sheet.worksheet(NAME_SIN_ASOCIADOS)
            logging.info(f"✅ Hoja '{NAME_SIN_ASOCIADOS}' detectada.")
            return worksheet
        except gspread.WorksheetNotFound:
            logging.info(f"📝 Creando hoja '{NAME_SIN_ASOCIADOS}' con encabezados...")
            # Crear hoja con 5 columnas
            worksheet = sheet.add_worksheet(title=NAME_SIN_ASOCIADOS, rows=1000, cols=5)
            
            # Definir encabezados A:E
            encabezados = ["Fecha", "Placa", "Cédula Asociado", "Cédula Propietario", "Estado"]
            # Rango A1:E1
            worksheet.update('A1:E1', [encabezados], value_input_option="RAW")
            logging.info(f"✅ Hoja '{NAME_SIN_ASOCIADOS}' creada correctamente.")
            return worksheet
            
    except Exception as e:
        logging.error(f"❌ Error garantizando hoja 'Sin Asociados': {e}")
        return None

def registrar_en_sin_asociados(worksheet, datos_registro):
    """Escribe un registro en la hoja 'Sin Asociados' evitando duplicados de placa"""
    try:
        if not worksheet: return False
        
        placa_consultada = datos_registro['placa'].strip().upper()
        
        # 1. Verificar duplicados (leer columna Placa = B)
        # Esto puede ser lento si hay miles de datos, pero es lo más seguro sin un DB
        logging.info(f"🔍 Verificando si placa {placa_consultada} ya está en 'Sin Asociados'...")
        placas_existentes = worksheet.col_values(2) # Columna B
        
        # Normalizar placas existentes para comparar
        placas_existentes_norm = [p.strip().upper() for p in placas_existentes]
        
        if placa_consultada in placas_existentes_norm:
            logging.info(f"⏭️ La placa {placa_consultada} ya está registrada en 'Sin Asociados'. Saltando escritura.")
            return True # Consideramos éxito porque ya está ahí
        
        # 2. Si no es duplicado, escribir datos
        fecha_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Fecha, Placa, Cédula Asociado, Cédula Propietario, Estado
        fila_datos = [
            fecha_actual,
            placa_consultada,
            datos_registro['cedula_asociado'],
            datos_registro['cedula_propietario'],
            "Sin personas asociadas" # Estado fijo para esta hoja
        ]
        
        # Append row es eficiente aquí
        worksheet.append_row(fila_datos, value_input_option="RAW")
        logging.info(f"✅ Placa {placa_consultada} registrada exitosamente en 'Sin Asociados'.")
        return True
        
    except Exception as e:
        logging.error(f"❌ Error registrando en 'Sin Asociados': {e}")
        return False

def actualizar_hoja_resultados(client, numero_fila, nuevo_estado, cedula_usada=""):
    """Actualiza la fila correspondiente en la hoja Resultados"""
    try:
        sheet = client.open_by_key(SHEET_ID_BASE)
        worksheet = sheet.worksheet(NAME_RESULTADOS)
        
        # Columna E es Estado (índice 5), Columna F es CedulaUsada (índice 6)
        
        # Actualizar Estado
        worksheet.update_cell(numero_fila, 5, nuevo_estado)
        # Actualizar Cédula Usada
        worksheet.update_cell(numero_fila, 6, cedula_usada)
        
        logging.info(f"✅ Hoja Resultados actualizada (Fila {numero_fila}): Estado={nuevo_estado}, Cédula={cedula_usada}")
        return True
    except Exception as e:
        logging.error(f"❌ Error actualizando hoja Resultados en fila {numero_fila}: {e}")
        return False

# ============================================================
# 3. LÓGICA DE VERIFICACIÓN AGRESIVA (CORE)
# ============================================================

def determinar_tipo_error(driver):
    """
    Analiza la página para distinguir entre:
    - 'sin_personas': Mensaje modal de datos no corresponden / sin propietarios
    - 'captcha_error': El captcha fue incorrecto o expiró
    - 'exito': Se cargaron los datos del vehículo
    - 'error_tecnico': Otro tipo de error o carga infinita
    
    Usa los XPaths definidos en Runt.py de forma implícita a través de la lógica de procesar_consulta_interno
    """
    # Esta función en realidad no se usa directamente, ya que procesar_consulta_interno
    # ya hace esta distinción y devuelve el 'estado' en el dict.
    # Se mantiene aquí solo por claridad conceptual de lo que el script requiere.
    pass

def verificar_registro_completo(driver, client, worksheet_sin_asoc, registro):
    """
    Ejecuta la estrategia de reintentos agresiva para UN registro.
    
    Estrategia:
    1. Intentar con Cédula Asociado.
    2. Si da 'captcha_incorrecto', reintentar hasta 3 veces (interno en Runt.py).
    3. Si el mensaje es 'Exitoso - Sin personas asociadas':
       a. Intentar con Cédula Propietario (si existe).
       b. Si Propietario también da 'Sin personas asociadas', REGISTRAR COMO SIN ASOCIADOS.
    4. Si es 'Exitoso' (con datos) en cualquier intento: REGISTRAR EN HOJAS PRINCIPALES.
    """
    placa = registro['placa']
    ced_asoc = registro['cedula_asociado']
    ced_prop = registro['cedula_propietario']
    fila_res = registro['numero_fila_resultados']
    
    logging.info(f"\n{'='*60}")
    logging.info(f"🔍 VERIFICANDO PLACA: {placa}")
    logging.info(f"   CedAsoc: {ced_asoc}, CedProp: {ced_prop}")
    logging.info(f"{'='*60}")
    
    # --- FASE 1: INTENTAR CON CÉDULA ASOCIADO ---
    logging.info(f"⏱️ Fase 1: Intentando con Cédula ASOCIADO ({ced_asoc})...")
    
    # Reutilizamos procesar_consulta_interno de Runt.py.
    # Esta función ya maneja internamente reintentos de captcha si detecta el modal correspondiente.
    # Le pasamos max_intentos_internos=3 para ser agresivos con el captcha.
    resultado_asoc, _ = procesar_consulta_interno(
        driver, 
        ced_asoc, 
        placa, 
        0, # fila_numero no se usa internamente para lógica de Sheets
        es_reintento=False, 
        max_intentos_internos=3 
    )
    
    # Analizar resultado de Fase 1
    if resultado_asoc:
        estado_asoc = resultado_asoc.get('estado')
        
        if estado_asoc == "Exitoso":
            # 🎉 Éxito con el Asociado!
            logging.info(f"✅ ¡ÉXITO recuperando datos con Cédula ASOCIADO para {placa}!")
            procesar_recuperacion_exitosa(client, resultado_asoc, ced_asoc, fila_res)
            return "recuperado_asoc"
            
        elif estado_asoc == "Exitoso - Sin personas asociadas":
            # Mensaje real de sin personas. Toca verificar con Propietario.
            logging.info(f"⚠️ Cédula ASOCIADO confirmó: Sin personas asociadas para {placa}.")
            
            # --- FASE 2: INTENTAR CON CÉDULA PROPIETARIO (Si aplica) ---
            if ced_prop and ced_prop.strip() != "" and ced_prop.lower() != "nan":
                logging.info(f"⏱️ Fase 2: Intentando con Cédula PROPIETARIO ({ced_prop})...")
                
                # Necesitamos reiniciar la página para nueva consulta
                try:
                    otra_consulta_btn = driver.find_element_by_xpath("//button[contains(., 'Otra consulta')]")
                    otra_consulta_btn.click()
                    time.sleep(2)
                except:
                    driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                    time.sleep(3)
                
                limpiar_todos_los_campos(driver)
                
                resultado_prop, _ = procesar_consulta_interno(
                    driver, 
                    ced_prop, 
                    placa, 
                    0, 
                    es_reintento=True, # Importante marcar como reintento
                    max_intentos_internos=3
                )
                
                if resultado_prop:
                    estado_prop = resultado_prop.get('estado')
                    
                    if estado_prop == "Exitoso":
                        # 🎉 Éxito con el Propietario! (El asociado estaba mal)
                        logging.info(f"✅ ¡ÉXITO recuperando datos con Cédula PROPIETARIO para {placa}!")
                        procesar_recuperacion_exitosa(client, resultado_prop, ced_prop, fila_res)
                        return "recuperado_prop"
                        
                    elif estado_prop == "Exitoso - Sin personas asociadas":
                        # 💥 AMBAS CÉDULAS CONFIRMAN SIN PERSONAS
                        logging.error(f"❌ CONFIRMADO: Ni Asociado ni Propietario tienen dueños activos para {placa}.")
                        
                        # Registrar en 'Sin Asociados'
                        registrar_en_sin_asociados(worksheet_sin_asoc, registro)
                        
                        # Actualizar 'Resultados' con el estado final correcto
                        actualizar_hoja_resultados(client, fila_res, "Sin personas asociadas", f"Asoc:{ced_asoc} / Prop:{ced_prop}")
                        return "sin_asociados_confirmado"
                    else:
                        logging.warning(f"⚠️ Cédula PROPIETARIO falló por error técnico ({estado_prop}) para {placa}.")
                        # No actualizamos Resultados aquí, dejamos el 'Falló' original para que se intente en otro ciclo.
                        return "fallo_tecnico_prop"
                else:
                    logging.warning(f"⚠️ No se obtuvo respuesta de Cédula PROPIETARIO para {placa}.")
                    return "error_consulta_prop"
                    
            else:
                # No hay cédula propietario registrada. El fallo con Asociado es definitivo por ahora.
                logging.error(f"❌ Cédula ASOCIADO confirmó sin personas y NO hay Cédula PROPIETARIO registrada para {placa}.")
                
                # Registrar en 'Sin Asociados' (asumimos que el único dato es el definitivo)
                registrar_en_sin_asociados(worksheet_sin_asoc, registro)
                
                # Actualizar 'Resultados'
                actualizar_hoja_resultados(client, fila_res, "Sin personas asociadas", f"Asoc:{ced_asoc}")
                return "sin_asociados_definitivo_asoc"
        
        else:
            # Estado es 'Falló - Error técnico' o similar (captcha falló tras reintentos, etc.)
            logging.warning(f"⚠️ Cédula ASOCIADO falló por error técnico ({estado_asoc}) para {placa}. No se pudo confirmar 'Sin personas'.")
            # Dejamos el 'Falló' original en Resultados.
            return "fallo_tecnico_asoc"
            
    else:
        # resultado_asoc es None (error crítico en la función)
        logging.warning(f"⚠️ Error crítico consultando Cédula ASOCIADO para {placa}.")
        return "error_critico_asoc"

def procesar_recuperacion_exitosa(client, resultado, cedula_exitosa, fila_resultados):
    """
    Lógica para guardar datos en hojas principales cuando la verificación tiene éxito.
    """
    placa = resultado['placa']
    try:
        # 1. Guardar en Datos Runt (reutilizando función de Runt.py)
        # Necesitamos pasar el cliente a la función modificado ligeramente si Runt.py usa creds directas.
        # Asumiremos que Runt.py funciona autónomamente.
        guardar_en_sheets([resultado])
        logging.info(f"   ✓ Datos Runt actualizados para {placa}")
        
        # 2. Guardar en Datos Vehiculo (reutilizando función de Runt.py)
        if 'datos_vehiculo' in resultado and resultado['datos_vehiculo']:
            escribir_datos_vehiculo_en_sheets(resultado['datos_vehiculo'], cedula_exitosa, placa)
            logging.info(f"   ✓ Datos Vehiculo actualizados para {placa}")
        
        # 3. Actualizar hoja Resultados (reescribir fila)
        # Cambiar 'Falló' por 'Funcionó' y poner la cédula que sirvió
        actualizar_hoja_resultados(client, fila_resultados, "Funcionó", cedula_exitosa)
        logging.info(f"   ✓ Hoja Resultados actualizada (Fila {fila_resultados}) para {placa}")
        
    except Exception as e:
        logging.error(f"❌ Error procesando éxito de {placa}: {e}")

# ============================================================
# 4. MANEJO DE ESTADO JSON (PROPIO)
# ============================================================

def cargar_estado_verificacion():
    """Carga el estado JSON de la carpeta Verificacion"""
    if os.path.exists(ESTADO_FILE):
        try:
            with open(ESTADO_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            logging.warning("⚠️ Archivo de estado JSON corrupto, creando nuevo...")
            return {}
    return {}

def guardar_estado_verificacion(estado):
    """Guarda el estado JSON en la carpeta Verificacion"""
    try:
        with open(ESTADO_FILE, "w", encoding="utf-8") as f:
            json.dump(estado, f, indent=2, ensure_ascii=False)
        logging.debug("💾 Estado de verificación guardado en JSON.")
    except Exception as e:
        logging.error(f"❌ Error guardando estado JSON: {e}")

# ============================================================
# 5. FUNCIÓN PRINCIPAL (MAIN)
# ============================================================

def main():
    """Flujo principal del verificador con reingreso visual"""
    logging.info("\n" + "="*80)
    logging.info("🚀 INICIANDO BOT VERIFICADOR DE FALLOS RUNT")
    logging.info("="*80)
    
    client = conectar_google_sheets()
    if not client: return
    
    registros_verificar = leer_registros_fallidos(client)
    if not registros_verificar:
        logging.info("✅ No hay registros pendientes.")
        return
    
    worksheet_sin_asoc = garantizar_hoja_sin_asociados(client)
    estado_json = cargar_estado_verificacion()
    
    driver = iniciar_driver()
    if not driver: return
    
    driver.maximize_window()
    
    # --- CAMBIO AQUÍ: Añadimos un contador ---
    contador_procesados = 0 

    try:
        for i, registro in enumerate(registros_verificar):
            # ═══ LÓGICA DE REINGRESO VISUAL CADA 5 PLACAS ═══
            if contador_procesados > 0 and contador_procesados % 5 == 0:
                logging.info("🔄 REINGRESO VISUAL: Refrescando página principal del Runt...")
                driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                time.sleep(5) # Espera para que veas la carga
            # ═══════════════════════════════════════════════

            resultado_verificacion = verificar_registro_completo(driver, client, worksheet_sin_asoc, registro)
            
            # Incrementamos el contador después de cada placa
            contador_procesados += 1
            
            # (El resto del código de guardado de estado sigue igual...)
            datos_json_actualizados = {
                'ultimo_intento': datetime.now().isoformat(),
                'resultado_verificacion': resultado_verificacion
            }
            estado_json[registro['placa']] = datos_json_actualizados
            guardar_estado_verificacion(estado_json)
            
            logging.info(f"✅ Avance: {i+1}/{len(registros_verificar)} | Sesión: {contador_procesados}/5")
            time.sleep(2)
            
    finally:
        cerrar_driver(driver)
        logging.info("🔚 Bot Verificador de Fallos finalizado.")

if __name__ == "__main__":
    main()