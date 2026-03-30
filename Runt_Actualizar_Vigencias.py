"""
Script para procesar vigencias de SOAT y Tecnomecanica EN BUCLE INFINITO
SOLO LEE desde las hojas 'Vigencias Soat' y 'Vigencias Tecnomecanica'
Columnas para lectura:
- C: Cédula Asociado
- D: Placa
- E: Cédula Propietario
- F: Estado Vigencia (filtra: No vigente, SE VENCE PRONTO, SE VENCE HOY)

ESCRIBE los resultados actualizados en la hoja 'Datos Runt'
USA SU PROPIO ARCHIVO DE ESTADO (estado_vigencias.json) - NO INTERFIERE CON Runt.py
"""

import logging
import time
import json
import os
import signal
import sys
from datetime import datetime
from typing import Optional, Dict, List, Tuple
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials


# Configuración logging - USAR ARCHIVOS PROPIOS
LOG_DIR = Path("logs_vigencias")
LOG_DIR.mkdir(exist_ok=True)

# Logger principal
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_DIR / "vigencias_automatizacion.log", encoding='utf-8')
    ]
)


# Importar funciones del Runt.py
from Runt import (
    iniciar_driver, cerrar_driver, limpiar_todos_los_campos,
    procesar_consulta_interno, GOOGLE_CREDS
)



# Logger separado para resultados exitosos
exitosos_logger = logging.getLogger('exitosos')
exitosos_logger.setLevel(logging.INFO)
exitosos_logger.propagate = False
exitosos_handler = logging.FileHandler(LOG_DIR / "exitosos.log", encoding='utf-8')
exitosos_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
exitosos_logger.addHandler(exitosos_handler)

# Logger separado para errores
errores_logger = logging.getLogger('errores')
errores_logger.setLevel(logging.INFO)
errores_logger.propagate = False
errores_handler = logging.FileHandler(LOG_DIR / "errores.log", encoding='utf-8')
errores_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
errores_logger.addHandler(errores_handler)

# Logger separado para resumen de ciclos
ciclos_logger = logging.getLogger('ciclos')
ciclos_logger.setLevel(logging.INFO)
ciclos_logger.propagate = False
ciclos_handler = logging.FileHandler(LOG_DIR / "ciclos.log", encoding='utf-8')
ciclos_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
ciclos_logger.addHandler(ciclos_handler)

# ID de la hoja de cálculo
SHEET_ID = "1vs414iH3QVeLoTcY2CExg4kD9eCkXZRRfax_WTlUXPk"

# Nombres de las hojas
SHEET_SOAT = "Vigencias Soat"
SHEET_TECNOMECANICA = "Vigencias Tecnomecanica"
SHEET_DATOS_RUNT = "Datos Runt"

# Estados a buscar en columna F (SOLO LECTURA)
ESTADOS_BUSCAR = ["No vigente", "SE VENCE PRONTO", "SE VENCE HOY"]

# Columnas para lectura (1-indexado)
COL_CEDULA_ASOCIADO = 3   # Columna C
COL_PLACA = 4              # Columna D
COL_CEDULA_PROPIETARIO = 5 # Columna E
COL_ESTADO_ACTUAL = 6      # Columna F

# Configuración del bucle
PAUSA_ENTRE_CICLOS = 3600  # 1 hora en segundos (ajustable)
PAUSA_CORTA = 5            # Segundos entre consultas
REINICIO_DRIVER_CADA_CICLOS = 10  # Reiniciar driver cada X ciclos completos

# ARCHIVO DE ESTADO PROPIO (NO INTERFIERE CON Runt.py)
ESTADO_VIGENCIAS_FILE = "estado_vigencias.json"


def estructura_estado_inicial():
    """Crea la estructura inicial del archivo de estado propio"""
    return {
        "version": "1.0",
        "tipo": "vigencias",
        "soat": {
            "procesadas": {},  # {"placa": {"estado": "exitoso", "fecha": "iso", "vigencia": "..."}}
            "ultima_ejecucion": None,
            "total_exitosas": 0,
            "total_fallidas": 0,
            "ultimo_ciclo": 0
        },
        "tecnomecanica": {
            "procesadas": {},
            "ultima_ejecucion": None,
            "total_exitosas": 0,
            "total_fallidas": 0,
            "ultimo_ciclo": 0
        },
        "estadisticas": {
            "total_ciclos": 0,
            "fecha_inicio": None,
            "fecha_ultimo_ciclo": None,
            "total_consultas": 0
        }
    }


def cargar_estado_vigencias():
    """Carga el estado de vigencias desde archivo propio"""
    if os.path.exists(ESTADO_VIGENCIAS_FILE):
        try:
            with open(ESTADO_VIGENCIAS_FILE, "r", encoding="utf-8") as f:
                estado = json.load(f)
                logging.debug(f"✅ Estado cargado desde {ESTADO_VIGENCIAS_FILE}")
                return estado
        except Exception as e:
            logging.warning(f"⚠️ Error cargando estado: {e}, creando nuevo")
            return estructura_estado_inicial()
    else:
        logging.info(f"📁 Creando nuevo archivo de estado: {ESTADO_VIGENCIAS_FILE}")
        return estructura_estado_inicial()


def guardar_estado_vigencias(estado):
    """Guarda el estado de vigencias en archivo propio"""
    try:
        with open(ESTADO_VIGENCIAS_FILE, "w", encoding="utf-8") as f:
            json.dump(estado, f, indent=2, ensure_ascii=False)
        logging.debug(f"💾 Estado guardado en {ESTADO_VIGENCIAS_FILE}")
    except Exception as e:
        logging.error(f"❌ Error guardando estado: {e}")


def reiniciar_estado_vigencias():
    """Reinicia el estado (útil para pruebas)"""
    estado = estructura_estado_inicial()
    guardar_estado_vigencias(estado)
    logging.info("🔄 Estado de vigencias REINICIADO")
    return estado


class VigenciaProcessor:
    """Clase para procesar vigencias de SOAT y Tecnomecanica en bucle infinito"""
    
    def __init__(self):
        """Inicializa el procesador"""
        self.client = None
        self.sheet = None
        self.driver = None
        self.contador_consultas_ciclo = 0
        self.contador_ciclos = 0
        self.ejecutando = True
        self.estado = cargar_estado_vigencias()
        
        # Configurar señal para cerrar gracefulmente
        signal.signal(signal.SIGINT, self.signal_handler)
        signal.signal(signal.SIGTERM, self.signal_handler)
        
    def signal_handler(self, signum, frame):
        """Maneja la señal de cierre para detener el bucle gracefulmente"""
        logging.info(f"\n⚠️ Señal de cierre recibida. Guardando estado final...")
        self.estado["estadisticas"]["fecha_ultimo_ciclo"] = datetime.now().isoformat()
        guardar_estado_vigencias(self.estado)
        self.ejecutando = False
        
    def conectar_google_sheets(self):
        """Conecta a Google Sheets"""
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
            
            self.client = gspread.authorize(creds)
            self.sheet = self.client.open_by_key(SHEET_ID)
            logging.info("✅ Conexión a Google Sheets establecida")
            return True
            
        except Exception as e:
            logging.error(f"❌ Error conectando a Google Sheets: {e}")
            return False
    
    def reiniciar_driver(self):
        """Reinicia el driver de Chrome"""
        try:
            if self.driver:
                cerrar_driver(self.driver)
                time.sleep(2)
            
            self.driver = iniciar_driver()
            if self.driver:
                self.driver.maximize_window()
                logging.info("✅ Driver reiniciado correctamente")
                return True
            else:
                logging.error("❌ No se pudo reiniciar el driver")
                return False
        except Exception as e:
            logging.error(f"❌ Error reiniciando driver: {e}")
            return False
    
    def leer_registros_desde_vigencias(self, sheet_name: str, tipo: str) -> List[Tuple[str, str, str, int]]:
        """
        SOLO LEE los registros NO procesados exitosamente desde la hoja de vigencias
        
        Args:
            sheet_name: Nombre de la hoja
            tipo: "soat" o "tecnomecanica"
            
        Returns:
            Lista de tuplas (cedula_asociado, cedula_propietario, placa, numero_fila)
        """
        try:
            worksheet = self.sheet.worksheet(sheet_name)
            todas_filas = worksheet.get_all_values()
            
            if not todas_filas or len(todas_filas) < 2:
                logging.warning(f"⚠️ La hoja {sheet_name} está vacía o no tiene datos")
                return []
            
            # Obtener placas ya procesadas exitosamente
            procesadas = self.estado[tipo]["procesadas"]
            
            registros_a_procesar = []
            saltadas = 0
            
            for idx, fila in enumerate(todas_filas[1:], start=2):
                # Leer valores
                cedula_asociado = fila[COL_CEDULA_ASOCIADO - 1].strip() if len(fila) > COL_CEDULA_ASOCIADO - 1 else ""
                placa = fila[COL_PLACA - 1].strip().upper() if len(fila) > COL_PLACA - 1 else ""
                cedula_propietario = fila[COL_CEDULA_PROPIETARIO - 1].strip() if len(fila) > COL_CEDULA_PROPIETARIO - 1 else ""
                estado_actual = fila[COL_ESTADO_ACTUAL - 1].strip() if len(fila) > COL_ESTADO_ACTUAL - 1 else ""
                
                if not placa or not cedula_asociado:
                    continue
                
                # Saltar si ya fue procesada exitosamente
                if placa in procesadas and procesadas[placa].get("estado") == "exitoso":
                    saltadas += 1
                    continue
                
                # Filtrar por estado
                if estado_actual in ESTADOS_BUSCAR:
                    registros_a_procesar.append((cedula_asociado, cedula_propietario, placa, idx))
            
            if saltadas > 0:
                logging.info(f"   ⏭️ {saltadas} placas ya procesadas exitosamente (saltadas)")
            
            return registros_a_procesar
            
        except gspread.WorksheetNotFound:
            logging.error(f"❌ Hoja {sheet_name} no encontrada")
            return []
        except Exception as e:
            logging.error(f"❌ Error leyendo datos de {sheet_name}: {e}")
            return []
    
    def obtener_vigencia_actual(self, resultado: Dict, tipo: str) -> str:
        """Extrae el estado de vigencia del resultado de la consulta"""
        try:
            if tipo == "soat":
                soat_datos = resultado.get("datos_soat", [])
                if soat_datos and len(soat_datos) > 4:
                    return soat_datos[4] if soat_datos[4] else "No disponible"
            else:
                rtm_datos = resultado.get("datos_técnicos", [])
                if rtm_datos and len(rtm_datos) > 4:
                    return rtm_datos[4] if rtm_datos[4] else "No disponible"
            return "No disponible"
        except Exception as e:
            return "Error en consulta"
    
    def guardar_o_actualizar_en_datos_runt(self, resultado: Dict, tipo: str = "soat"):
        """Guarda o actualiza el resultado en la hoja Datos Runt"""
        try:
            worksheet = self.sheet.worksheet(SHEET_DATOS_RUNT)
            todas_filas = worksheet.get_all_values()
            
            placa_buscar = resultado.get("placa", "").upper()
            numero_fila_existente = None
            datos_existentes = None
            
            for idx, fila_exist in enumerate(todas_filas[1:], start=2):
                if len(fila_exist) > 2:
                    placa_en_fila = fila_exist[2].strip().upper() if fila_exist[2] else ""
                    if placa_en_fila == placa_buscar:
                        numero_fila_existente = idx
                        datos_existentes = fila_exist
                        break
            
            if tipo == "soat":
                soat_data_nuevos = resultado.get("datos_soat", ["No disponible"] * 7)
                
                if numero_fila_existente and datos_existentes:
                    rtm_data_actuales = datos_existentes[11:18] if len(datos_existentes) > 18 else ["No disponible"] * 7
                    fila_completa = [
                        resultado.get("Tiempo ejecucion", datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                        resultado.get("cedula", "No disponible"),
                        resultado.get("placa", "No disponible"),
                        resultado.get("cilindraje", "No disponible"),
                    ] + soat_data_nuevos + rtm_data_actuales + [resultado.get("estado", "Pendiente")]
                    rango = f"A{numero_fila_existente}:S{numero_fila_existente}"
                    worksheet.update([fila_completa], range_name=rango, value_input_option="RAW")
                else:
                    fila_completa = [
                        resultado.get("Tiempo ejecucion", datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                        resultado.get("cedula", "No disponible"),
                        resultado.get("placa", "No disponible"),
                        resultado.get("cilindraje", "No disponible"),
                    ] + soat_data_nuevos + ["No disponible"] * 7 + [resultado.get("estado", "Pendiente")]
                    fila_nueva = len(todas_filas) + 1
                    rango = f"A{fila_nueva}:S{fila_nueva}"
                    worksheet.update([fila_completa], range_name=rango, value_input_option="RAW")
                    
            else:
                rtm_data_nuevos = resultado.get("datos_técnicos", ["No disponible"] * 7)
                
                if numero_fila_existente and datos_existentes:
                    soat_data_actuales = datos_existentes[4:11] if len(datos_existentes) > 11 else ["No disponible"] * 7
                    fila_completa = [
                        resultado.get("Tiempo ejecucion", datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                        resultado.get("cedula", "No disponible"),
                        resultado.get("placa", "No disponible"),
                        resultado.get("cilindraje", "No disponible"),
                    ] + soat_data_actuales + rtm_data_nuevos + [resultado.get("estado", "Pendiente")]
                    rango = f"A{numero_fila_existente}:S{numero_fila_existente}"
                    worksheet.update([fila_completa], range_name=rango, value_input_option="RAW")
                else:
                    fila_completa = [
                        resultado.get("Tiempo ejecucion", datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                        resultado.get("cedula", "No disponible"),
                        resultado.get("placa", "No disponible"),
                        resultado.get("cilindraje", "No disponible"),
                    ] + ["No disponible"] * 7 + rtm_data_nuevos + [resultado.get("estado", "Pendiente")]
                    fila_nueva = len(todas_filas) + 1
                    rango = f"A{fila_nueva}:S{fila_nueva}"
                    worksheet.update([fila_completa], range_name=rango, value_input_option="RAW")
            
        except Exception as e:
            logging.error(f"❌ Error guardando en Datos Runt: {e}")
    
    def procesar_consulta_vehiculo(self, cedula: str, placa: str, cedula_propietario: str = None) -> Tuple[Optional[Dict], bool]:
        """Procesa una consulta de vehículo con reintento usando propietario si falla"""
        if not self.driver:
            return None, False
        
        try:
            # Primer intento con cédula asociado
            resultado, _ = procesar_consulta_interno(
                self.driver, cedula, placa, 0,
                es_reintento=False, max_intentos_internos=2
            )
            
            if resultado and resultado.get("estado") == "Exitoso":
                return resultado, True
            
            # Si falló por "Sin personas asociadas" y tenemos cédula propietario
            elif resultado and resultado.get("estado") == "Exitoso - Sin personas asociadas" and cedula_propietario:
                # Reiniciar para nueva consulta
                try:
                    self.driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                    time.sleep(3)
                    limpiar_todos_los_campos(self.driver)
                    time.sleep(1)
                except Exception as e:
                    return None, False
                
                # Segundo intento con cédula propietario
                resultado_prop, _ = procesar_consulta_interno(
                    self.driver, cedula_propietario, placa, 0,
                    es_reintento=True, max_intentos_internos=1
                )
                
                if resultado_prop and resultado_prop.get("estado") == "Exitoso":
                    return resultado_prop, True
                return resultado_prop, False
            
            return resultado, False
                
        except Exception as e:
            logging.error(f"❌ Error en consulta: {e}")
            return None, False
    
    def reiniciar_sesion_periodica(self):
        """Reinicia la sesión periódicamente cada 5 consultas"""
        if self.contador_consultas_ciclo > 0 and self.contador_consultas_ciclo % 5 == 0 and self.driver:
            try:
                self.driver.get("https://www.google.com")
                time.sleep(2)
                self.driver.get("https://portalpublico.runt.gov.co/#/consulta-vehiculo/consulta/consulta-ciudadana")
                time.sleep(4)
                limpiar_todos_los_campos(self.driver)
                time.sleep(2)
            except Exception as e:
                logging.error(f"❌ Error reiniciando sesión: {e}")
    
    def procesar_tipo_vigencia(self, sheet_name: str, tipo: str) -> Tuple[int, int]:
        """
        Procesa un tipo específico de vigencia (SOAT o RTM)
        
        Returns:
            Tupla (exitosos, fallos)
        """
        registros = self.leer_registros_desde_vigencias(sheet_name, tipo)
        
        if not registros:
            logging.info(f"📊 {tipo.upper()}: No hay registros pendientes")
            return 0, 0
        
        logging.info(f"📊 {tipo.upper()}: {len(registros)} registros para procesar")
        
        exitosos = 0
        fallos = 0
        
        for cedula_asociado, cedula_propietario, placa, fila_origen in registros:
            if not self.ejecutando:
                return exitosos, fallos
            
            self.contador_consultas_ciclo += 1
            
            # Reiniciar sesión periódicamente
            self.reiniciar_sesion_periodica()
            
            logging.info(f"🔍 [{self.contador_consultas_ciclo}] {tipo.upper()} - {placa}")
            
            resultado, exito = self.procesar_consulta_vehiculo(cedula_asociado, placa, cedula_propietario)
            
            if exito and resultado:
                nuevo_estado = self.obtener_vigencia_actual(resultado, tipo)
                self.guardar_o_actualizar_en_datos_runt(resultado, tipo)
                
                # Guardar en estado propio
                self.estado[tipo]["procesadas"][placa] = {
                    "estado": "exitoso",
                    "fecha": datetime.now().isoformat(),
                    "vigencia": nuevo_estado,
                    "ciclo": self.contador_ciclos + 1
                }
                self.estado[tipo]["total_exitosas"] += 1
                self.estado["estadisticas"]["total_consultas"] += 1
                
                exitosos += 1
                
                # Log en archivo separado
                exitosos_logger.info(f"{tipo.upper()} - {placa} - Vigencia: {nuevo_estado}")
                logging.info(f"✅ {tipo.upper()} - {placa} EXITOSO - Vigencia: {nuevo_estado}")
                
            else:
                # Guardar fallo en estado
                self.estado[tipo]["procesadas"][placa] = {
                    "estado": "fallido",
                    "fecha": datetime.now().isoformat(),
                    "ciclo": self.contador_ciclos + 1
                }
                self.estado[tipo]["total_fallidas"] += 1
                self.estado["estadisticas"]["total_consultas"] += 1
                
                fallos += 1
                
                # Log en archivo separado
                errores_logger.info(f"{tipo.upper()} - {placa} - FALLÓ")
                logging.error(f"❌ {tipo.upper()} - {placa} FALLÓ")
            
            # Guardar estado después de cada consulta
            guardar_estado_vigencias(self.estado)
            
            time.sleep(PAUSA_CORTA)
        
        return exitosos, fallos
    
    def procesar_ciclo_completo(self):
        """Procesa un ciclo completo (SOAT + Tecnomecanica)"""
        ciclo_numero = self.contador_ciclos + 1
        
        logging.info(f"\n{'='*80}")
        logging.info(f"🔄 INICIANDO CICLO #{ciclo_numero}")
        logging.info(f"📅 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logging.info(f"{'='*80}")
        
        # Reiniciar contador de consultas del ciclo
        self.contador_consultas_ciclo = 0
        
        # Procesar SOAT
        logging.info(f"\n📌 FASE 1: VIGENCIAS SOAT")
        exitosos_soat, fallos_soat = self.procesar_tipo_vigencia(SHEET_SOAT, "soat")
        
        # Pausa entre fases
        time.sleep(5)
        
        # Procesar Tecnomecanica
        logging.info(f"\n📌 FASE 2: VIGENCIAS TECNOMECANICA")
        exitosos_rtm, fallos_rtm = self.procesar_tipo_vigencia(SHEET_TECNOMECANICA, "tecnomecanica")
        
        # Actualizar estadísticas del ciclo
        self.estado["estadisticas"]["total_ciclos"] = ciclo_numero
        self.estado["estadisticas"]["fecha_ultimo_ciclo"] = datetime.now().isoformat()
        if self.estado["estadisticas"]["fecha_inicio"] is None:
            self.estado["estadisticas"]["fecha_inicio"] = datetime.now().isoformat()
        
        self.estado["soat"]["ultima_ejecucion"] = datetime.now().isoformat()
        self.estado["soat"]["ultimo_ciclo"] = ciclo_numero
        self.estado["tecnomecanica"]["ultima_ejecucion"] = datetime.now().isoformat()
        self.estado["tecnomecanica"]["ultimo_ciclo"] = ciclo_numero
        
        guardar_estado_vigencias(self.estado)
        
        # Log de resumen del ciclo
        resumen = (f"CICLO #{ciclo_numero} - SOAT: ✅{exitosos_soat} ❌{fallos_soat} | "
                   f"RTM: ✅{exitosos_rtm} ❌{fallos_rtm} | "
                   f"Total consultas: {self.contador_consultas_ciclo}")
        ciclos_logger.info(resumen)
        
        logging.info(f"\n{'='*80}")
        logging.info(f"✅ CICLO #{ciclo_numero} COMPLETADO")
        logging.info(f"   SOAT: ✅ {exitosos_soat} | ❌ {fallos_soat}")
        logging.info(f"   RTM:  ✅ {exitosos_rtm} | ❌ {fallos_rtm}")
        logging.info(f"   Total consultas: {self.contador_consultas_ciclo}")
        logging.info(f"{'='*80}\n")
        
        self.contador_ciclos = ciclo_numero
    
    def ejecutar_bucle_infinito(self):
        """Ejecuta el bucle infinito de procesamiento"""
        logging.info("="*80)
        logging.info("🚀 PROCESADOR DE VIGENCIAS - MODO BUCLE INFINITO")
        logging.info(f"   📁 Archivo de estado propio: {ESTADO_VIGENCIAS_FILE}")
        logging.info(f"   📁 Logs en carpeta: logs/")
        logging.info(f"   ⏱️  Pausa entre ciclos: {PAUSA_ENTRE_CICLOS} segundos")
        logging.info(f"   🔄 Reinicio de driver cada: {REINICIO_DRIVER_CADA_CICLOS} ciclos")
        logging.info("   🛑 Presiona Ctrl+C para detener")
        logging.info("="*80)
        
        # Mostrar estadísticas actuales
        logging.info(f"\n📊 ESTADÍSTICAS ACTUALES:")
        logging.info(f"   SOAT - Exitosas: {self.estado['soat']['total_exitosas']} | Fallidas: {self.estado['soat']['total_fallidas']}")
        logging.info(f"   RTM  - Exitosas: {self.estado['tecnomecanica']['total_exitosas']} | Fallidas: {self.estado['tecnomecanica']['total_fallidas']}")
        logging.info(f"   Total ciclos completados: {self.estado['estadisticas']['total_ciclos']}")
        logging.info(f"   Total consultas realizadas: {self.estado['estadisticas']['total_consultas']}\n")
        
        # Conectar a Google Sheets
        if not self.conectar_google_sheets():
            logging.error("❌ No se pudo conectar a Google Sheets")
            return
        
        # Iniciar driver
        if not self.reiniciar_driver():
            logging.error("❌ No se pudo iniciar el driver")
            return
        
        try:
            while self.ejecutando:
                try:
                    # Procesar un ciclo completo
                    self.procesar_ciclo_completo()
                    
                    # Reiniciar driver cada X ciclos
                    if self.contador_ciclos % REINICIO_DRIVER_CADA_CICLOS == 0:
                        logging.info(f"\n🔄 Reiniciando driver después de {self.contador_ciclos} ciclos...")
                        self.reiniciar_driver()
                    
                    if self.ejecutando:
                        logging.info(f"💤 Esperando {PAUSA_ENTRE_CICLOS} segundos antes del próximo ciclo...")
                        
                        # Esperar con cuenta regresiva
                        for i in range(PAUSA_ENTRE_CICLOS, 0, -30):
                            if not self.ejecutando:
                                break
                            if i % 300 == 0 or i <= 60:  # Cada 5 minutos o último minuto
                                minutos = i // 60
                                segundos = i % 60
                                logging.info(f"   ⏳ Próximo ciclo en {minutos}m {segundos}s")
                            time.sleep(30)
                        
                except Exception as e:
                    logging.error(f"❌ Error en ciclo: {e}")
                    logging.info("🔄 Reintentando en 60 segundos...")
                    time.sleep(60)
                    
                    # Intentar reconectar
                    try:
                        self.reiniciar_driver()
                        self.conectar_google_sheets()
                    except:
                        pass
                    
        except KeyboardInterrupt:
            logging.info("\n⚠️ Proceso detenido por el usuario")
        finally:
            self.estado["estadisticas"]["fecha_ultimo_ciclo"] = datetime.now().isoformat()
            guardar_estado_vigencias(self.estado)
            if self.driver:
                cerrar_driver(self.driver)
            logging.info("✅ Proceso finalizado. Estado guardado.")


def main():
    """Función principal - Inicia el bucle infinito"""
    processor = VigenciaProcessor()
    processor.ejecutar_bucle_infinito()


if __name__ == "__main__":
    main()