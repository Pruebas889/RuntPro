import logging
import gspread
from google.oauth2.service_account import Credentials
from pathlib import Path
from datetime import datetime

# ═════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ═════════════════════════════════════════════════════════════

BASE_PATH = Path(r"C:\Users\cmarroquin\Music\RuntPro")
GOOGLE_CREDS = BASE_PATH / "prueba-de-gmail-486215-345473339c47.json"
LOGS_FOLDER = BASE_PATH / "Log_duplicados"
LOGS_FOLDER.mkdir(parents=True, exist_ok=True)


# ═════════════════════════════════════════════════════════════
# LOGGING
# ═════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(str(LOGS_FOLDER / "limpiar_duplicados.log"), encoding='utf-8')
    ]
)

# ═════════════════════════════════════════════════════════════
# FUNCIONES DE LIMPIEZA
# ═════════════════════════════════════════════════════════════

def conectar_sheets(sheet_id):
    """Conecta con Google Sheets"""
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
        sheet = client.open_by_key(sheet_id)
        return sheet

    except Exception as e:
        logging.error(f"❌ Error conectando a Sheets: {e}")
        return None


def limpiar_duplicados_datos_runt():
    """
    🧹 Limpia duplicados en 'Datos Runt'
    - Clave de duplicado: Columna C (PLACA) + Columna B (CÉDULA)
    - Mantiene el PRIMERO, elimina los posteriores
    """
    
    logging.info("\n" + "="*70)
    logging.info("🧹 LIMPIANDO DUPLICADOS EN 'Datos Runt'")
    logging.info("="*70)
    
    sheet = conectar_sheets("1vs414iH3QVeLoTcY2CExg4kD9eCkXZRRfax_WTlUXPk")
    if not sheet:
        return
    
    try:
        worksheet = sheet.worksheet("Datos Runt")
        logging.info("✅ Hoja 'Datos Runt' encontrada")
    except gspread.WorksheetNotFound:
        logging.error("❌ Hoja 'Datos Runt' no encontrada")
        return
    
    # Obtener todas las filas
    todas_filas = worksheet.get_all_values()
    logging.info(f"📊 Total de filas: {len(todas_filas)}")
    
    if len(todas_filas) <= 1:
        logging.info("ℹ️  Solo hay encabezados, no hay datos por limpiar")
        return
    
    # Separar encabezado y datos
    encabezado = todas_filas[0]
    datos = todas_filas[1:]
    
    logging.info(f"📋 Encabezado: {encabezado}")
    logging.info(f"📊 Datos: {len(datos)} registros")
    
    # ═══ IDENTIFICAR DUPLICADOS ═══
    # Clave: Columna B (índice 1) = cédula, Columna C (índice 2) = placa
    vistos = {}
    filas_a_eliminar = []
    
    for idx, fila in enumerate(datos, start=2):  # Comienza en fila 2 (después del encabezado)
        if len(fila) < 3:
            continue
        
        cedula = fila[1].strip() if len(fila) > 1 else ""
        placa = fila[2].strip() if len(fila) > 2 else ""
        
        clave = f"{cedula}|{placa}"
        
        if clave in vistos:
            # Es duplicado
            logging.warning(f"   ❌ DUPLICADO encontrado en fila {idx}: {placa} / {cedula}")
            filas_a_eliminar.append(idx)
        else:
            # Primera vez que se ve
            logging.info(f"   ✓ Fila {idx}: {placa} / {cedula} (MANTENER)")
            vistos[clave] = idx
    
    # ═══ ELIMINAR FILAS DUPLICADAS (BATCH) ═══
    if filas_a_eliminar:
        logging.warning(f"\n🔄 ELIMINANDO {len(filas_a_eliminar)} filas duplicadas...")
        
        # ⭐ USAR BATCH DELETE EN LUGAR DE DELETE_ROWS UNO POR UNO
        try:
            # Convertir números de fila a índices de Google Sheets (0-based, sin encabezado)
            indices_a_eliminar = [num_fila - 1 for num_fila in sorted(filas_a_eliminar, reverse=True)]
            
            logging.info(f"   📋 Indices a eliminar: {indices_a_eliminar}")
            
            # Crear requests de delete
            requests = []
            for index in indices_a_eliminar:
                requests.append({
                    "deleteDimension": {
                        "range": {
                            "sheetId": worksheet.id,
                            "dimension": "ROWS",
                            "startIndex": index,
                            "endIndex": index + 1
                        }
                    }
                })
            
            # Ejecutar batch update
            if requests:
                response = worksheet.client.batch_update(worksheet.spreadsheet_id, {"requests": requests})
                logging.info(f"✅ Se eliminaron {len(filas_a_eliminar)} duplicados (batch)")
                
                for num_fila in filas_a_eliminar:
                    logging.info(f"   🗑️  Eliminada fila {num_fila}")
        
        except Exception as e:
            logging.error(f"❌ Error en batch delete: {e}")
            logging.warning("🔄 Intentando método alternativo (uno por uno)...")
            
            # Fallback: Eliminar de atrás para adelante
            for num_fila in sorted(filas_a_eliminar, reverse=True):
                try:
                    worksheet.delete_row(num_fila)
                    logging.info(f"   🗑️  Eliminada fila {num_fila}")
                except Exception as e2:
                    logging.error(f"   ❌ Error eliminando fila {num_fila}: {e2}")
    else:
        logging.info("✅ No se encontraron duplicados en 'Datos Runt'")


def limpiar_duplicados_datos_vehiculo():
    """
    🧹 Limpia duplicados en 'Datos Vehiculo'
    - Clave de duplicado: Columna A (PLACA)
    - Mantiene el PRIMERO, elimina los posteriores
    """
    
    logging.info("\n" + "="*70)
    logging.info("🧹 LIMPIANDO DUPLICADOS EN 'Datos Vehiculo'")
    logging.info("="*70)
    
    sheet = conectar_sheets("1vs414iH3QVeLoTcY2CExg4kD9eCkXZRRfax_WTlUXPk")
    if not sheet:
        return
    
    try:
        worksheet = sheet.worksheet("Datos Vehiculo")
        logging.info("✅ Hoja 'Datos Vehiculo' encontrada")
    except gspread.WorksheetNotFound:
        logging.error("❌ Hoja 'Datos Vehiculo' no encontrada")
        return
    
    # Obtener todas las filas
    todas_filas = worksheet.get_all_values()
    logging.info(f"📊 Total de filas: {len(todas_filas)}")
    
    if len(todas_filas) <= 1:
        logging.info("ℹ️  Solo hay encabezados, no hay datos por limpiar")
        return
    
    # Separar encabezado y datos
    encabezado = todas_filas[0]
    datos = todas_filas[1:]
    
    logging.info(f"📋 Encabezado: {encabezado[:5]}...")
    logging.info(f"📊 Datos: {len(datos)} registros")
    
    # ═══ IDENTIFICAR DUPLICADOS ═══
    # Clave: Columna A (índice 0) = placa
    vistos = {}
    filas_a_eliminar = []
    
    for idx, fila in enumerate(datos, start=2):  # Comienza en fila 2 (después del encabezado)
        if len(fila) < 1 or not fila[0].strip():
            continue
        
        placa = fila[0].strip()
        
        if placa in vistos:
            # Es duplicado
            logging.warning(f"   ❌ DUPLICADO encontrado en fila {idx}: {placa}")
            filas_a_eliminar.append(idx)
        else:
            # Primera vez que se ve
            logging.info(f"   ✓ Fila {idx}: {placa} (MANTENER)")
            vistos[placa] = idx
    
        # ═══ ELIMINAR FILAS DUPLICADAS (BATCH) ═══
    if filas_a_eliminar:
        logging.warning(f"\n🔄 ELIMINANDO {len(filas_a_eliminar)} filas duplicadas...")
        
        try:
            indices_a_eliminar = [num_fila - 1 for num_fila in sorted(filas_a_eliminar, reverse=True)]
            
            requests = []
            for index in indices_a_eliminar:
                requests.append({
                    "deleteDimension": {
                        "range": {
                            "sheetId": worksheet.id,
                            "dimension": "ROWS",
                            "startIndex": index,
                            "endIndex": index + 1
                        }
                    }
                })
            
            if requests:
                response = worksheet.client.batch_update(worksheet.spreadsheet_id, {"requests": requests})
                logging.info(f"✅ Se eliminaron {len(filas_a_eliminar)} duplicados (batch)")
                
                for num_fila in filas_a_eliminar:
                    logging.info(f"   🗑️  Eliminada fila {num_fila}")
        
        except Exception as e:
            logging.error(f"❌ Error en batch delete: {e}")
            logging.warning("🔄 Intentando método alternativo...")
            
            for num_fila in sorted(filas_a_eliminar, reverse=True):
                try:
                    worksheet.delete_row(num_fila)
                    logging.info(f"   🗑️  Eliminada fila {num_fila}")
                except Exception as e2:
                    logging.error(f"   ❌ Error eliminando fila {num_fila}: {e2}")
    else:
        logging.info("✅ No se encontraron duplicados en 'Datos Vehiculo'")


def limpiar_duplicados_resultados():
    """
    🧹 Limpia duplicados en 'Resultados'
    - Clave de duplicado: Columna A (CÉDULA ASOCIADO) + Columna C (PLACA)
    - Mantiene el PRIMERO, elimina los posteriores
    """
    
    logging.info("\n" + "="*70)
    logging.info("🧹 LIMPIANDO DUPLICADOS EN 'Resultados'")
    logging.info("="*70)
    
    sheet = conectar_sheets("1vs414iH3QVeLoTcY2CExg4kD9eCkXZRRfax_WTlUXPk")
    if not sheet:
        return
    
    try:
        worksheet = sheet.worksheet("Resultados")
        logging.info("✅ Hoja 'Resultados' encontrada")
    except gspread.WorksheetNotFound:
        logging.error("❌ Hoja 'Resultados' no encontrada")
        return
    
    # Obtener todas las filas
    todas_filas = worksheet.get_all_values()
    logging.info(f"📊 Total de filas: {len(todas_filas)}")
    
    if len(todas_filas) <= 1:
        logging.info("ℹ️  Solo hay encabezados, no hay datos por limpiar")
        return
    
    # Separar encabezado y datos
    encabezado = todas_filas[0]
    datos = todas_filas[1:]
    
    logging.info(f"📋 Encabezado: {encabezado}")
    logging.info(f"📊 Datos: {len(datos)} registros")
    
    # ═══ IDENTIFICAR DUPLICADOS ═══
    # Clave: Columna A (índice 0) = cédula asociado, Columna C (índice 2) = placa
    vistos = {}
    filas_a_eliminar = []
    
    for idx, fila in enumerate(datos, start=2):  # Comienza en fila 2 (después del encabezado)
        if len(fila) < 3:
            continue
        
        cedula_asociado = fila[0].strip() if len(fila) > 0 else ""
        placa = fila[2].strip() if len(fila) > 2 else ""
        
        if not cedula_asociado or not placa:
            continue
        
        clave = f"{cedula_asociado}|{placa}"
        
        if clave in vistos:
            # Es duplicado
            logging.warning(f"   ❌ DUPLICADO encontrado en fila {idx}: {placa} / {cedula_asociado}")
            filas_a_eliminar.append(idx)
        else:
            # Primera vez que se ve
            logging.info(f"   ✓ Fila {idx}: {placa} / {cedula_asociado} (MANTENER)")
            vistos[clave] = idx
    
        # ═══ ELIMINAR FILAS DUPLICADAS (BATCH) ═══
    if filas_a_eliminar:
        logging.warning(f"\n🔄 ELIMINANDO {len(filas_a_eliminar)} filas duplicadas...")
        
        try:
            indices_a_eliminar = [num_fila - 1 for num_fila in sorted(filas_a_eliminar, reverse=True)]
            
            requests = []
            for index in indices_a_eliminar:
                requests.append({
                    "deleteDimension": {
                        "range": {
                            "sheetId": worksheet.id,
                            "dimension": "ROWS",
                            "startIndex": index,
                            "endIndex": index + 1
                        }
                    }
                })
            
            if requests:
                response = worksheet.client.batch_update(worksheet.spreadsheet_id, {"requests": requests})
                logging.info(f"✅ Se eliminaron {len(filas_a_eliminar)} duplicados (batch)")
                
                for num_fila in filas_a_eliminar:
                    logging.info(f"   🗑️  Eliminada fila {num_fila}")
        
        except Exception as e:
            logging.error(f"❌ Error en batch delete: {e}")
            logging.warning("🔄 Intentando método alternativo...")
            
            for num_fila in sorted(filas_a_eliminar, reverse=True):
                try:
                    worksheet.delete_row(num_fila)
                    logging.info(f"   🗑️  Eliminada fila {num_fila}")
                except Exception as e2:
                    logging.error(f"   ❌ Error eliminando fila {num_fila}: {e2}")
    else:
        logging.info("✅ No se encontraron duplicados en 'Resultados'")


def main():
    """Función principal"""
    
    logging.info("\n" + "="*70)
    logging.info("🧹 INICIANDO LIMPIEZA DE DUPLICADOS")
    logging.info(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("="*70)
    
    # Limpiar las 3 hojas
    limpiar_duplicados_datos_runt()
    limpiar_duplicados_datos_vehiculo()
    limpiar_duplicados_resultados()
    
    # ═══ REPORTE FINAL ═══
    logging.info("\n" + "="*70)
    logging.info("✅ LIMPIEZA COMPLETADA")
    logging.info("="*70)
    logging.info("📋 Se procesaron:")
    logging.info("   ✓ Datos Runt")
    logging.info("   ✓ Datos Vehiculo")
    logging.info("   ✓ Resultados")
    logging.info("="*70 + "\n")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.error(f"❌ Error: {e}", exc_info=True)