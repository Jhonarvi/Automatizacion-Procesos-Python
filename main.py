import os
import sys
import time

import gspread
import pandas as pd
from gspread.utils import rowcol_to_a1
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Variables de Ruta Globales
user_profile = os.path.expanduser("~")
RUTA_ESCRITORIO = os.path.join(user_profile, "Desktop") 
RUTA_DESCARGAS = os.path.join(RUTA_ESCRITORIO, "Automatizacion_TEMP")

# =========================================================================
# --- CONFIGURACIÓN DE ACCESO Y ARCHIVOS ---
# =========================================================================
URL_LOGIN = "https://campus.talentotechvalle.co/auth/login/monitor"
USUARIO = "CORREO_DEL_USUARIO"  # REEMPLAZA ESTO CON EL CORREO REAL
CONTRASENA = "CONTRASEÑA_DEL_USUARIO"  # REEMPLAZA ESTO CON LA CONTRASEÑA REAL
URL_DASHBOARD = "https://campus.talentotechvalle.co/dashboard/monitor"

GOOGLE_SHEET_ID = 'ID_DE_TU_GOOGLE_SHEET'  # REEMPLAZA ESTO CON EL ID REAL DE TU GOOGLE SHEET (la parte larga en la URL de tu hoja)
CREDENCIALES_JSON = 'CREDENCIALES_GOOGLE.json'  # REEMPLAZA ESTO CON EL NOMBRE REAL DE TU ARCHIVO JSON DE CREDENCIALES GENERADO EN GOOGLE CLOUD 
ARCHIVO_EXCEL_DESCARGADO = os.path.join(RUTA_DESCARGAS, 'informe_asistencias.xlsx')

SELECTORES = {
    'usuario': (By.NAME, "email"),
    'contrasena': (By.NAME, "password"),
    'boton_login': (By.XPATH, "//button[contains(., 'Ingresar')]"),
    'enlace_bootcamps': (By.XPATH, "//a[contains(normalize-space(.), 'Ir a bootcamps')]"),
    'checkbox_despliegue': (By.XPATH, "//input[@type='checkbox' and @name='IAP1-221']"),
    'enlace_sesiones': (By.XPATH, "//a[contains(text(), 'Sesiones')]"),
    'boton_reporte': (By.XPATH, "//button[contains(text(), 'Generar informe de asistencia')]")
}
# =========================================================================

def wait_for_file_to_be_ready(filepath, timeout=60, delay=5): 
    """Espera hasta que el archivo esté completamente descargado y liberado."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            with open(filepath, 'rb') as f:
                print("-> ÉXITO: Archivo descargado y liberado por el sistema.")
                return True 
        except IOError:
            print(f"-> Archivo bloqueado o incompleto. Esperando {delay}s...")
            time.sleep(delay)
        except FileNotFoundError:
            print(f"-> Archivo no encontrado. Esperando {delay}s...")
            time.sleep(delay)
    print(f"ERROR: Tiempo de espera agotado ({timeout}s). El archivo sigue bloqueado o no existe.")
    return False

class AsistenciaAgent:
    def __init__(self, url_login, usuario, contrasena, url_dashboard, selectores):
        self.url_login = url_login
        self.usuario = usuario
        self.contrasena = contrasena
        self.url_dashboard = url_dashboard
        self.selectores = selectores
        self.driver = None 
        self.ventana_principal = None

    def iniciar_navegador(self):
        """Abre el navegador Chrome, configura la descarga y maximiza la ventana."""
        print("1. Inicializando navegador...")
        
        # Crear la carpeta temporal
        os.makedirs(RUTA_DESCARGAS, exist_ok=True)
        
        # Configuración de Chrome personalizado
        chrome_options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": RUTA_DESCARGAS,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True 
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        self.driver = webdriver.Chrome(options=chrome_options) 
        self.driver.maximize_window() 
        
        self.driver.get(self.url_login)
        self.ventana_principal = self.driver.current_window_handle
        print("2. Navegador iniciado y maximizado.")
        return self.driver

    def ejecutar_login(self):
        """Realiza el login en la plataforma."""
        driver = self.driver
        if not driver: 
            return None
        
        try:
            print("3. Iniciando proceso de login...")
            wait = WebDriverWait(driver, 15)
            
            wait.until(EC.presence_of_element_located(self.selectores['usuario'])).send_keys(self.usuario)
            driver.find_element(*self.selectores['contrasena']).send_keys(self.contrasena)
            wait.until(EC.element_to_be_clickable(self.selectores['boton_login'])).click()
            wait.until(EC.url_to_be(self.url_dashboard))
            
            print("4. ÉXITO: Login completado. En el dashboard.")
            time.sleep(3) 
            return driver
            
        except Exception as e:
            print(f"5. ERROR en el proceso de login: {e}")
            self.cerrar_navegador()
            return None

    def navegar_y_descargar(self):
        """Navega por la plataforma y descarga el informe de asistencia."""
        driver = self.driver
        if not driver: 
            return False
        
        wait = WebDriverWait(driver, 30)
        
        try:
            print("5. Iniciando navegación hacia el reporte...")
            
            # Pre-limpieza del archivo anterior
            if os.path.exists(ARCHIVO_EXCEL_DESCARGADO):
                os.remove(ARCHIVO_EXCEL_DESCARGADO)
                print(f"-> Archivo anterior eliminado: {ARCHIVO_EXCEL_DESCARGADO}")
            
            # 5A. Clic en "Ir a bootcamps"
            print("-> Paso 5A: Clic en 'Ir a bootcamps'")
            enlace_bootcamps = wait.until(EC.element_to_be_clickable(self.selectores['enlace_bootcamps']))
            enlace_bootcamps.click()
            time.sleep(3)
            
            # 5B. Clic en checkbox de despliegue IAP1-221
            print("-> Paso 5B: Buscando checkbox IAP1-221...")
            checkbox_despliegue = wait.until(EC.presence_of_element_located(self.selectores['checkbox_despliegue']))
            
            checkbox_name = checkbox_despliegue.get_attribute('name')
            print(f"-> Checkbox encontrado: {checkbox_name}")
            
            if checkbox_name != "IAP1-221":
                raise Exception(f"ERROR: Se seleccionó el checkbox incorrecto: {checkbox_name}")
            
            driver.execute_script("""
                var element = arguments[0];
                var headerOffset = 150;
                var elementPosition = element.getBoundingClientRect().top;
                var offsetPosition = elementPosition + window.pageYOffset - headerOffset;
                window.scrollTo({
                    top: offsetPosition,
                    behavior: 'smooth'
                });
            """, checkbox_despliegue)
            time.sleep(2)
            
            driver.execute_script("arguments[0].click();", checkbox_despliegue)
            print("-> Menú IAP1-221 desplegado, esperando expansión...")
            time.sleep(3)
            
            print("-> Realizando segundo scroll para visualizar opciones del menú...")
            driver.execute_script("window.scrollBy(0, 250);")
            time.sleep(2)
            
            # 5C. Clic en enlace "Sesiones"
            print("-> Paso 5C: Buscando enlaces 'Sesiones' visibles...")
            enlaces_sesiones = driver.find_elements(*self.selectores['enlace_sesiones'])
            print(f"-> Se encontraron {len(enlaces_sesiones)} enlaces 'Sesiones'")
            
            enlace_sesiones = None
            for idx, enlace in enumerate(enlaces_sesiones):
                es_visible = enlace.is_displayed()
                if es_visible:
                    enlace_sesiones = enlace
                    print(f"-> Seleccionado enlace 'Sesiones' #{idx + 1} (visible)")
                    break
            
            if not enlace_sesiones:
                raise Exception("No se encontró ningún enlace 'Sesiones' visible")
            
            print("-> Haciendo clic en 'Sesiones'...")
            enlace_sesiones.click()
            time.sleep(1)
            
            self.ventana_principal = driver.current_window_handle
            print(f"-> Ventana principal guardada: {self.ventana_principal}")
            print("-> Clic realizado. Esperando nueva ventana...")
            time.sleep(1)
            
            # 5D. Cambio a la nueva ventana
            wait.until(EC.number_of_windows_to_be(2))
            time.sleep(2)
            
            for ventana in driver.window_handles:
                if ventana != self.ventana_principal:
                    print(f"-> Cambiando a la ventana secundaria: {ventana}")
                    driver.switch_to.window(ventana)
                    break
            
            time.sleep(3)
            
            # 5E. Esperar y hacer clic en el botón de reporte
            print("-> Paso 5E: Esperando botón 'Generar informe de asistencia'")
            boton_reporte = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(self.selectores['boton_reporte'])
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", boton_reporte)
            time.sleep(1)
            boton_reporte.click()
            print("6. Clic en Generar Reporte realizado. Descargando...")
            
            if not wait_for_file_to_be_ready(ARCHIVO_EXCEL_DESCARGADO):
                raise Exception("El archivo no pudo ser descargado o liberado a tiempo.")
            
            print("7. ÉXITO: Archivo descargado correctamente.")
            return True 

        except Exception as e:
            print(f"ERROR durante la navegación/descarga: {e}")
            try:
                if driver and len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(self.ventana_principal)
            except:
                pass
            return False

    def cerrar_navegador(self):
        """Cierra el navegador si está activo."""
        if self.driver:
            try:
                self.driver.quit()
                print("13. Navegador cerrado.")
            except:
                print("13. Error al cerrar navegador (puede ya estar cerrado).")


def procesar_sesion_individual(session_number_to_process, xls, worksheet):
    """Procesa y actualiza una única sesión."""
    try:
        # 1. Encontrar la pestaña de Excel para la sesión
        target_sheet_in_excel = None
        sheet_names = xls.sheet_names
        
        for sheet_name in sheet_names:
            try:
                partes = sheet_name.split('-')
                session_str = partes[0].replace('Sesión', '').strip()
                if int(session_str) == session_number_to_process:
                    target_sheet_in_excel = sheet_name
                    break
            except (ValueError, IndexError):
                continue 

        if not target_sheet_in_excel:
            print(f"-> ADVERTENCIA: No se encontró una pestaña en Excel para la sesión {session_number_to_process}. Saltando...")
            return

        print(f"-> Pestaña de Excel encontrada: '{target_sheet_in_excel}'")

        # 2. Extraer datos de asistencia
        df_asistencia = pd.read_excel(xls, sheet_name=target_sheet_in_excel, usecols='J', header=0)
        asistencia_data = df_asistencia.iloc[0:33].values.flatten().tolist()
        print(f"-> {len(asistencia_data)} registros de asistencia extraídos del Excel.")

        # VALIDACIÓN: Convertir valores de 3 a 4
        asistencia_data = [4 if x == 3 else x for x in asistencia_data]
        
        # 3. Encontrar columna en Google Sheets
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            raise Exception("La hoja de Google no tiene la fila 2 para los números de sesión.")

        session_row = all_values[1]
        target_col_index = -1
        for i, cell_value in enumerate(session_row):
            if cell_value.strip() == str(session_number_to_process):
                target_col_index = i + 1 
                break
        
        if target_col_index == -1:
            print(f"-> ADVERTENCIA: No se pudo encontrar la columna para la sesión '{session_number_to_process}' en Google Sheet. Saltando...")
            return

        # 4. Actualizar la columna
        start_cell = rowcol_to_a1(3, target_col_index)
        end_cell = rowcol_to_a1(35, target_col_index)
        target_range = f'{start_cell}:{end_cell}'
        
        data_to_upload = [[val] for val in asistencia_data]
        worksheet.update(values=data_to_upload, range_name=target_range, value_input_option='USER_ENTERED')
        print(f"-> ÉXITO: Datos de la sesión {session_number_to_process} cargados en Google Sheets.")

    except Exception as e:
        print(f"-> ERROR procesando la sesión {session_number_to_process}: {e}")

def determinar_sesion_actual(xls):
    """(ANTIGUA) Encuentra el número de sesión más alto correspondiente a la fecha de hoy."""
    from datetime import datetime
    fecha_hoy_str = datetime.now().strftime('%Y-%m-%d')
    print(f"-> Buscando la sesión de hoy ({fecha_hoy_str}) para determinar el final del rango...")

    sheet_names = xls.sheet_names
    sheets_de_hoy = [name for name in sheet_names if fecha_hoy_str in name]

    if not sheets_de_hoy:
        raise Exception(f"No se encontró ninguna pestaña en el Excel con la fecha de hoy ({fecha_hoy_str}).")

    sesion_actual = -1
    for sheet_name in sheets_de_hoy:
        try:
            partes = sheet_name.split('-')
            session_str = partes[0].replace('Sesión', '').strip()
            session_num = int(session_str)
            if session_num > sesion_actual:
                sesion_actual = session_num
        except (ValueError, IndexError):
            continue
    
    if sesion_actual == -1:
        raise Exception(f"De las pestañas encontradas, ninguna tenía un formato de sesión válido.")
    
    return sesion_actual

# --- NUEVA FUNCIÓN AGREGADA ---
def determinar_ultima_sesion_absoluta(xls):
    """
    (NUEVA) Encuentra el número de la última sesión en todo el archivo Excel,
    ignorando completamente las fechas. Ideal para cuando el bootcamp ya finalizó.
    """
    print("-> Buscando la última clase registrada en todo el Excel (Fin del bootcamp)...")
    max_session = 0
    sheet_names = xls.sheet_names
    
    for sheet_name in sheet_names:
        try:
            partes = sheet_name.split('-')
            session_str = partes[0].replace('Sesión', '').strip()
            session_num = int(session_str)
            if session_num > max_session:
                max_session = session_num
        except (ValueError, IndexError):
            continue
            
    if max_session == 0:
        raise Exception("No se encontró ninguna pestaña con el formato 'Sesión X' en el Excel.")
        
    print(f"-> La última sesión absoluta encontrada es: {max_session}")
    return max_session
# ------------------------------

def procesar_multiples_sesiones(xls, worksheet, sesion_inicial, sesion_final):
    """Itera a través de un rango de sesiones y llama al procesador individual."""
    print(f"-> Rango de procesamiento definido: Sesión INICIAL={sesion_inicial}, Sesión FINAL={sesion_final}")

    if sesion_inicial > sesion_final:
        print(f"-> ADVERTENCIA: La sesión inicial ({sesion_inicial}) es mayor que la final ({sesion_final}).")
        return

    for session_number in range(sesion_inicial, sesion_final + 1):
        print("\n" + "="*50)
        print(f"--- PROCESANDO SESIÓN: {session_number} ---")
        procesar_sesion_individual(
            session_number_to_process=session_number,
            xls=xls,
            worksheet=worksheet
        )
        print(f"--- FIN DEL PROCESO PARA LA SESIÓN: {session_number} ---")
    
    print("\n" + "="*50)

# --- CÓDIGO DE EJECUCIÓN (MAIN) ---
if __name__ == "__main__":
    
    # --- MENÚ INTERACTIVO ---
    print("========================================")
    print("  AUTOMATIZACIÓN DE ASISTENCIA TALENTO")
    print("========================================")
    print("Por favor, seleccione el modo de ejecución:")
    print("  1. Actualizar TODAS las sesiones (desde la 1 hasta hoy)")
    print("  2. Actualizar SOLO la sesión de hoy")
    print("  3. Actualizar TODO el bootcamp (De la Sesión 1 hasta la última clase)")
    print("========================================")
    
    choice = input("Ingrese su opción (1, 2 o 3) y presione Enter: ").strip()
    
    if choice not in ['1', '2', '3']:
        print("Opción no válida. Saliendo del script.")
        sys.exit(1)

    agent = AsistenciaAgent(
        URL_LOGIN, USUARIO, CONTRASENA, URL_DASHBOARD, SELECTORES
    )
    
    driver_activo = agent.iniciar_navegador() 
    driver_activo = agent.ejecutar_login()
    
    if driver_activo:
        xls = None
        try:
            descarga_exitosa = agent.navegar_y_descargar()
            
            if descarga_exitosa:
                print("9. Iniciando procesamiento y carga a Google Sheets...")
                
                print("-> Abriendo Excel y conectando a Google Sheets...")
                for intento in range(10):
                    try:
                        xls = pd.ExcelFile(ARCHIVO_EXCEL_DESCARGADO)
                        break
                    except (IOError, FileNotFoundError, PermissionError) as e:
                        if intento == 9: raise Exception(f"El archivo Excel se mantuvo bloqueado: {e}")
                        time.sleep(2)
                
                if not xls: raise Exception("No se pudo abrir el archivo Excel.")

                if not os.path.exists(CREDENCIALES_JSON):
                    raise FileNotFoundError(f"No se encuentra el archivo de credenciales: {CREDENCIALES_JSON}")
                gc = gspread.service_account(filename=CREDENCIALES_JSON)
                spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
                worksheet = spreadsheet.worksheet('ASISTENCIA')
                print(f"-> Conectado a la hoja: '{worksheet.title}'")

                # --- 2. DETERMINAR RANGO DE SESIONES A PROCESAR SEGÚN LA OPCIÓN ---
                sesion_inicial = 1
                sesion_final = 1
                
                if choice == '1':
                    print("\n>>> MODO SELECCIONADO: Actualizar TODAS las sesiones hasta HOY.")
                    sesion_final = determinar_sesion_actual(xls)
                    sesion_inicial = 1
                elif choice == '2':
                    print("\n>>> MODO SELECCIONADO: Actualizar SOLO la sesión de HOY.")
                    sesion_final = determinar_sesion_actual(xls)
                    sesion_inicial = sesion_final
                elif choice == '3':
                    # ¡AQUÍ ESTÁ EL CAMBIO! Usamos la nueva función para la Opción 3
                    print("\n>>> MODO SELECCIONADO: Actualizar TODO EL BOOTCAMP (Desde la primera hasta la última clase).")
                    sesion_final = determinar_ultima_sesion_absoluta(xls)
                    sesion_inicial = 1

                # --- 3. LLAMAR AL PROCESADOR ---
                procesar_multiples_sesiones(xls, worksheet, sesion_inicial, sesion_final)

                # --- MENSAJE DE ÉXITO PERSONALIZADO ---
                if sesion_inicial <= sesion_final:
                    if choice == '1':
                        print(f"10. ÉXITO: Sesiones de la 1 a la {sesion_final} (Hoy) procesadas.")
                    elif choice == '2':
                        print(f"10. ÉXITO: La sesión de hoy ({sesion_final}) ha sido procesada.")
                    elif choice == '3':
                        print(f"10. ÉXITO: Todas las clases del bootcamp (1 a la {sesion_final}) fueron procesadas.")

                print("14. Proceso AUTOMATIZADO COMPLETADO.")

        except Exception as e:
            print(f"\n¡ALERTA FINAL! EL PROCESO FALLÓ COMPLETAMENTE: {e}")
            import traceback
            traceback.print_exc()
        
        finally:
            if xls is not None:
                xls.close()
            
            time.sleep(2)
            try:
                if os.path.exists(ARCHIVO_EXCEL_DESCARGADO):
                    os.remove(ARCHIVO_EXCEL_DESCARGADO)
                    print("12. Archivo Excel temporal eliminado.")
            except PermissionError:
                print("12. ADVERTENCIA: No se pudo eliminar el archivo Excel temporal.")

            agent.cerrar_navegador()
            
    else:
        print("El script terminó debido a un fallo en el login.")