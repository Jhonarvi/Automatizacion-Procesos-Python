import os
import pywhatkit
import pandas as pd
import time
import pyautogui
from google import genai

# ==========================================
# 1. CONFIGURACIÓN Y SEGURIDAD
# ==========================================
# ⚠️ NUNCA quemes la API Key en el código. Úsala como variable de entorno.
API_KEY = os.getenv("GEMINI_API_KEY", "TU_API_KEY_AQUI_SOLO_LOCAL")

# Coordenadas de contingencia para el botón enviar (Depende de la resolución del monitor)
COORDENADA_X = 0
COORDENADA_Y = 0

# Inicialización del cliente de Gemini
client = genai.Client(api_key=API_KEY)

# ==========================================
# 2. FUNCIONES AUXILIARES
# ==========================================

def barra_de_progreso(segundos, mensaje):
    """Muestra una cuenta regresiva visual en la consola para evitar bloqueos."""
    print(f"   ⏳ {mensaje}")
    for i in range(segundos, 0, -1):
        print(f"\r   ... {i} ...", end="")
        time.sleep(1)
    print("\r   ✅ Listo.                  ")

def enviar_recordatorio_whatsapp(numero_completo, texto):
    """Ejecuta la automatización del navegador y el teclado para enviar el mensaje."""
    try:
        print(f"   -> Abriendo WhatsApp para: {numero_completo}")
        print("   -> ⚠️ NO MUEVAS EL MOUSE. ESPERANDO 20 SEGUNDOS...")

        # Abrimos WhatsApp Web
        pywhatkit.sendwhatmsg_instantly(
            numero_completo,
            texto,
            wait_time=20,
            tab_close=False
        )

        barra_de_progreso(5, "Cargando el chat...")

        # --- FLUJO DE ENVÍO ---
        if COORDENADA_X > 0 and COORDENADA_Y > 0:
            print("   -> Ejecutando clic físico por coordenadas...")
            pyautogui.click(x=COORDENADA_X, y=COORDENADA_Y)
            time.sleep(0.5)
            pyautogui.click()
        else:
            print("   -> Ejecutando ENTER virtual...")
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.press('enter')

        print(f"   -> ✅ ¡Mensaje enviado con éxito!")

    except Exception as e:
        print(f"   -> ❌ Error en la interfaz de WhatsApp: {e}")

# ==========================================
# 3. FLUJO PRINCIPAL (ETL e IA)
# ==========================================

def main():
    try:
        # Lectura de la base de datos de campistas
        df_campistas = pd.read_excel('campistas.xlsx', dtype={'TELÉFONO': str})
    except Exception as e:
        print(f"❌ Error leyendo el archivo Excel: {e}")
        return

    print(f"📊 Total de contactos a procesar: {len(df_campistas)}")
    print("🛑 ¡NO TOQUES EL MOUSE! Tienes 5 segundos para cancelar (Ctrl+C)...")
    time.sleep(5)

    for index, row in df_campistas.iterrows():
        nombre = str(row['NOMBRE']).title()
        telefono = str(row['TELÉFONO']).replace(" ", "").strip()
        
        # Estandarización del prefijo internacional (Colombia)
        if not telefono.startswith('+57'): 
            telefono = '+57' + telefono

        print(f"\n[{index + 1}/{len(df_campistas)}] Generando mensaje para: {nombre}")

        # --- GENERACIÓN DE CONTENIDO CON IA ---
        try:
            time.sleep(2) # Pausa de seguridad para la API
            
            # Prompt de sistema estructurado
            prompt = (
                f"Actúa como un monitor de un bootcamp tecnológico. Escribe un saludo muy corto y amigable "
                f"para el campista {nombre}. Invítalo a asistir a la mentoría de nivelación (tras el receso de diciembre) "
                f"para empezar el 2026 con toda la energía. La cita es el 17 de enero de 2 pm a 3 pm. "
                f"Recuérdale que debe registrar su asistencia en el chat de la videollamada. "
                f"Incluye este enlace de la reunión: [ENLACE MEET]. "
                f"Usa emojis. Usa un solo asterisco de apertura y cierre para *resaltar* palabras clave."
            )

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt
            )
            texto_personalizado = response.text
            print("   -> IA generó el mensaje correctamente.")

        except Exception as e:
            print(f"   ⚠️ Falló la conexión con Gemini. Usando mensaje de respaldo. Error: {e}")
            texto_personalizado = f"Hola {nombre}, te invitamos a la mentoría de nivelación este 17 de enero de 2pm a 3pm. Link: [ENLACE MEET]. ¡No faltes!"

        # --- EJECUCIÓN ---
        enviar_recordatorio_whatsapp(telefono, texto_personalizado)
        barra_de_progreso(10, "Pausa de seguridad entre contactos...")

    print("\n🎉 Proceso de mensajería finalizado con éxito.")

if __name__ == "__main__":
    main()