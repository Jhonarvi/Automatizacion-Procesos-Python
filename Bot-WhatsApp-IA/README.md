# 🤖 Bot de Mensajería Masiva con Inteligencia Artificial

Este script en Python automatiza el envío de recordatorios y notificaciones masivas por WhatsApp, utilizando la API de Google Gemini para redactar mensajes dinámicos y personalizados para cada usuario, evitando los bloqueos por *spam* que generan las plantillas estáticas.

## 🎯 El Problema
La gestión de comunicaciones en bootcamps educativos requería enviar recordatorios diarios (clases, nivelaciones, enlaces) a más de 30 campistas. Hacerlo de forma manual consumía tiempo excesivo, y usar mensajes genéricos de "copiar y pegar" disminuía la cercanía y el interés de los estudiantes por participar.

## 🚀 La Solución
Se desarrolló un bot híbrido que combina la manipulación de datos, la Inteligencia Artificial y la automatización de la interfaz de usuario (RPA básico). 

**Flujo del Bot:**
1. Lee una base de datos local (`.xlsx`) extrayendo y limpiando los números telefónicos.
2. Se conecta a la API de **Gemini 2.5 Flash** enviando un prompt dinámico con el nombre de cada campista para generar un mensaje amigable, único y contextualizado.
3. Utiliza automatización de navegador y teclado para abrir WhatsApp Web chat por chat, pegando y enviando el mensaje generado sin intervención humana.

## 🛠️ Tecnologías y Librerías Utilizadas
* **Python 3.x**
* **Google GenAI API:** Para la generación de texto dinámico y natural (Prompt Engineering).
* **Pandas:** Para la extracción y estandarización de datos desde Excel.
* **PyWhatKit & PyAutoGUI:** Para la orquestación del navegador, control de tiempos de carga y simulación de eventos físicos del teclado.

## ⚙️ Configuración y Uso Local

> ⚠️ **Importante:** La API Key de Google no se incluye por motivos de seguridad. Debes configurarla como variable de entorno.

1. Clona el repositorio e instala las dependencias:
   ```bash
   pip install pywhatkit pandas pyautogui google-genai