# 🤖 Automatización de Asistencia: Bootcamps Talento Tech Valle

Este repositorio contiene un script ETL (Extract, Transform, Load) en Python diseñado para automatizar la gestión y el reporte de asistencia diaria de los campistas, eliminando horas de trabajo manual repetitivo.

## 🎯 El Problema
El control de asistencia requería un proceso manual, tedioso y propenso a errores humanos que consistía en:
* Ingresar diariamente al campus virtual y navegar por múltiples menús para descargar el reporte en Excel.
* Filtrar manualmente hojas de cálculo con datos de aproximadamente 30 campistas por cohorte.
* Mapear estados numéricos del sistema a formatos visuales de validación.
* Copiar y pegar celda por celda la información depurada en un Google Sheet compartido para el análisis de la coordinación.

## 🚀 La Solución
Se desarrolló un bot que automatiza el flujo completo de extracción, limpieza y carga de datos. Al ejecutar el script en consola, el usuario cuenta con un menú interactivo que permite actualizar el histórico completo, un rango de fechas, o únicamente la sesión del día. 

**Flujo automatizado:**
1. **Extracción (Web Scraping):** El script utiliza Selenium para abrir un navegador en segundo plano, autenticarse de forma segura en la plataforma, navegar por las rutas del campus y descargar el archivo de Excel crudo.
2. **Transformación:** Mediante Pandas y Openpyxl, el bot lee el archivo, limpia los datos y transforma los valores numéricos del sistema a los formatos requeridos.
3. **Carga (API Sincronización):** Se integra con Google Cloud (mediante una Service Account) y la librería Gspread para conectarse al Google Sheet de la coordinación e inyectar los datos limpios en tiempo real, cerrando la sesión y eliminando los archivos temporales al finalizar.

## 🛠️ Tecnologías Utilizadas
* **Python 3.x**
* **Selenium:** Para la automatización del navegador web y descarga de reportes.
* **Pandas & Openpyxl:** Para la manipulación, lectura y limpieza de datos en Excel.
* **Gspread (Google Sheets API):** Para la manipulación e inyección de datos en la nube.

> **Nota de desarrollo:** La lógica base de este script fue estructurada y acelerada utilizando herramientas de Inteligencia Artificial (Prompt Engineering), adaptando el código final a los requerimientos específicos de la lógica de negocio.

## ⚙️ Cómo ejecutarlo localmente

> ⚠️ **Seguridad:** El archivo de credenciales JSON de la cuenta de servicio de Google Cloud no se incluye en este repositorio. Debes generar el tuyo en GCP.

1. Clona este repositorio:
   ```bash
   git clone [https://github.com/Jhonarvi/Automatizacion-Procesos-Python.git](https://github.com/Jhonarvi/Automatizacion-Procesos-Python.git)

### 2. Bot de Mensajería Masiva con IA (WhatsApp)
* **El Problema:** El envío manual de recordatorios diarios a más de 30 campistas consumía mucho tiempo y el uso de plantillas genéricas reducía la interacción de los estudiantes.
* **La Solución:** Script en Python que lee una base de datos telefónica, utiliza la API de **Google Gemini 2.5 Flash** para redactar mensajes amigables y 100% personalizados por estudiante, y ejecuta el envío automático a través de WhatsApp Web (RPA básico).
* **Tecnologías:** Python, Google GenAI API (Prompt Engineering), Pandas, PyWhatKit, PyAutoGUI.
