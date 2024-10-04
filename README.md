# Proyecto de Generación de Presentaciones y Documentos

Este proyecto permite la creación automática de presentaciones en **PowerPoint** y documentos en **PDF** utilizando **Flask** y varias bibliotecas de Python.

## Requisitos previos

Antes de comenzar, asegúrate de tener instalado lo siguiente:

- **Python 3.x**: El proyecto requiere Python 3.
- **Virtualenv**: Para gestionar las dependencias de manera aislada.
- **LibreOffice**: Utilizamos LibreOffice para algunas operaciones de procesamiento de documentos. Puedes descargarlo desde el siguiente enlace:
  - [Descargar LibreOffice 24.2.6](https://es.libreoffice.org/descarga/libreoffice/)

## Instalación

Sigue los siguientes pasos para configurar y ejecutar el proyecto en tu entorno local:

1. **Clonar el repositorio**  
   Clona el repositorio en tu máquina local usando el siguiente comando:

   ```bash
   git clone <url_del_repositorio>
   cd <nombre_del_proyecto>

2. **Crear un entorno virtual y activalo**
- virtualenv -p python3 env_tesis
- env_tesis\Scripts\activate

3. **Instalar las dependencias**
Instala las dependencias necesarias ejecutando los siguientes comandos:

- pip install Flask
- pip install Flask-Limiter
- pip install openai
- pip install python-pptx
- pip install python-dotenv
- pip install python-docx

Verificar las dependencias instaladas

- pip list

**Notas adicionales**
LibreOffice: Algunas funcionalidades dependen de LibreOffice para la manipulación de documentos. Asegúrate de tenerlo instalado en tu sistema antes de ejecutar el proyecto. Puedes descargar la versión correcta desde el siguiente enlace:
LibreOffice 24.2.6 para Windows

**NOTA IMPORTANTE AGREGAR LAS SIGUIENTES CARPETAS EN LA RUTA RAIZ**

- Cache
- GeneratedPdf
- GeneratedPresentations
