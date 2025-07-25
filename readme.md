# Extractor y Descargador de CURPs


**Autores:** Claude Sonnet y Gemini Assistant
**Editor:** Leonardo González Muñoz

Una aplicación de escritorio con interfaz gráfica (GUI) construida en Python y Tkinter que automatiza la extracción y descarga de Claves Únicas de Registro de Población (CURP) de México.

La herramienta permite a los usuarios extraer CURPs desde archivos de imagen (JPG, PNG) y PDF, ingresarlas manualmente, y descargar los documentos oficiales en formato PDF desde el portal del gobierno.



## Características Principales

-   **Extracción OCR:** Utiliza Tesseract OCR para extraer CURPs de imágenes y documentos PDF.
-   **Entrada Manual y Masiva:** Permite ingresar una o múltiples CURPs directamente en la interfaz.
-   **Validación de Formato:** Verifica en tiempo real que las CURPs ingresadas cumplan con el formato estándar de 18 caracteres.
-   **Descarga de PDFs Oficiales:** Automatiza la descarga de los documentos PDF oficiales de las CURPs a través de Selenium y ChromeDriver.
-   **Gestión de Resultados:** Muestra los resultados en una tabla interactiva donde se pueden seleccionar, copiar o eliminar CURPs.
-   **Exportación a CSV:** Permite exportar la lista de CURPs y su estatus a un archivo CSV compatible con Excel.
-   **Interfaz Intuitiva:** Interfaz gráfica fácil de usar construida con Tkinter, adaptable a diferentes tamaños de pantalla.

## Requisitos Previos

-   **Python 3.7+**
-   **Tesseract OCR:** El motor de reconocimiento óptico de caracteres.
-   **Google Chrome:** Necesario para la descarga automatizada de PDFs.


## Instalación automática para Windows
**Instalar Tesseract OCR:**
    -   **Windows:** Descarga e instala desde la wiki de Tesseract de UB-Mannheim. Asegúrate de que la ruta de instalación (`C:\Program Files\Tesseract-OCR\tesseract.exe`) coincida con la del script o añádela a la variable de entorno `PATH` del sistema.
    - Descarga y abre el archivo ExtractorCURP.exe


## Instalación desde cero

Sigue estos pasos para configurar el entorno de desarrollo.

1.  **Clonar el repositorio:**
    ```bash
    git clone https://github.com/tu-usuario/curp_downloader.git
    cd curp_downloader
    ```

2.  **Instalar Tesseract OCR:**
    -   **Windows:** Descarga e instala desde la wiki de Tesseract de UB-Mannheim. Asegúrate de que la ruta de instalación (`C:\Program Files\Tesseract-OCR\tesseract.exe`) coincida con la del script o añádela a la variable de entorno `PATH` del sistema.
    -   **macOS:** `brew install tesseract`
    -   **Linux (Debian/Ubuntu):** `sudo apt install tesseract-ocr`

3.  **Crear un entorno virtual (Recomendado):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # En Windows: venv\Scripts\activate
    ```

4.  **Instalar las dependencias de Python:**
    Crea un archivo `requirements.txt`  y luego instala las dependencias:
    ```bash
    pip install -r requirements.txt
    ```

## Uso

1.  Asegúrate de que todas las dependencias estén instaladas.
2.  Ejecuta el script principal desde tu terminal:
    ```bash
    python curp_extractor.py
    ```
3.  **Uso de la aplicación:**
    -   **Subir Archivos:** Haz clic en "Subir archivos" para seleccionar imágenes o PDFs. La aplicación buscará y extraerá las CURPs.
    -   **Entrada Manual:** Escribe una CURP en el campo de texto y presiona "Agregar CURP" o la tecla `Enter`.
    -   **Entrada Masiva:** Pega una lista de CURPs (una por línea) en el área de texto grande y haz clic en "Agregar todas las CURPs".
    -   **Descargar PDFs:** Selecciona las CURPs deseadas usando las casillas de verificación y haz clic en "Descargar PDFs seleccionados". Se te pedirá que elijas una carpeta de destino.
    -   **Exportar:** Usa los botones "Copiar CURPs" o "Descargar como CSV" para exportar los resultados.

## Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.
