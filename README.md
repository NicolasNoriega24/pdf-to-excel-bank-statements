PDF to Excel - Bank Statements
Descripción
Script en Python que automatiza la extracción de datos desde extractos bancarios en formato PDF y los convierte en tablas estructuradas en Excel listas para análisis.
Luego fue pensado para ser trabajado por el equipo de trabajo para que se coloquen los PDFs en las carpetas indicadas para el que el script ejecute y devuelva los excel correspondientes.

Problema
Los extractos bancarios en PDF no permiten un análisis eficiente ni automatizado.
Esto obliga a realizar cargas manuales, aumentando el tiempo de trabajo y la probabilidad de errores.

Solución
Se desarrolló un proceso automatizado que:

Lee archivos PDF
Identifica la estructura de los datos mediante coordenadas
Reconstruye la información en formato tabular
Genera un archivo Excel listo para análisis
Lee la carpeta de recepción de los PDF
Devuelve los Excel en las carpetas que sufrieron modificaciones
Disminuye el tiempo de trabajo operativo en un 80%

Tecnologías utilizadas
Python
pandas
pdfplumber
openpyxl
watchdog

Cómo usar el proyecto
Clonar el repositorio:
git clone https://github.com/TU-USUARIO/pdf-to-excel-bank-statements.gi

