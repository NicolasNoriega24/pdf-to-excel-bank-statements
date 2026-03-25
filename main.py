"""
Proyecto: PDF to Excel - Bank Statements
Autor: Nicolás Noriega

Descripción:
Script que extrae datos desde PDFs de extractos bancarios
y los convierte en tablas estructuradas en Excel. Pensado para trabajos en servidor
donde se alojan los PDFs en las carpetas crrespondientes y se devuelve los excel listos
para ser trabajados.
"""
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
import subprocess
import os

# =========================
# CONFIGURACIÓN DE BANCOS
# =========================
BANCOS = {
    "MercadoPago": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Mercado Pago"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Mercado Pago"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosMERCADOPAGO.py")
    },
    "Comafi": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Comafi"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Comafi"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosCOMAFI.py")
    },
    "Bancor": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Bancor"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Bancor"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosBANCOR.py")
    },
    "Macro": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Macro"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Macro"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosMACRO.py")
    },
    "BBVA": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\BBVA"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\BBVA"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosBBVA.py")
    },
    "Galicia": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Galicia"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Galicia"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosGALICIA.py")
    },
    "Santander": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Santander"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Santander"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosSANTANDER.py")
    },
    "Credicoop": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Credicoop"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Credicoop"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosCREDICOOP.py")
    },
    "Nacion": {
        "entrada": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Nacion"),
        "salida": Path(r"C:\Users\Lenovo123\Desktop\Proyecto Automatizacion\Extractos Bancarios\Nacion"),
        "script": Path(r"C:\Users\Lenovo123\Desktop\Sabemosdecampo\BotExtactosBancarios\PDFExtractosNACION.py")
    }
}

# =========================
# CONFIGURACIÓN DE LIMPIEZA
# =========================
LIMPIEZA_INTERVALO = 60 * 60   # segundos, ej: 1 hora
EXTENSIONES_LIMPIEZA = [".pdf", ".xlsx"]  # extensiones a borrar

# =========================
# HANDLER DE EVENTOS
# =========================
class MiHandler(FileSystemEventHandler):
    def __init__(self, bancos_config):
        super().__init__()
        self.bancos_config = bancos_config

    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return

        pdf_path = Path(event.src_path)

        for banco, cfg in self.bancos_config.items():
            if pdf_path.parent == cfg["entrada"]:
                print(f"Nuevo archivo detectado en {banco}: {pdf_path.name}")
                salida_excel = cfg["salida"] / (pdf_path.stem + ".xlsx")
                try:
                    subprocess.run(
                        ["python", str(cfg["script"]), str(pdf_path), str(salida_excel)],
                        check=True
                    )
                    print(f"Procesado y guardado en: {salida_excel}")
                except Exception as e:
                    print(f"Error procesando {pdf_path.name}: {e}")
                break
        else:
            print(f"Archivo {pdf_path.name} no pertenece a ninguna carpeta configurada.")

# =========================
# FUNCIONES DE LIMPIEZA
# =========================
def limpiar_carpetas():
    for banco, cfg in BANCOS.items():
        for carpeta in [cfg["entrada"], cfg["salida"]]:
            for ext in EXTENSIONES_LIMPIEZA:
                for f in carpeta.glob(f"*{ext}"):
                    try:
                        f.unlink()
                        print(f"Borrado archivo: {f}")
                    except Exception as e:
                        print(f"No se pudo borrar {f}: {e}")

# =========================
# MAIN LOOP
# =========================
if __name__ == "__main__":
    # Crear carpetas si no existen
    for cfg in BANCOS.values():
        os.makedirs(cfg["entrada"], exist_ok=True)
        os.makedirs(cfg["salida"], exist_ok=True)

    observer = Observer()
    event_handler = MiHandler(BANCOS)

    for cfg in BANCOS.values():
        observer.schedule(event_handler, str(cfg["entrada"]), recursive=False)

    observer.start()
    print("Monitorizando carpetas de todos los bancos...")

    ultima_limpieza = time.time()

    try:
        while True:
            time.sleep(1)
            # Ejecutar limpieza si pasó el intervalo
            if time.time() - ultima_limpieza >= LIMPIEZA_INTERVALO:
                print("Iniciando limpieza de carpetas...")
                limpiar_carpetas()
                ultima_limpieza = time.time()

    except KeyboardInterrupt:
        observer.stop()
    observer.join()
