import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Años y meses que se van a recorrer para descargar los Excels, adapta los años a tu caso
ANOS = ["2026", "2025", "2024", "2023", "2022", "2021", "2020", "2019", "2018", "2017"]
MESES = ["Diciembre", "Noviembre", "Octubre", "Septiembre", "Agosto", "Julio", "Junio", "Mayo", "Abril", "Marzo", "Febrero", "Enero"]

# Carpeta donde se guardarán los archivos Excel descargados (misma carpeta que los PDFs, puedes cambiarla si quieres separar)
CARPETA_DESCARGAS = os.path.join(os.getcwd(), "Gastos_Comunes_Edifito")
if not os.path.exists(CARPETA_DESCARGAS):
    os.makedirs(CARPETA_DESCARGAS)

# Opciones de Chrome para que las descargas se hagan automáticamente en la carpeta indicada
chrome_options = Options()
prefs = {
    "download.default_directory": CARPETA_DESCARGAS,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "plugins.always_open_pdf_externally": True  # para PDFs no afecta, pero lo dejamos por si acaso
}
chrome_options.add_experimental_option("prefs", prefs)

# Iniciar el navegador con el driver gestionado automáticamente
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.maximize_window()
wait = WebDriverWait(driver, 15)

try:
    # Ir a la página principal de clientes Edifito (si falla, verifica la URL de inicio de sesión correcta)
    driver.get("https://clientes.edifito.com")
    
    print("Esperando inicio de sesión manual...")
    # El usuario debe iniciar sesión manualmente y luego presionar Enter para continuar
    input("Presione ENTER después de iniciar sesión para continuar...")

    # Recorrer cada año y cada mes
    for anio in ANOS:
        for mes in MESES:
            try:
                # Navegar a la sección de consulta de documentos
                driver.get("https://clientes.edifito.com/main.php?prm0=0x2200")
                
                # Esperar a que los selectores de año y mes estén visibles
                select_anio_el = wait.until(EC.visibility_of_element_located((By.ID, "anio")))
                select_mes_el = wait.until(EC.visibility_of_element_located((By.ID, "mes")))
                
                # Guardamos una referencia al body para detectar si la página cambia
                cuerpo_antiguo = driver.find_element(By.TAG_NAME, 'body')

                # Seleccionar el año y el mes en los desplegables
                Select(select_anio_el).select_by_visible_text(anio)
                Select(select_mes_el).select_by_visible_text(mes)

                # Enviar el formulario
                driver.execute_script("document.month.submit();")
                
                # Esperar a que la página se actualice (el body antiguo desaparezca)
                wait.until(EC.staleness_of(cuerpo_antiguo))
                
                # Verificar que la selección de año y mes realmente se haya aplicado
                nuevo_anio = wait.until(EC.presence_of_element_located((By.ID, "anio")))
                nuevo_mes = driver.find_element(By.ID, "mes")
                
                anio_actual = Select(nuevo_anio).first_selected_option.text
                mes_actual = Select(nuevo_mes).first_selected_option.text

                # Si coincide, hay datos y podemos descargar el Excel
                if anio_actual == anio and mes_actual == mes:
                    print(f"Descargando Excel: {mes} {anio}")
                    driver.execute_script("excelLey.submit();")   # <-- Aquí se pide el Excel en vez del PDF
                    time.sleep(5)  # Esperar a que se complete la descarga
                else:
                    print(f"No hay registros para: {mes} {anio}")

            except Exception:
                # Si algo falla (mes sin datos, error de carga, etc.), se pasa al siguiente
                continue

    print("Ready mi rey.")

except Exception as e:
    print(f"Error en la ejecución: {e}")

finally:
    # Cerrar el navegador pase lo que pase
    driver.quit()
