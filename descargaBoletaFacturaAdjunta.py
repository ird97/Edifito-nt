import os
import re
import urllib.parse
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Períodos históricos: Edifito tiene rendiciones desde marzo 2018 hasta febrero 2026
ANIOS = ["2026", "2025", "2024", "2023", "2022", "2021", "2020", "2019", "2018"]
MESES = [
    "Diciembre", "Noviembre", "Octubre", "Septiembre",
    "Agosto", "Julio", "Junio", "Mayo",
    "Abril", "Marzo", "Febrero", "Enero"
]

# Carpeta temporal donde se guardan los documentos descargados
BASE_DIR = os.path.join(os.getcwd(), "boletasgarcas_excel")
os.makedirs(BASE_DIR, exist_ok=True)


def iniciar_driver():
    # Configura Chrome para que descargue archivos sin abrirlos en el navegador
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": BASE_DIR,
        "plugins.always_open_pdf_externally": True
    })

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver


def sincronizar_sesion(driver):
    # Copia las cookies del navegador a una sesión de requests para descargar archivos
    session = requests.Session()

    for cookie in driver.get_cookies():
        session.cookies.set(cookie["name"], cookie["value"])

    # Header básico para no ir con las manos vacías
    session.headers.update({
        "User-Agent": "Mozilla/5.0"
    })

    return session


def descargar_archivo(session, url, destino):
    # Descarga un archivo verificando que sea válido y no una página de error
    try:
        r = session.get(url, stream=True, timeout=20)

        if r.status_code != 200:
            return False

        # Si la respuesta es HTML, probablemente sea un error o redirección
        if "text/html" in r.headers.get("Content-Type", ""):
            return False

        with open(destino, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)

        return True

    except Exception:
        return False


def procesar_periodo(driver, wait, session, anio, mes):
    # Navega a la página de consulta de documentos para un mes y año específicos
    driver.get("https://clientes.edifito.com/main.php?prm0=0x2200")

    # Espera a que los selectores estén disponibles
    elemento_anio = wait.until(EC.presence_of_element_located((By.ID, "anio")))
    elemento_mes = wait.until(EC.presence_of_element_located((By.ID, "mes")))

    # Guardamos el HTML actual para detectar cuándo se recarga la página
    html_antiguo = driver.find_element(By.TAG_NAME, "html")

    # Selecciona el período
    Select(elemento_anio).select_by_visible_text(anio)
    Select(elemento_mes).select_by_visible_text(mes)

    # Envía el formulario usando JavaScript
    driver.execute_script("document.month.submit();")

    # Espera a que la página se recargue (el HTML antiguo ya no esté presente)
    wait.until(EC.staleness_of(html_antiguo))

    # Verifica que el servidor realmente cargó el período solicitado
    nuevo_anio = wait.until(EC.presence_of_element_located((By.ID, "anio")))
    nuevo_mes = wait.until(EC.presence_of_element_located((By.ID, "mes")))

    if (Select(nuevo_anio).first_selected_option.text != anio or
            Select(nuevo_mes).first_selected_option.text != mes):
        # Si no coincide, es porque no hay datos para ese período
        return

    # Crea la carpeta para guardar los archivos de este período
    carpeta = os.path.join(BASE_DIR, anio, mes)
    os.makedirs(carpeta, exist_ok=True)

    # Espera a que aparezca la tabla con los documentos
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "table-hover")))
    elementos = driver.find_elements(By.CLASS_NAME, "lnr-magnifier")

    for idx, el in enumerate(elementos):
        accion = el.get_attribute("onclick")

        # Solo interesan los elementos que llaman a la función verImg
        if not accion or "verImg" not in accion:
            continue

        # Extrae la URL del archivo y el tipo desde el código JavaScript
        url_match = re.search(r"verImg\('(.*?)'", accion)
        tipo_match = re.search(r"verImg\('.*?','(.*?)'\)", accion)

        if not url_match:
            continue

        url_aws = url_match.group(1)
        tipo = tipo_match.group(1) if tipo_match else "pdf"

        # Codifica la URL de S3 para pasarla por el endpoint de Edifito
        url_codificada = urllib.parse.quote(url_aws, safe="")
        url_final = f"https://clientes.edifito.com/verFactura.php?img={url_codificada}&t={tipo}"

        # Nombre del archivo a partir del último segmento de la URL
        nombre = url_aws.split("/")[-1].split("?")[0]
        extension = tipo if tipo in ["pdf", "zip"] else "jpg"

        ruta = os.path.join(carpeta, f"{nombre}.{extension}")

        # Descarga el archivo usando requests
        descargar_archivo(session, url_final, ruta)


def main():
    driver = iniciar_driver()
    wait = WebDriverWait(driver, 15)

    try:
        # Página de acceso para que el usuario inicie sesión manualmente, si falla, simplemente busca donde iniciar sesión
        driver.get("https://clientes.edifito.com/")
        input("Inicie sesión y presione ENTER para continuar...")

        # Sincroniza la sesión autenticada con requests
        session = sincronizar_sesion(driver)

        # Recorre todos los períodos y descarga los documentos encontrados
        for anio in ANIOS:
            for mes in MESES:
                try:
                    procesar_periodo(driver, wait, session, anio, mes)
                except Exception:
                    # Si falla un período, se sigue con el siguiente sin detener todo
                    continue

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
