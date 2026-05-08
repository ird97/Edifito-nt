import os
import shutil
import openpyxl

# Años y meses que vamos a buscar en los archivos, en mi caso son estos años pero en el tuyo pueden ser distintos
AÑOS = ["2026", "2025", "2024", "2023", "2022", "2021", "2020", "2019", "2018"]
MESES = ["Diciembre", "Noviembre", "Octubre", "Septiembre", "Agosto", "Julio", "Junio", "Mayo", "Abril", "Marzo", "Febrero", "Enero"]

# Carpeta donde están los excels sueltos y desordenados
ORIGEN = os.path.join(os.getcwd(), "Gastos_Comunes_Edifito")
# Carpeta de destino donde todo quedará ordenadito
DESTINO = os.path.join(os.getcwd(), "ggcc_ordenaditos")

# Si no existe la carpeta origen, no hay nada que hacer
if not os.path.exists(ORIGEN):
    print(f"No encontre la carpeta: {ORIGEN}")
    exit()

# Creamos la carpeta destino y una subcarpeta para los archivos que no se puedan clasificar
os.makedirs(DESTINO, exist_ok=True)
carpeta_huerfanos = os.path.join(DESTINO, "Sin_Clasificar")
os.makedirs(carpeta_huerfanos, exist_ok=True)

# Juntamos solo los archivos .xlsx (nada de temporales ni otras cosas)
archivos = [f for f in os.listdir(ORIGEN) if f.endswith('.xlsx')]

if not archivos:
    print("No hay nada que mover.")
    exit()

print(f"Ordenando bip bip bip {len(archivos)} archivos...\n")

# Recorremos cada archivo para leer su contenido y adivinar mes y año
for archivo in archivos:
    ruta_origen = os.path.join(ORIGEN, archivo)
    mes_encontrado = None
    anio_encontrado = None
    
    try:
        # Abrimos el Excel en modo solo lectura y sin fórmulas, para ir más rápido
        wb = openpyxl.load_workbook(ruta_origen, data_only=True, read_only=True)
        hoja = wb.active
        
        # Recolectamos el texto de las primeras 30 filas en una sola cadena
        texto_interno = ""
        for fila in hoja.iter_rows(min_row=1, max_row=30, values_only=True):
            for celda in fila:
                if celda and isinstance(celda, str):
                    texto_interno += celda.lower() + " "
        wb.close()
        
        # Buscamos primero el año dentro del texto
        for anio in AÑOS:
            if anio in texto_interno:
                anio_encontrado = anio
                break
                
        # Después buscamos el mes, también dentro del texto
        for mes in MESES:
            if mes.lower() in texto_interno:
                mes_encontrado = mes
                break

        # Si encontramos ambos datos, movemos el archivo a su carpeta correspondiente
        if mes_encontrado and anio_encontrado:
            ruta_final = os.path.join(DESTINO, anio_encontrado, mes_encontrado)
            os.makedirs(ruta_final, exist_ok=True)
            
            nombre_nuevo = f"Gasto_Comun_{mes_encontrado}_{anio_encontrado}.xlsx"
            ruta_destino = os.path.join(ruta_final, nombre_nuevo)
            
            # Si ya existe un archivo con ese nombre, le ponemos un numerito para no sobrescribir
            cont = 1
            while os.path.exists(ruta_destino):
                nombre_nuevo = f"Gasto_Comun_{mes_encontrado}_{anio_encontrado}_Copia{cont}.xlsx"
                ruta_destino = os.path.join(ruta_final, nombre_nuevo)
                cont += 1
            
            shutil.move(ruta_origen, ruta_destino)
            print(f"Listo: {nombre_nuevo}")
        else:
            # Si falta el mes o el año, lo mandamos a la carpeta de huérfanos
            shutil.move(ruta_origen, os.path.join(carpeta_huerfanos, archivo))
            print(f"Sin fecha: {archivo}")
            
    except Exception:
        # Si algo sale mal al leer el archivo, también va a huérfanos
        shutil.move(ruta_origen, os.path.join(carpeta_huerfanos, archivo))
        print(f"Error con: {archivo}")

print("\nTerminado.")
