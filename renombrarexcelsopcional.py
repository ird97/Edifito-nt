import os
import re

# Ruta principal donde se encuentran las carpetas por año con los archivos Excel. Cambiar por el que sea en tu caso
ruta_base = "/Users/ian/ojoalpiojoenelocho/ANEXOS/ANEXO 3.1 (excels)"

# Nombres de los meses, ordenados de enero a diciembre
meses = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

# Recorrer el rango de años que haya info disponible en tu plataforma, en mi caso los puse que recorriera cada año desde 2018 hasta 2026
for anio in range(2018, 2027):
    carpeta = os.path.join(ruta_base, str(anio))

    # Si la carpeta del año no existe, se omite
    if not os.path.isdir(carpeta):
        continue

    print(f"\nProcesando {anio}")

    # Listar todos los archivos .xlsx ignorando los temporales de oficina
    archivos = [
        archivo for archivo in os.listdir(carpeta)
        if archivo.lower().endswith(".xlsx") and not archivo.startswith("~$")
    ]

    for archivo in archivos:
        nombre = archivo.lower()
        mes_detectado = None

        # Buscar el nombre del mes escrito en el nombre del .xlsx
        for mes in meses:
            if mes in nombre:
                mes_detectado = mes
                break

        # Si no se pilló un nombre de mes, intentar buscar un número del 1 al 12
        if not mes_detectado:
            numeros = re.findall(r'(?:^|_|-)(\d{1,2})(?:_|-)', nombre)
            for numero in numeros:
                valor = int(numero)
                if 1 <= valor <= 12:
                    mes_detectado = meses[valor - 1]
                    break

        # Si no se pudo pillar el mes, avisa y se sigue con el siguiente .xlsx
        if not mes_detectado:
            print(f"No se detectó mes: {archivo}")
            continue

        # Generar el nuevo nombre base: "mesaño"
        nuevo_nombre_base = f"{mes_detectado}{anio}"
        nuevo_nombre = f"{nuevo_nombre_base}.xlsx"
        
        ruta_vieja = os.path.join(carpeta, archivo)
        ruta_nueva = os.path.join(carpeta, nuevo_nombre)

        # Si el nombre ya tiene el formato correcto, no se renombra
        if archivo == nuevo_nombre:
            print(f"Correcto: {archivo}")
            continue

        # Si el nombre del desrino ya existe, se agrega un sufijo numérico para evitar sobrescritura
        contador = 1
        while os.path.exists(ruta_nueva):
            if ruta_vieja.lower() == ruta_nueva.lower():
                break
            
            nuevo_nombre = f"{nuevo_nombre_base}_{contador}.xlsx"
            ruta_nueva = os.path.join(carpeta, nuevo_nombre)
            contador += 1

        # Renombrar el .xlsx
        try:
            os.rename(ruta_vieja, ruta_nueva)
            print(f"{archivo} -> {nuevo_nombre}")
        except Exception as e:
            print(f"Error con {archivo}: {e}")

print("\nProceso terminado")
