import os
import pandas as pd
import re

# Ruta donde están las carpetas con los archivos Excel por año
ruta_base = "/Users/ian/ojoalpiojoenelocho/ANEXOS/ANEXO 3.1 (excels)"

# Años que vamos a analizar
anios = [str(a) for a in range(2018, 2027)]

# Nombres de meses en minúscula (tal como aparecen en los nombres de archivo)
meses_nombres = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Aquí iremos acumulando los registros limpios de todos los archivos
datos_consolidados = []

def limpiar_monto(valor):
    """Convierte valores como '$ 1.234.567' o '1,234.567' a float"""
    if pd.isna(valor): return 0
    # Quitamos símbolos de moneda, puntos de miles y cambiamos coma decimal por punto
    s = str(valor).replace('$', '').replace('.', '').replace(' ', '').replace(',', '.')
    try:
        return float(re.sub(r'[^-0-9.]', '', s))
    except:
        return 0

print("SUPERCOMPUTADORA INICIANDO ANAL PROFUNDO")

# Recorremos cada año y cada mes buscando el archivo correspondiente
for anio in anios:
    ruta_anio = os.path.join(ruta_base, anio)
    if not os.path.isdir(ruta_anio): continue
    
    for mes in meses_nombres:
        nombre_archivo = f"{mes}{anio}.xlsx"
        ruta_archivo = os.path.join(ruta_anio, nombre_archivo)
        
        if os.path.exists(ruta_archivo):
            try:
                # Leemos el archivo sin conocer aún dónde empieza la tabla real
                df = pd.read_excel(ruta_archivo)
                col_monto = None
                
                # Buscamos en las primeras 20 filas la columna que diga "VALOR"
                for i in range(20):
                    temp_df = pd.read_excel(ruta_archivo, skiprows=i)
                    posibles = [c for c in temp_df.columns if "VALOR" in str(c).upper()]
                    if posibles:
                        df = temp_df
                        col_monto = posibles[0]
                        break
                
                if col_monto:
                    # Convertimos los montos a números limpios
                    df[col_monto] = df[col_monto].apply(limpiar_monto)
                    
                    # Eliminamos filas sin nombre de proveedor o con monto cero
                    df = df.dropna(subset=['Nombre'])
                    df = df[df[col_monto] > 0]
                    
                    # Agregamos columnas de año, mes textual y periodo en formato YYYY-MM
                    df['Año'] = int(anio)
                    df['Mes'] = mes
                    df['Periodo'] = f"{anio}-{str(meses_nombres.index(mes)+1).zfill(2)}"
                    
                    # Unificamos el nombre de la columna de monto
                    df = df.rename(columns={col_monto: 'Monto'})
                    
                    datos_consolidados.append(df[['Año', 'Mes', 'Periodo', 'Nombre', 'Monto']])
                    print(f"{mes} {anio} procesado.")
            except Exception as e:
                print(f"Error en {nombre_archivo}: {e}")
        else:
            print(f"Faltante: {mes} {anio}")

# Si logramos extraer información, armamos el reporte
if datos_consolidados:
    df_final = pd.concat(datos_consolidados, ignore_index=True)
    
    # 1. Resumen mensual con evolución gasto a gasto
    resumen_mensual = df_final.groupby(['Año', 'Periodo', 'Mes'])['Monto'].sum().reset_index()
    resumen_mensual = resumen_mensual.sort_values('Periodo')
    resumen_mensual['Variación $'] = resumen_mensual['Monto'].diff()
    resumen_mensual['% Incremento'] = resumen_mensual['Monto'].pct_change() * 100

    # 2. Ranking de proveedores (quién recibe más)
    resumen_proveedor = df_final.groupby('Nombre')['Monto'].agg(['sum', 'count', 'mean']).sort_values('sum', ascending=False)
    resumen_proveedor.columns = ['Total Pagado', 'Cant. Meses', 'Promedio Mensual']

    # Guardamos todo en un Excel con varias hojas
    ruta_salida = os.path.join(ruta_base, "REPORTE_AUDITORIA_FINAL.xlsx")
    with pd.ExcelWriter(ruta_salida) as writer:
        df_final.to_excel(writer, sheet_name='Data_Cruda', index=False)
        resumen_mensual.to_excel(writer, sheet_name='Evolucion_Mensual', index=False)
        resumen_proveedor.to_excel(writer, sheet_name='Analisis_Proveedores')
    
    print(f"\nProcesamiento completado. Archivo generado en: {ruta_salida}")
else:
    print("No se pudo extraer ninguna información.")
