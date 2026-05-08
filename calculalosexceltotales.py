import pandas as pd
from pathlib import Path

# Carpeta donde se encuentran los archivos Excel organizados
ruta_base = Path("/Users/ian/Auditoria_Excel_Organizado/excel/2026")

# Buscar todos los archivos .xlsx dentro de la estructura
archivos = list(ruta_base.rglob("*.xlsx"))

print(f"Archivos encontrados: {len(archivos)}")

registros = []

# Procesar cada archivo Excel
for archivo in archivos:
    try:
        # Leer sin encabezado porque los formatos pueden variar
        df = pd.read_excel(archivo, header=None)

        for _, row in df.iterrows():
            fila = row.tolist()

            fecha = fila[3]
            monto = fila[5]

            # Incluir solo filas que contengan una fecha con guiones y un monto válido
            if isinstance(fecha, str) and "-" in fecha and pd.notna(monto):

                registros.append({
                    "archivo": archivo.name,
                    "proveedor": str(fila[1]).strip(),
                    "fecha": fecha,
                    "descripcion": str(fila[4]).strip(),
                    "monto": float(monto)
                })

    except Exception as e:
        print(f"Error en {archivo}: {e}")

# Construir el DataFrame con todos los registros extraídos
df = pd.DataFrame(registros)

if df.empty:
    print("No se extrajo información")
    exit()

# Normalizar las fechas a tipo datetime (asumiendo día primero)
df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True)

# Agregar columnas de año y mes para los análisis
df["año"] = df["fecha"].dt.year
df["mes"] = df["fecha"].dt.month

# Calcular totales por año y por mes
total_anual = df.groupby("año")["monto"].sum()
total_mensual = df.groupby(["año", "mes"])["monto"].sum()

# Ranking de proveedores según monto total
top_proveedores = (
    df.groupby("proveedor")["monto"]
    .sum()
    .sort_values(ascending=False)
)

# Detectar meses faltantes para cada año
faltantes_por_año = {}

for año in sorted(df["año"].dropna().unique()):
    año = int(año)

    meses_presentes = set(
        df[df["año"] == año]["mes"].astype(int).unique()
    )

    meses_esperados = set(range(1, 13))

    faltantes = meses_esperados - meses_presentes

    faltantes_por_año[año] = sorted(faltantes)

# Exportar todos los resultados a un archivo Excel
with pd.ExcelWriter("revisionywe_historica.xlsx") as writer:
    df.to_excel(writer, sheet_name="base", index=False)
    total_anual.to_excel(writer, sheet_name="total_anual")
    total_mensual.to_excel(writer, sheet_name="total_mensual")
    top_proveedores.to_excel(writer, sheet_name="top_proveedores")

    pd.DataFrame([
        {"año": k, "meses_faltantes": v}
        for k, v in faltantes_por_año.items()
    ]).to_excel(writer, sheet_name="faltantes", index=False)

print("Auditoría generada: auditoria_final_historica.xlsx")

# Mostrar en consola los meses faltantes por año
print("\nMeses faltantes por año:")

for año, faltantes in faltantes_por_año.items():
    print(año, "->", faltantes)
