# Edifito'nt

Extracción, "normalización" y organización de los antecedentes financieros que la plataforma Edifito.com dificulta acceder 

La arquitectura del sitio está pensada para que cualquier extracción masiva sea difícil y poco eficiente. No existen endpoints RESTful ni URLs persistentes para los documentos; en su lugar el sistema se apoya en formularios con campos ocultos, scripts que ejecutan métodos POST y parámetros de consulta ofuscados en hexadecimal como 0x2200. Las validaciones de referrer (strict-origin-when-cross-origin) y los timers de sesión del lado del cliente impiden cualquier parsing estático. Con Selenium, se mantuvo la sesión.

El inicio de sesión puede automatizarse con la cookie de sesión pero para mi fue más rápido iniciar sesión en cada automatización con Selenium para asegurarme de que tuviera el camino correcto, así que en todos los scripts la autenticación es manual. Una vez logueado, arranca la parte de seleccionar año y mes, enviar el formulario con JavaScript y esperar a que la página recargue por completo. Edifito omite en que un mes no tiene información, nos muestra un aviso o una pantalla vacía, recarga silenciosamente el último mes válido que el servidor tenga en caché. Si no validas que el mes devuelto coincide con el que pediste, puedes archivar datos que no corresponden. Los scripts chequean justo eso comparando los selectores después de cada recarga y solo siguen adelante si el período es correcto.

Los archivos que se bajan también vienen con lo suyo. Si elijes descargar las cartolas en PDF, todas se llaman gasto_comun.pdf y el sistema operativo les agrega el (1), (2) como gasto_comun(1).pdf, gasto_comun(2).pdf y así. Si seleccionas el "Excel" (xlsx), el nombre es Gasto_Comun.xlsx sin ninguna referencia al mes ni al año, y puedes descargar el mismo período infinitas veces sin que la plataforma te advierta. Al repetir una descarga, el navegador empieza a renombrar automáticamente con Copia1, Copia2, etc. Un desorden en nombres. 

Las planillas Excel que entrega Edifito son un ejemplo de cómo no:

-encabezados que cambian de un archivo a otro

-subtotales incrustados en las mismas columnas del detalle

-jerarquías tipo 2.1, 2.1.1 donde hay filas en un mismo bloque

-celdas fusionadas a mansalva

-montos que aparecen en columnas distintas (SUBTOTAL, TOTAL, VALOR)


# calculalosexceltotales.py

Procesa todos los archivos "Excel" (.xlsx) ubicados dentro de una carpeta base (estructura anidada) sin asumir un formato fijo de columnas. Lee cada hoja sin encabezado, identifica filas donde la cuarta columna contenga una fecha con guiones y la sexta columna un monto numérico, y las consolida en un único DataFrame. Luego normaliza las fechas, extrae año y mes, y calcula totales anuales y mensuales, un ranking de proveedores por monto acumulado, y detecta los meses que faltan en cada año presente. Finalmente exporta todos los resultados a un archivo Excel con hojas separadas ("base", "total_anual", "total_mensual", "top_proveedores" y "faltantes") y muestra en consola el resumen de meses ausentes.

# calculoexcelopcional.py 

Analiza archivos contables mensuales organizados en carpetas por año, donde cada archivo se nombra como "mesaño.xlsx" (ej. "enero2024.xlsx"). Para cada archivo, busca dinámicamente la posición de la tabla real dentro de las primeras 20 filas para una columna que contenga la palabra "VALOR". Convierte los montos desde formatos locales (puntos de miles, comas decimales) a flotantes, filtra filas con proveedor y monto positivo, y agrega columnas de período. Junta todos los registros y genera una evolución mensual con diferencias y porcentajes de incremento intermensual, y un ranking de proveedores con total pagado, frecuencia y promedio. El resultado se guarda en un Excel con tres hojas (datos crudos, evolución y análisis de proveedores).

# descargaBoletaFacturaAdjunta.py

Automatiza la descarga de documentos adjuntos (boletas, facturas, imágenes) desde el portal de Edifito. Inicia una sesión manual del usuario mediante Selenium, luego replica las cookies de autenticación en una sesión de requests. Itera por años y meses, selecciona cada período en un formulario web, y, tras la recarga, extrae de la tabla resultante todos los enlaces a archivos mediante expresiones regulares que capturan URLs de AWS S3 y su tipo (pdf/jpg/zip). Codifica la URL y la pasa a través del endpoint verFactura.php para descargar cada archivo binario de vuelta con requests, guardándolo en una estructura de carpetas año/mes/. Si un período no tiene datos o la página no carga, continúa sin detenerse.

# descargaExcels.py 

Descarga los archivos Excel de gastos comunes desde la plataforma Edifito. Requiere que el usuario inicie sesión manualmente y luego recorre todas las combinaciones de año y mes. Para cada período, navega a la página de consulta, selecciona año y mes mediante los desplegables correspondientes, envía el formulario, espera la recarga completa y verifica que la selección se haya aplicado. Si el servidor devuelve datos, ejecuta excelLey.submit() para forzar la descarga del archivo Excel (en lugar del PDF). Las descargas se guardan en una carpeta predefinida, con una pausa de 5 segundos entre cada una para evitar sobrecargas.

# descargaPDFs.py 

Similar al script anterior, pero para la descarga de los archivos PDF. Con la misma lógica de sesión manual y recorrido de años y meses, utiliza pdfLey.submit() en lugar del llamado al Excel. Está configurado para descargar automáticamente en la carpeta especificada sin intervención del usuario, y omite los períodos en los que no hay información disponible.

# ordenaexcel.py

Organiza un conjunto de archivos Excel de gastos comunes que originalmente están en una carpeta plana. Lee cada archivo con openpyxl en modo sólo lectura y extrae todo el texto de las primeras 30 filas. Sobre ese texto busca la presencia de un año (de una lista predefinida) y un mes (en castellano). Si encuentra ambos, mueve el archivo a una estructura de directorios Año/Mes bajo una carpeta destino, renombrándolo a un formato estándar Gasto_Comun_Mes_Año.xlsx. Si el nombre destino ya existe, añade un sufijo incremental para evitar sobrescritura. Los archivos donde no se puede determinar el año o el mes, o que producen un error de lectura, se trasladan a una subcarpeta Sin_Clasificar.

# renombrarexcelsopcional.py

Recorre una jerarquía de carpetas por año (en mi caso es de 2018 a 2026) y renombra archivos .xlsx (Excels) que no sigan la nomenclatura esperada mesaño.xlsx. Para cada archivo, intenta detectar el mes primero buscando el nombre del mes en español dentro del nombre del archivo; si no lo encuentra, extrae números de 1 al 12 (obviamente 1 siendo enero y 12 diciembre) mediante una expresión regular y lo asigna. Una vez identificado el mes y conocido el año por la carpeta contenedora, crea el nuevo nombre y renombra para evitar sobrescribir agregando un sufijo numérico si el nuevo nombre ya existe. Omite los archivos que ya tienen el formato deseado e informa aquellos cuyo mes no pudo ser determinado.



