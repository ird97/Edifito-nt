# Edifito'nt

Este repositorio documenta el proceso de extracción, normalización y organización de los antecedentes financieros que la plataforma Edifito (Mi Conserje SpA) pone al alcance de los que pagamos GG.CC. 

Los gastos comunes en Chile son desproporcionadamente altos y eso tiene una explicación bastante lamentable, muchas administraciones y comités manejan los fondos con mala fe, y Edifito parece ponerle un peaje a esa práctica entregando una plataforma por 100mil pesos al mes que vuelve cualquier auditoría un cacho. La información está fragmentada, la navegación es rígida y acceder a registros viejos requiere decenas de interacciones manuales. No existe un botón para bajarse todo de una sola vez sino que te obligan a pedir mes por mes, año por año, con clics que recargan la página entera cada vez. Obviamente lo primero que hice fue solicitarla al administrador, este no quiso por "razones misteriosas", el comité ídem, y la empresa responsable de la plataforma "Mi Conserje SpA" se nego después de ene insistencia. 

La arquitectura del sitio está pensada para que cualquier extracción masiva sea irritante, gastadora de tiempo e igual que caminar sobre huevos, un pequeño error u olvido y te perdiste. No existen endpoints RESTful ni URLs persistentes para los documentos; en su lugar el sistema se apoya en formularios con campos ocultos, scripts que ejecutan métodos POST y parámetros de consulta ofuscados en hexadecimal como 0x2200. Las validaciones de referrer (strict-origin-when-cross-origin) y los timers de sesión del lado del cliente impiden cualquier parsing estático, por lo que tuve que orquestar la máquina de estados con un navegador real mediante Selenium. Me puse en la situación de que había que mantener la sesión viva pero enhorabuena no se me corto, lidiar con modales que bloquean la interfaz e interceptar flujos de descarga que solo se disparan después de invocar funciones como javascript:pdf.submit() o validar().

<img width="1699" height="918" alt="Captura de pantalla 2026-05-07 a la(s) 21 35 39" src="https://github.com/user-attachments/assets/e4cabb61-5355-4913-ac34-67eafbedcaf7" />


<img width="1695" height="899" alt="Captura de pantalla 2026-05-07 a la(s) 21 37 37" src="https://github.com/user-attachments/assets/6aa2b39c-e183-48fa-bc9c-1f1db194649b" />

El inicio de sesión puede automatizarse con la cookie de sesión pero para mi fue más rápido iniciar sesión en cada automatización con selenium para asegurarme de que tuviera el camino correcto, así que en todos los scripts la autenticación es manual. Una vez logueado, arranca la parte de seleccionar año y mes, enviar el formulario con JavaScript y esperar a que la página recargue por completo. Edifito comete un error grave cuando un mes no tiene información, pues en vez de mostrar un aviso o una pantalla vacía, recarga silenciosamente el último mes válido que el servidor tenga en caché. Si no validas que el mes devuelto coincide con el que pediste, terminás archivando datos que no corresponden. Los scripts chequean justo eso comparando los selectores después de cada recarga y solo siguen adelante si el período es correcto.

Los archivos que se bajan también vienen con lo suyo. Si elijes descargar las cartolas en PDF, todas se llaman gasto_comun.pdf y el sistema operativo les agrega el (1), (2) como gasto_comun(1).pdf, gasto_comun(2).pdf y así. Si seleccionas el Excel, el nombre es Gasto_Comun.xlsx sin ninguna referencia al mes ni al año, y puedes descargar el mismo período infinitas veces sin que la plataforma te advierta. Encima, al repetir una descarga, el navegador empieza a renombrar automáticamente con Copia1, Copia2, dejando un caos de nombres. Yo mismo metí la pata al no verificar si el archivo ya existía antes de volver a descargar, y después el script que ordena por fecha interna a veces leyó mal el año y terminé con archivos llamados 2019_04...abril_2026.xlsx.

Las planillas Excel que entrega Edifito son un ejemplo de manual de cómo no estructurar datos para un análisis porque:
-Encabezados que cambian de un archivo a otro
-subtotales incrustados en las mismas columnas del detalle
-jerarquías tipo 2.1, 2.1.1 donde hay filas “papito” y filas “mijito” en un mismo bloque
-celdas fusionadas a mansalva
-montos que aparecen en columnas distintas (SUBTOTAL, TOTAL, VALOR)

Sumado a la pésima labor del administrador que ingresa a los proveedores con razones sociales repetidas pero escritas de forma distinta mes a mes. 

<img width="802" height="993" alt="Captura de pantalla 2026-05-07 a la(s) 21 46 52" src="https://github.com/user-attachments/assets/c6031923-3d08-4437-bd4d-c30519c235a3" />

Todo eso obligó a postprocesar los archivos con openpyxl para leerles el contenido real y con pandas para consolidar la información financiera en una estructura plana y auditable.

# calculalosexceltotales.py
Procesa todos los archivos Excel (.xlsx) ubicados dentro de una carpeta base (estructura anidada) sin asumir un formato fijo de columnas. Lee cada hoja sin encabezado, identifica filas donde la cuarta columna contenga una fecha con guiones y la sexta columna un monto numérico, y las consolida en un único DataFrame. Luego normaliza las fechas, extrae año y mes, y calcula totales anuales y mensuales, un ranking de proveedores por monto acumulado, y detecta los meses que faltan en cada año presente. Finalmente exporta todos los resultados a un archivo Excel con hojas separadas ("base", "total_anual", "total_mensual", "top_proveedores" y "faltantes") y muestra en consola el resumen de meses ausentes.

# calculoexcelopcional.py 
Analiza archivos contables mensuales organizados en carpetas por año, donde cada archivo se nombra como "mesaño.xlsx" (ej. "enero2024.xlsx"). Para cada archivo, busca dinámicamente la posición de la tabla real dentro de las primeras 20 filas identificando una columna que contenga la palabra "VALOR". Convierte los montos desde formatos locales (puntos de miles, comas decimales) a flotantes, filtra filas con proveedor y monto positivo, y añade columnas normalizadas de período. Consolida todos los registros y genera dos análisis de salida: una evolución mensual con diferencias y porcentajes de incremento intermensual, y un ranking de proveedores con total pagado, frecuencia y promedio. El resultado se guarda en un Excel con tres hojas (datos crudos, evolución y análisis de proveedores).

# descargaBoletaFacturaAdjunta.py
Automatiza la descarga masiva de documentos adjuntos (boletas, facturas, imágenes) desde el portal de Edifito. Inicia una sesión manual del usuario mediante Selenium, luego replica las cookies de autenticación en una sesión de requests. Itera por años y meses, selecciona cada período en un formulario web, y, tras la recarga, extrae de la tabla resultante todos los enlaces a archivos mediante expresiones regulares que capturan URLs de AWS S3 y su tipo (pdf/jpg/zip). Codifica la URL y la pasa a través del endpoint verFactura.php para descargar cada archivo binario de vuelta con requests, guardándolo en una estructura de carpetas año/mes/. Si un período no tiene datos o la página no carga, continúa sin detenerse.

# descargaExcels.py 
Script de automatización con Selenium para descargar los archivos Excel de gastos comunes desde la plataforma Edifito. Requiere que el usuario inicie sesión manualmente y luego recorre todas las combinaciones de año y mes. Para cada período, navega a la página de consulta, selecciona año y mes mediante los desplegables correspondientes, envía el formulario, espera la recarga completa y verifica que la selección se haya aplicado. Si el servidor devuelve datos, ejecuta excelLey.submit() para forzar la descarga del archivo Excel (en lugar del PDF). Las descargas se guardan en una carpeta predefinida, con una pausa de 5 segundos entre cada una para evitar sobrecargas.

# descargaPDFs.py 
Similar al script anterior, pero orientado a la descarga de los archivos PDF de gastos comunes. Con la misma lógica de sesión manual y recorrido de años y meses, utiliza pdfLey.submit() en lugar del llamado al Excel. Está configurado para descargar automáticamente en la carpeta especificada sin intervención del usuario, y omite los períodos en los que no hay información disponible.

# ordenaexcel.py
Organiza un conjunto de archivos Excel de gastos comunes que originalmente están en una carpeta plana. Lee cada archivo con openpyxl en modo solo lectura y extrae todo el texto de las primeras 30 filas. Sobre ese texto busca la presencia de un año (de una lista predefinida) y un mes (en español). Si encuentra ambos, mueve el archivo a una estructura de directorios Año/Mes bajo una carpeta destino, renombrándolo a un formato estándar Gasto_Comun_Mes_Año.xlsx. Si el nombre destino ya existe, añade un sufijo incremental para evitar sobrescritura. Los archivos donde no se puede determinar el año o el mes, o que producen un error de lectura, se trasladan a una subcarpeta Sin_Clasificar.

# renombrarexcelsopcional.py
Recorre una jerarquía de carpetas por año (en mi caso es de 2018 a 2026) y renombra archivos .xlsx (excels) que no sigan la nomenclatura esperada mesaño.xlsx. Para cada archivo, intenta detectar el mes primero buscando el nombre del mes en español dentro del nombre del archivo; si no lo encuentra, extrae números de 1 al 12 (obviamente 1 siendo enero y 12 diciembre) mediante una expresión regular y lo asigna. Una vez identificado el mes y conocido el año por la carpeta contenedora, crea el nuevo nombre y renombra para evitar sobrescribir agregando un sufijo numérico si el nuevo nombre ya existe. Omite los archivos que ya tienen el formato deseado e informa aquellos cuyo mes no pudo ser determinado.

# Requisitos:

1. Para correr estos scripts lo fundamental es que tengas Python instalado en tu máquina, porque sin eso no hay arranque. Python ya trae su instalador de paquetes, así que con eso te alcanza. Después necesitás Google Chrome, porque varios de los scripts lo usan para moverse por internet y hacer descargas. También vas a precisar conexión a internet para los que entran a Edifito, y por supuesto tu usuario y contraseña para iniciar sesión en el momento que te lo pida la terminal.

2. Con eso listo, abrís una terminal y escribís pip install pandas openpyxl selenium webdriver-manager requests. Eso baja todo lo que hace falta para que los scripts funcionen. Lo demás ya viene de fábrica con Python.

3. Antes de lanzar cualquier cosa, cambiá las rutas que veas en los archivos —las que dicen /Users/ian/... o menciones a carpetas como Gastos_Comunes_Edifito— por las que vos tengas. Cada script espera los archivos de cierta manera: los que procesan planillas ya descargadas necesitan una carpeta por año con los excel adentro llamados "enero2024.xlsx" y así; el que ordena archivos sueltos, en cambio, necesita que tires todos los excel desordenados en una misma carpeta. Fijate también que donde vayas a guardar resultados o descargas tengas permiso para escribir, porque los scripts crean carpetas y mueven cosas sin avisar.

4. Cuando ejecutes los que usan Selenium, se te abre Chrome, iniciás sesión en Edifito a mano, volvés a la terminal, presionás Enter y el script arranca. No toques las pausas que traen (los time.sleep), están puestas para no exigir demasiado al servidor y que no te bloqueen. Y ojo que descargaBoletaFacturaAdjunta.py descarga de todo: PDFs, JPGs, ZIPs; los otros son especializados, solo bajan PDFs o solo Excel.

A bajar esos gastos comunes! 

# Ejemplo de mi resultado con calculalosexceltotales.py

<img width="998" height="834" alt="Captura de pantalla 2026-05-07 a la(s) 22 40 16" src="https://github.com/user-attachments/assets/d479ee4f-03a6-4c91-9b5f-102f9c9b4551" />


<img width="756" height="552" alt="Captura de pantalla 2026-05-07 a la(s) 22 41 04" src="https://github.com/user-attachments/assets/5b5ea764-8901-480f-bda8-de595d0c9fef" />


<img width="894" height="995" alt="Captura de pantalla 2026-05-07 a la(s) 22 41 50" src="https://github.com/user-attachments/assets/e3102a2d-abd1-49ff-9ff1-a83aa84502ab" />


<img width="1083" height="946" alt="Captura de pantalla 2026-05-07 a la(s) 22 43 10" src="https://github.com/user-attachments/assets/aacbadab-2e2b-4508-9294-f33bba7ed57d" />


<img width="848" height="576" alt="Captura de pantalla 2026-05-07 a la(s) 22 43 34" src="https://github.com/user-attachments/assets/34c18bee-2239-48d3-a213-70c69d43afef" />

# Ejemplo descarga cartolas

<img width="701" height="457" alt="Captura de pantalla 2026-05-07 a la(s) 22 44 31" src="https://github.com/user-attachments/assets/b316a6e7-b7b4-4f1d-8caf-a8c9dc56dafb" />


<img width="844" height="929" alt="Captura de pantalla 2026-05-07 a la(s) 22 49 04" src="https://github.com/user-attachments/assets/196498ba-b321-4aa4-a3ab-5531bd710a26" />


<img width="699" height="428" alt="Captura de pantalla 2026-05-07 a la(s) 22 49 50" src="https://github.com/user-attachments/assets/55949a97-b399-4d1e-8547-4c46de4545b0" />


# Con los excels una vez ya organizados, se puede usar software o codigo para herramientas de auditoria 


<img width="1200" height="500" alt="ANEXO 3 7 (graficoevolucion)" src="https://github.com/user-attachments/assets/959184d2-4a81-4e36-864c-aaa4feaa2eab" />


