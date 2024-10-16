# Obtiene la hora de inicio
$horaInicio = Get-Date

Write-Host "Hora de inicio: $($horaInicio.ToString('HH:mm:ss'))"

# Ruta de la carpeta que contiene los archivos Excel
$carpetaExcel = "\excels"

# Crea un objeto Excel
$excel = New-Object -ComObject Excel.Application

# Obtiene la lista de archivos Excel en la carpeta
$archivos = Get-ChildItem $carpetaExcel -Filter *.xlsx

Write-Host "Comienza la búsqueda en $carpetaExcel"
#Contador de scripts
$contador = 0
# Recorre cada archivo Excel en la carpeta
foreach ($archivo in $archivos) {
	# Ruta y nombre del archivo de texto plano de salida
	$archivoSalida = ""
	# Inicializa una cadena para almacenar el contenido de la columna
	$contenidoColumna = "/*
 * (EST): Estructura. 
 * (DAT): Modificación Datos.
 * (QRY): Consultas.
*/
-------------------------------------------------------------------------------------
/*
 * LINK TAREA: 
 * DESCRIPCIÓN: 
 * 
 *
 * AUTOR: Nicolangelo Famiglietti Acuña
 * FECHA CREACIÓN: 
 * FECHA DESPLIEGUE DESARROLLO: 	
 * FECHA DESPLIEGUE PRE-PRODUCCIÓN: 
 * FECHA DESPLIEGUE PRODUCCIÓN:
*/
-------------------------------------------------------------------------------------
---
-------------------------------------------------
--- rollback
-------------------------------------------------
BEGIN TRAN
"
    $workbook = $excel.Workbooks.Open($archivo.FullName)
    Write-Host "Excel: $($archivo.Name)"
	$contadorLineasTotales = 0
    # Recorre cada hoja en el libro
    foreach ($worksheet in $workbook.Worksheets) {
		# Nombre del archivo sin la extensión
		$nombreArchivoSinExtension = [System.IO.Path]::GetFileNameWithoutExtension($archivo.Name)

		# Columna que deseas copiar (por ejemplo, "A" para la primera columna)
		$columnaACopiar = "W"
				
		# Obtiene el rango de la columna especificada
		$range = $worksheet.Range("${columnaACopiar}:${columnaACopiar}")
		# Encuentra la última fila con contenido en la columna
		$lastRow = $worksheet.UsedRange.Rows.Count
		#Write-Host "Cantidad de filas encontradas: ${lastRow}"
		# Recorre solo las celdas con contenido en la columna
		Write-Host "Hoja: $($worksheet.Name)"
		$nomHoja = 1
		for ($i = 1; $i -le $lastRow; $i++) {
			$cell = $range.Item($i, 1)
			$contenidoCelda = $cell.Text
			# Verifica si la longitud del contenido de la celda es mayor que 70
            if ($contenidoCelda.Length -gt 70 -or $contenidoCelda -like '*DECLARE*') {
				if($nomHoja -eq 1){
					# Agrega un comentario con el nombre de la hoja
					$contenidoColumna += "---Tabla: $($worksheet.Name)`r`n"	
					$nomHoja = 0
				}
				# Reemplaza las apariciones de '' por NULL
				$contenidoCelda = $contenidoCelda -replace "''", "NULL" 
				# Reemplaza las apariciones de 'NULL' por NULL
				$contenidoCelda = $contenidoCelda -replace "'NULL'", "NULL" 				
				# Reemplaza las apariciones de %% por ''
				$contenidoCelda = $contenidoCelda -replace "%%", "''"  
				# Reemplaza las apariciones de $$ por salto de línea
				$contenidoCelda = $contenidoCelda -replace '\$\$', "`r`n"
				# Reemplaza las palabras que contengan ' por doble '
				$contenidoCelda = $contenidoCelda -replace "(?<=\w)'(?=\w)", "''"
				$contenidoColumna += $contenidoCelda + "`r`n" # Agrega un salto de línea
				$contadorLineasTotales++
				# Verifica si $contadorLineasTotales es múltiplo de 45 y agrega "GO"
				if ($contadorLineasTotales % 45 -eq 0) {
					$contenidoColumna += "GO`r`n"
					foreach ($variable in $variables) {
						$contenidoColumna += $variable + "`r`n"
					}					
				}
			}
		}
    }
	$contenidoColumna += "GO`r`n"
	$contenidoColumna += "commit`r`n"
	Write-Host "Contador de scripts generados: $($contador)"
	# Crea un archivo de texto con el contenido de la columna
	$archivoSalida = "\script generado\20240718-XXXX-001-DAT-$($nombreArchivoSinExtension) - correcciones.sql"
	$contador++
	$contenidoColumna | Out-File -FilePath $archivoSalida
	Write-Host "***--Cerramos el archivo $($archivo.Name) y liberamos memoria"
    # Cierra el libro de trabajo sin guardar cambios
	$workbook.Close($false)
	# Libera los objetos COM
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range)
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
}
Write-Host "***--Cierra la aplicación Excel y terminamos el proceso"
Write-Host "***--FIN"
# Cierra la aplicación Excel
$excel.Quit()
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
# Obtiene la hora de fin
$horaFin = Get-Date

Write-Host "Hora de fin: $($horaFin.ToString('HH:mm:ss'))"

# Calcula y muestra la duración de la ejecución
$duracion = $horaFin - $horaInicio
Write-Host "Duración de la ejecución: $($duracion)"