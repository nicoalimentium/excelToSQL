$carpeta = ""

Get-ChildItem -Path $carpeta -Filter "*.ps1" | ForEach-Object {
    if ($_.Name -ne "ejecuta-todos-los-scripts.ps1") {
		Write-Host "Ejecutando script: $($_.Name)"
		try{
			pwsh.exe -File $_.FullName
		}
		catch{
			Write-Error $_.Exception.Message
		}
		Write-Host "Script finalizado: $($_.Name)"		
	}
	else{
		Write-Host "Omitiendo script: $($_.Name)"
	}
}
