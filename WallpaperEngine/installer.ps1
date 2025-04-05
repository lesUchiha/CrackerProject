# Configura los nombres
$rarUrl = "https://store-sa-sao-1.gofile.io/download/web/9204c5a0-e2da-4e50-8eb5-ffda18edc07e/Wallpaper.Engine.v2.5.28.rar"  # <-- Cambia esto por la URL real del RAR
$folderName = "Wallpaper Engine"
$exeName = "wallpaper32.exe"
$destination = "$env:APPDATA\$folderName"
$shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Wallpaper Engine.lnk"

# Crear carpeta si no existe
if (-Not (Test-Path $destination)) {
    New-Item -ItemType Directory -Path $destination | Out-Null
}

# Ruta temporal para el .rar
$tempRar = "$env:TEMP\wallpaper_engine.rar"

Write-Host "Downloading Wallpaper Engine..."
Invoke-WebRequest -Uri $rarUrl -OutFile $tempRar

Write-Host "Extracting files using 7-Zip..."
# Ruta a 7z.exe (asegúrate de que 7-Zip esté instalado y su ruta esté en el PATH)
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

# Extraer el archivo .rar
Start-Process -FilePath $sevenZipPath -ArgumentList "x", $tempRar, "-o$destination", "-y" -Wait

# Elimina el archivo .rar temporal
Remove-Item $tempRar

# Crear acceso directo
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "$destination\$exeName"
$shortcut.WorkingDirectory = $destination
$shortcut.Save()

Write-Host "Wallpaper Engine has been successfully downloaded. You can find the application in the Start menu." -ForegroundColor Green
