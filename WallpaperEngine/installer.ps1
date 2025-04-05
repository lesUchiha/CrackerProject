# Configura los nombres
$zipUrl = "https://store-sa-sao-1.gofile.io/download/web/9204c5a0-e2da-4e50-8eb5-ffda18edc07e/Wallpaper.Engine.v2.5.28.rar"  # <-- Cambia esto por la URL real del ZIP
$folderName = "Wallpaper Engine"
$exeName = "wallpaper32.exe"
$destination = "$env:APPDATA\$folderName"
$shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Wallpaper Engine.lnk"

# Crear carpeta si no existe
if (-Not (Test-Path $destination)) {
    New-Item -ItemType Directory -Path $destination | Out-Null
}

# Ruta temporal para el .zip
$tempZip = "$env:TEMP\wallpaper_engine.zip"

Write-Host "Downloading Wallpaper Engine..."
Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip

Write-Host "Extracting files..."
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $destination)

# Elimina el ZIP temporal
Remove-Item $tempZip

# Crear acceso directo
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "$destination\$exeName"
$shortcut.WorkingDirectory = $destination
$shortcut.Save()

Write-Host "Wallpaper Engine has been successfully downloaded. You can find the application in the Start menu." -ForegroundColor Green
