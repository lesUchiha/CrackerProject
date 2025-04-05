# Configura los nombres
$rarUrl = "https://store-sa-sao-1.gofile.io/download/web/9204c5a0-e2da-4e50-8eb5-ffda18edc07e/Wallpaper.Engine.v2.5.28.rar"
$folderName = "Wallpaper Engine"
$exeName = "wallpaper32.exe"
$destination = "$env:APPDATA\$folderName"
$shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Wallpaper Engine.lnk"

# Ruta temporal para el .rar
$tempRar = "$env:TEMP\wallpaper_engine.rar"

# Ruta a 7z.exe (asegúrate de que 7-Zip esté instalado y su ruta esté en el PATH)
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

# Función para instalar 7-Zip si no está instalado
function Install-7Zip {
    Write-Host "7-Zip is not installed. Installing 7-Zip..."
    $installerUrl = "https://www.7-zip.org/a/7z1900-x64.exe"
    $installerPath = "$env:TEMP\7z_installer.exe"
    
    # Descargar el instalador de 7-Zip
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop
    
    # Instalar 7-Zip
    Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait
    
    # Verificar que 7-Zip se haya instalado correctamente
    if (Test-Path $sevenZipPath) {
        Write-Host "7-Zip installed successfully." -ForegroundColor Green
    } else {
        throw "Failed to install 7-Zip."
    }
}

try {
    # Crear carpeta si no existe
    if (-Not (Test-Path $destination)) {
        New-Item -ItemType Directory -Path $destination | Out-Null
    }

    # Descargar el archivo .rar
    Write-Host "Downloading Wallpaper Engine..."
    Invoke-WebRequest -Uri $rarUrl -OutFile $tempRar -ErrorAction Stop

    # Verificar si 7z.exe existe
    if (-Not (Test-Path $sevenZipPath)) {
        Install-7Zip  # Instalar 7-Zip si no se encuentra
    }

    # Extraer el archivo .rar
    Write-Host "Extracting files using 7-Zip..."
    Start-Process -FilePath $sevenZipPath -ArgumentList "x", $tempRar, "-o$destination", "-y" -Wait

    # Eliminar el archivo .rar temporal
    Remove-Item $tempRar -ErrorAction SilentlyContinue

    # Crear acceso directo
    Write-Host "Creating shortcut..."
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "$destination\$exeName"
    $shortcut.WorkingDirectory = $destination
    $shortcut.Save()

    Write-Host "Wallpaper Engine has been successfully downloaded and installed." -ForegroundColor Green
} catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    if (Test-Path $tempRar) {
        Remove-Item $tempRar -ErrorAction SilentlyContinue
    }
}
