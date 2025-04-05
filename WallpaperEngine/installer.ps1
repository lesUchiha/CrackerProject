# Configuración de nombres y rutas
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$rarUrl       = "https://store-sa-sao-1.gofile.io/download/web/9204c5a0-e2da-4e50-8eb5-ffda18edc07e/Wallpaper.Engine.v2.5.28.rar"
$folderName   = "Wallpaper Engine"
$exeName      = "wallpaper32.exe"
$destination  = "$env:APPDATA\$folderName"
$shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Wallpaper Engine.lnk"

# Ruta temporal para guardar el archivo .rar con su nombre real
$tempRar = "$env:TEMP\Wallpaper.Engine.v2.5.28.rar"

# Rutas a los programas de extracción
$winRarPath   = "C:\Program Files\WinRAR\WinRAR.exe"
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

# Función que intenta extraer con WinRAR y, si falla, con 7-Zip
function Extract-WithFallback {
    param (
        [string]$rarFile,
        [string]$destinationFolder
    )

    # Intentar extraer con WinRAR
    if (Test-Path $winRarPath) {
        Write-Host "Intentando extraer con WinRAR..."
        $process = Start-Process -FilePath $winRarPath -ArgumentList "x", $rarFile, "$destinationFolder\", "-y" -PassThru -Wait
        Write-Host "Código de salida de WinRAR: $($process.ExitCode)"
        if ($process.ExitCode -eq 0) {
            Write-Host "Extracción exitosa usando WinRAR." -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "Fallo en la extracción con WinRAR. Código de salida: $($process.ExitCode)" -ForegroundColor Red
        }
    }
    
    # Si falla WinRAR, intenta extraer con 7-Zip
    if (Test-Path $sevenZipPath) {
        Write-Host "Intentando extraer con 7-Zip..."
        $process = Start-Process -FilePath $sevenZipPath -ArgumentList "x", $rarFile, "-o$destinationFolder", "-y" -PassThru -Wait
        Write-Host "Código de salida de 7-Zip: $($process.ExitCode)"
        if ($process.ExitCode -eq 0) {
            Write-Host "Extracción exitosa usando 7-Zip." -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "Fallo en la extracción con 7-Zip. Código de salida: $($process.ExitCode)" -ForegroundColor Red
        }
    }
    
    Write-Host "La extracción falló con ambos métodos. Asegúrate de tener instalado WinRAR o 7-Zip." -ForegroundColor Red
    return $false
}

try {
    # Crear la carpeta de destino (no forzamos que contenga los archivos internos, se respeta la estructura del RAR)
    if (-Not (Test-Path $destination)) {
        New-Item -ItemType Directory -Path $destination | Out-Null
    }

    Write-Host "Descargando Wallpaper Engine..."
    $headers = @{
      "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    # Forzamos UseBasicParsing y establecemos encabezados
    Invoke-WebRequest -Uri $rarUrl -OutFile $tempRar -UseBasicParsing -Headers $headers -Verbose -ErrorAction Stop

    # Esperar 2 segundos para confirmar que se guardó el archivo
    Start-Sleep -Seconds 2

    # Verificar que el archivo se descargó correctamente
    if (-Not (Test-Path $tempRar)) {
        Write-Host "La descarga falló. No se encontró el archivo en: $tempRar" -ForegroundColor Red
        throw "El archivo descargado está ausente o corrupto."
    }

    # Verificar el tamaño del archivo descargado (ejemplo: debe ser mayor a 100 MB)
    $downloadedSize = (Get-Item $tempRar).Length
    Write-Host "Tamaño del archivo descargado: $downloadedSize bytes" -ForegroundColor Yellow
    # 100 MB = 104857600 bytes (ajusta este valor si es necesario)
    if ($downloadedSize -lt 104857600) {
        throw "El archivo descargado es muy pequeño (solo $downloadedSize bytes). Probablemente se descargó una respuesta HTML en lugar del archivo real."
    }
    Write-Host "Descarga completada. Archivo guardado en: $tempRar" -ForegroundColor Green

    # Intentar extraer el contenido usando la función de extracción
    Write-Host "Intentando extraer el contenido..."
    if (-Not (Extract-WithFallback -rarFile $tempRar -destinationFolder $destination)) {
        throw "La extracción falló usando ambos métodos."
    }

    # Verificar que se hayan extraído archivos (incluyendo subcarpetas)
    $extractedFiles = Get-ChildItem -Path $destination -Recurse
    if ($extractedFiles.Count -eq 0) {
        throw "No se extrajeron archivos. Revisa el contenido del .rar."
    }
    else {
        Write-Host "Archivos extraídos exitosamente en: $destination" -ForegroundColor Green
    }

    # Eliminar el archivo .rar temporal
    Remove-Item $tempRar -ErrorAction SilentlyContinue

    # Crear el acceso directo en el menú de inicio
    Write-Host "Creando acceso directo..."
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "$destination\$exeName"
    $shortcut.WorkingDirectory = $destination
    $shortcut.Save()

    Write-Host "Wallpaper Engine se ha descargado, instalado y se ha creado un acceso directo." -ForegroundColor Green
}
catch {
    Write-Host "Ocurrió un error: $_" -ForegroundColor Red
    if (Test-Path $tempRar) {
        Remove-Item $tempRar -ErrorAction SilentlyContinue
    }
}
