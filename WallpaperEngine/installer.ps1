# Configura los nombres
$rarUrl = "https://store-sa-sao-1.gofile.io/download/web/9204c5a0-e2da-4e50-8eb5-ffda18edc07e/Wallpaper.Engine.v2.5.28.rar"
$folderName = "Wallpaper Engine"
$exeName = "wallpaper32.exe"
$destination = "$env:APPDATA\$folderName"
$shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Wallpaper Engine.lnk"

# Ruta temporal para el .rar
$tempRar = "$env:TEMP\wallpaper_engine.rar"

# Rutas a WinRAR y 7-Zip
$winRarPath = "C:\Program Files\WinRAR\WinRAR.exe"
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

# Función para intentar con WinRAR y luego con 7-Zip
function Extract-WithFallback {
    param (
        [string]$rarFile,
        [string]$destinationFolder
    )

    # Intentar con WinRAR primero
    if (Test-Path $winRarPath) {
        Write-Host "Attempting to extract with WinRAR..."
        $process = Start-Process -FilePath $winRarPath -ArgumentList "x", $rarFile, "$destinationFolder\", "-y" -PassThru -Wait
        if ($process.ExitCode -eq 0) {
            Write-Host "Extraction successful using WinRAR." -ForegroundColor Green
            return $true
        } else {
            Write-Host "WinRAR extraction failed. Exit code: $($process.ExitCode)" -ForegroundColor Red
        }
    }

    # Si falló con WinRAR, intentar con 7-Zip
    if (Test-Path $sevenZipPath) {
        Write-Host "Attempting to extract with 7-Zip..."
        $process = Start-Process -FilePath $sevenZipPath -ArgumentList "x", $rarFile, "-o$destinationFolder", "-y" -PassThru -Wait
        if ($process.ExitCode -eq 0) {
            Write-Host "Extraction successful using 7-Zip." -ForegroundColor Green
            return $true
        } else {
            Write-Host "7-Zip extraction failed. Exit code: $($process.ExitCode)" -ForegroundColor Red
        }
    }

    # Si ambos fallaron
    Write-Host "Both WinRAR and 7-Zip extraction failed. Please ensure that one of them is installed." -ForegroundColor Red
    return $false
}

try {
    # Descargar el archivo .rar
    Write-Host "Downloading Wallpaper Engine..."
    Invoke-WebRequest -Uri $rarUrl -OutFile $tempRar -ErrorAction Stop

    # Intentar extraer con WinRAR o 7-Zip
    if (-Not (Extract-WithFallback -rarFile $tempRar -destinationFolder $destination)) {
        throw "Extraction failed with both WinRAR and 7-Zip."
    }

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
