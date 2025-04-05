# Naming and path configuration
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$rarUrl       = "https://download1501.mediafire.com/kjp6zqip77qgKyV1mXPQBCie7CzyygfGXme8ATMk6AXBXacDih39Bi0l3LL7RyuhXJgDLE4uK1VY8t8TlWYvWLRuSV10NWBcu2HH106weqCakbgFg8o4H5HgW3pwoSY3zyHMeMTJiADNtfsrfXQlwoimdz03sFGXzUS28xhwYLltHUg/5ck9nxlqcg3c8rf/Wallpaper.Engine.v2.5.28.rar"
$folderName   = "Wallpaper Engine"
$exeName      = "wallpaper32.exe"
$destination  = "$env:APPDATA\$folderName"
$shortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Wallpaper Engine.lnk"

# Temporary path to save the .rar file with its real name
$tempRar = "$env:TEMP\Wallpaper.Engine.v2.5.28.rar"

# Paths to extraction programs
$winRarPath   = "C:\Program Files\WinRAR\WinRAR.exe"
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

# Function to extract with WinRAR and fallback to 7-Zip if needed
function Extract-WithFallback {
    param (
        [string]$rarFile,
        [string]$destinationFolder
    )

    # Try extracting with WinRAR
    if (Test-Path $winRarPath) {
        Write-Host "Trying to extract with WinRAR..."
        $process = Start-Process -FilePath $winRarPath -ArgumentList "x", $rarFile, "$destinationFolder\", "-y" -PassThru -Wait
        Write-Host "WinRAR exit code: $($process.ExitCode)"
        if ($process.ExitCode -eq 0) {
            Write-Host "Extraction successful using WinRAR." -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "Extraction failed with WinRAR. Exit code: $($process.ExitCode)" -ForegroundColor Red
        }
    }
    
    # If WinRAR failed, try 7-Zip
    if (Test-Path $sevenZipPath) {
        Write-Host "Trying to extract with 7-Zip..."
        $process = Start-Process -FilePath $sevenZipPath -ArgumentList "x", $rarFile, "-o$destinationFolder", "-y" -PassThru -Wait
        Write-Host "7-Zip exit code: $($process.ExitCode)"
        if ($process.ExitCode -eq 0) {
            Write-Host "Extraction successful using 7-Zip." -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "Extraction failed with 7-Zip. Exit code: $($process.ExitCode)" -ForegroundColor Red
        }
    }
    
    Write-Host "Extraction failed with both methods. Make sure you have WinRAR or 7-Zip installed." -ForegroundColor Red
    return $false
}

try {
    # Create destination folder (doesn't force any internal structure, respects the RAR content)
    if (-Not (Test-Path $destination)) {
        New-Item -ItemType Directory -Path $destination | Out-Null
    }

    Write-Host "Downloading Wallpaper Engine..."
    $headers = @{
      "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    Invoke-WebRequest -Uri $rarUrl -OutFile $tempRar -UseBasicParsing -Headers $headers -Verbose -ErrorAction Stop

    # Wait 2 seconds to ensure file was saved
    Start-Sleep -Seconds 2

    # Check that the file was downloaded correctly
    if (-Not (Test-Path $tempRar)) {
        Write-Host "Download failed. File not found at: $tempRar" -ForegroundColor Red
        throw "Downloaded file is missing or corrupt."
    }

    # Check file size (should be over 100MB)
    $downloadedSize = (Get-Item $tempRar).Length
    Write-Host "Downloaded file size: $downloadedSize bytes" -ForegroundColor Yellow
    if ($downloadedSize -lt 104857600) {
        throw "Downloaded file is too small ($downloadedSize bytes). It’s likely an HTML response instead of the actual file."
    }
    Write-Host "Download completed. File saved to: $tempRar" -ForegroundColor Green

    # Try to extract using the function
    Write-Host "Attempting to extract contents..."
    if (-Not (Extract-WithFallback -rarFile $tempRar -destinationFolder $destination)) {
        throw "Extraction failed using both methods."
    }

    # Check if files were extracted
    $extractedFiles = Get-ChildItem -Path $destination -Recurse
    if ($extractedFiles.Count -eq 0) {
        throw "No files were extracted. Check the .rar content."
    }
    else {
        Write-Host "Files successfully extracted to: $destination" -ForegroundColor Green
    }

    # Delete temporary .rar
    Remove-Item $tempRar -ErrorAction SilentlyContinue

    # Create Start Menu shortcut
    Write-Host "Creating shortcut..."
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "$destination\$exeName"
    $shortcut.WorkingDirectory = $destination
    $shortcut.Save()

    Write-Host "Wallpaper Engine has been downloaded, installed, and a shortcut has been created." -ForegroundColor Green
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
    if (Test-Path $tempRar) {
        Remove-Item $tempRar -ErrorAction SilentlyContinue
    }
}
