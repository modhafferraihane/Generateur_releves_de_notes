# Installe et lance le Generateur de releves de notes depuis GitHub, sans
# rien telecharger manuellement. A executer dans PowerShell avec :
#
#   irm https://raw.githubusercontent.com/modhafferraihane/Generateur_releves_de_notes/main/install.ps1 | iex
#
$ErrorActionPreference = "Stop"

$RepoZipUrl = "https://github.com/modhafferraihane/Generateur_releves_de_notes/archive/refs/heads/main.zip"
$Dest = Join-Path $env:USERPROFILE "Generateur_releves_de_notes"

Write-Host "Telechargement du programme..." -ForegroundColor Cyan
$zipPath = Join-Path $env:TEMP "releves_install.zip"
Invoke-WebRequest -Uri $RepoZipUrl -OutFile $zipPath -UseBasicParsing

$extractPath = Join-Path $env:TEMP "releves_install_extract"
if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
$extractedFolder = Get-ChildItem $extractPath | Select-Object -First 1

if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
Move-Item $extractedFolder.FullName $Dest
Remove-Item $zipPath, $extractPath -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Verification de Python..." -ForegroundColor Cyan
function Find-PythonExe {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd -and ((& $cmd.Source --version) -notmatch "0\.0\.0\.0")) { return $cmd.Source }
    $installed = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "Python3*" } | Sort-Object Name -Descending | Select-Object -First 1
    if ($installed) { return (Join-Path $installed.FullName "python.exe") }
    return $null
}

$PythonExe = Find-PythonExe
if (-not $PythonExe) {
    Write-Host "Python n'est pas installe : installation en cours (ca peut prendre quelques minutes)..." -ForegroundColor Cyan
    winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
    $PythonExe = Find-PythonExe
    if (-not $PythonExe) {
        Write-Host "Python a ete installe mais n'a pas pu etre localise automatiquement." -ForegroundColor Yellow
        Write-Host "Fermez PowerShell, rouvrez-le, et relancez la commande d'installation." -ForegroundColor Yellow
        exit 1
    }
}

Write-Host "Installation des dependances (premiere fois seulement)..." -ForegroundColor Cyan
& $PythonExe -m pip install --quiet --disable-pip-version-check -r (Join-Path $Dest "requirements.txt")

# Raccourci pour relancer le programme plus tard sans repasser par PowerShell.
$LauncherPath = Join-Path $Dest "Lancer le programme.bat"
@"
@echo off
cd /d "%~dp0"
start "" cmd /c "timeout /t 2 >nul & start http://127.0.0.1:5000"
"$PythonExe" app.py
pause
"@ | Set-Content -Path $LauncherPath -Encoding ASCII

Copy-Item $LauncherPath (Join-Path ([Environment]::GetFolderPath("Desktop")) "Generateur de releves de notes.bat") -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Installation terminee !" -ForegroundColor Green
Write-Host "Lancement du programme..." -ForegroundColor Cyan
Start-Process -FilePath $PythonExe -ArgumentList "app.py" -WorkingDirectory $Dest
Start-Sleep -Seconds 3
Start-Process "http://127.0.0.1:5000"

Write-Host ""
Write-Host "Le site est ouvert dans votre navigateur (http://127.0.0.1:5000)." -ForegroundColor Green
Write-Host "La prochaine fois, double-cliquez sur l'icone 'Generateur de releves de notes' sur le Bureau." -ForegroundColor Green
