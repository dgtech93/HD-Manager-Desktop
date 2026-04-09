param(
    [string]$Version = "1.0.5"
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RootDir = Split-Path -Parent $ScriptDir

Write-Host "==> Build release HD Manager Desktop" -ForegroundColor Cyan
Write-Host "Root: $RootDir"
Write-Host "Version: $Version"

Set-Location $RootDir

function Resolve-IsccPath {
    $cmd = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    throw "ISCC.exe non trovato. Installa Inno Setup 6 oppure aggiungi iscc.exe al PATH."
}

Write-Host "==> Verifica dipendenze build" -ForegroundColor Yellow
python -m PyInstaller --version | Out-Null

Write-Host "==> Pulizia cartelle build" -ForegroundColor Yellow
if (Test-Path "$RootDir\build") { Remove-Item "$RootDir\build" -Recurse -Force }
if (Test-Path "$RootDir\dist\HDManagerDesktop") { Remove-Item "$RootDir\dist\HDManagerDesktop" -Recurse -Force }
if (Test-Path "$RootDir\dist-installer") { Remove-Item "$RootDir\dist-installer" -Recurse -Force }

Write-Host "==> PyInstaller (onedir)" -ForegroundColor Yellow
python -m PyInstaller `
    --noconfirm `
    --clean `
    --name HDManagerDesktop `
    --onedir `
    --windowed `
    --icon "app\assets\image.ico" `
    --add-data "app\assets;app\assets" `
    "main.py"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller terminato con errore (exit code: $LASTEXITCODE)."
}

$BuildOutput = Join-Path $RootDir "dist\HDManagerDesktop"
if (-not (Test-Path $BuildOutput)) {
    throw "Build PyInstaller non trovata in '$BuildOutput'. Verifica eventuali errori PyInstaller."
}

$BuiltExe = Join-Path $BuildOutput "HDManagerDesktop.exe"
if (-not (Test-Path $BuiltExe)) {
    throw "Exe non trovato in '$BuiltExe'. Verifica il nome output PyInstaller."
}

$IsccPath = Resolve-IsccPath
Write-Host "==> Inno Setup compile ($IsccPath)" -ForegroundColor Yellow

& $IsccPath `
    "/DAppVersion=$Version" `
    "/DBuildRoot=dist\HDManagerDesktop" `
    "$RootDir\installer.iss"
if ($LASTEXITCODE -ne 0) {
    throw "ISCC terminato con errore (exit code: $LASTEXITCODE)."
}

Write-Host ""
Write-Host "Build completata." -ForegroundColor Green
Write-Host "Installer generato in: $RootDir\dist-installer" -ForegroundColor Green
