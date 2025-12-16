# local_build.ps1
# Automates the setup and build process for TaxChecker locally

Write-Host "Starting Local Build Process..." -ForegroundColor Cyan

# 1. Install Dependencies
Write-Host "Installing Python dependencies..."
pip install -r requirements.txt
pip install pyinstaller

# 2. Download EasyOCR Models
$modelDir = "easyocr_models"
if (-not (Test-Path $modelDir)) {
    New-Item -ItemType Directory -Force -Path $modelDir | Out-Null
}
Write-Host "Downloading EasyOCR models (if missing)..."
$g2Path = "$modelDir/english_g2.pth"
$mltPath = "$modelDir/craft_mlt_25k.pth"

if (-not (Test-Path $g2Path)) {
    Write-Host "Downloading english_g2.pth..."
    Invoke-WebRequest -Uri "https://huggingface.co/xiaoyao9184/easyocr/resolve/main/english_g2.pth" -OutFile $g2Path
}
if (-not (Test-Path $mltPath)) {
    Write-Host "Downloading craft_mlt_25k.pth..."
    Invoke-WebRequest -Uri "https://huggingface.co/xiaoyao9184/easyocr/resolve/main/craft_mlt_25k.pth" -OutFile $mltPath
}

# 3. Download EdgeDriver
$driverDir = "edgedriver/windows"
$driverFile = "$driverDir/msedgedriver.exe"
if (-not (Test-Path $driverDir)) {
    New-Item -ItemType Directory -Force -Path $driverDir | Out-Null
}

if (-not (Test-Path $driverFile)) {
    Write-Host "Downloading EdgeDriver..."
    $edgeVersion = "131.0.2903.99"
    $url = "https://msedgedriver.azureedge.net/$edgeVersion/edgedriver_win64.zip"
    $zipPath = "edgedriver.zip"
    
    Invoke-WebRequest -Uri $url -OutFile $zipPath
    Expand-Archive -Path $zipPath -DestinationPath "edgedriver_temp" -Force
    Move-Item "edgedriver_temp/msedgedriver.exe" $driverDir -Force
    Remove-Item "edgedriver_temp" -Recurse -Force
    Remove-Item $zipPath -Force
} else {
    Write-Host "EdgeDriver already exists."
}

# 4. Run PyInstaller
Write-Host "Running PyInstaller..."
pyinstaller --noconfirm TaxChecker.spec

if (Test-Path "dist/TaxChecker/TaxChecker.exe") {
    Write-Host "Build Successful! Output at dist/TaxChecker/TaxChecker.exe" -ForegroundColor Green
} else {
    Write-Host "Build Failed!" -ForegroundColor Red
}
