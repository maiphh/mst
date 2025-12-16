# local_build.ps1
# Automates the setup and build process for TaxChecker locally

Write-Host "Starting Local Build Process..." -ForegroundColor Cyan

# 1. Install Dependencies
Write-Host "Installing Python dependencies..."
pip install -r requirements.txt
pip install pyinstaller

# 2. Download RapidOCR Models
$modelDir = "models"
if (-not (Test-Path $modelDir)) {
    New-Item -ItemType Directory -Force -Path $modelDir | Out-Null
}
Write-Host "Downloading RapidOCR models (if missing)..."
$detPath = "$modelDir/en_PP-OCRv3_det_infer.onnx"
$recPath = "$modelDir/en_PP-OCRv4_rec_infer.onnx"

if (-not (Test-Path $detPath)) {
    Write-Host "Downloading detection model..."
    Invoke-WebRequest -Uri "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.4.0/onnx/PP-OCRv4/det/en_PP-OCRv3_det_infer.onnx" -OutFile $detPath
}
if (-not (Test-Path $recPath)) {
    Write-Host "Downloading recognition model..."
    Invoke-WebRequest -Uri "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.4.0/onnx/PP-OCRv4/rec/en_PP-OCRv4_rec_infer.onnx" -OutFile $recPath
}

# 3. Run PyInstaller
Write-Host "Running PyInstaller..."
pyinstaller --noconfirm TaxChecker.spec

if (Test-Path "dist/TaxChecker/TaxChecker.exe") {
    Write-Host "Build Successful! Output at dist/TaxChecker/TaxChecker.exe" -ForegroundColor Green
} else {
    Write-Host "Build Failed!" -ForegroundColor Red
}
