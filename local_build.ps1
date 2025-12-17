# local_build.ps1
# Automates the setup and build process for TaxChecker locally

Write-Host "Starting Local Build Process..." -ForegroundColor Cyan

# 1. Install Dependencies
Write-Host "Installing Python dependencies..."
pip install -r requirements.txt
pip install pyinstaller pyinstaller-hooks-contrib

# 2. Download RapidOCR Models (English - better for captcha)
$modelDir = "models"
if (-not (Test-Path $modelDir)) {
    New-Item -ItemType Directory -Force -Path $modelDir | Out-Null
}
Write-Host "Downloading RapidOCR English models (if missing)..."
$detPath = "$modelDir/en_PP-OCRv3_det_infer.onnx"
$recPath = "$modelDir/en_PP-OCRv4_rec_infer.onnx"

if (-not (Test-Path $detPath)) {
    Write-Host "Downloading English detection model..."
    Invoke-WebRequest -Uri "https://huggingface.co/SWHL/RapidOCR/resolve/main/PP-OCRv3/en/en_PP-OCRv3_det_infer.onnx" -OutFile $detPath
}
if (-not (Test-Path $recPath)) {
    Write-Host "Downloading English recognition model..."
    Invoke-WebRequest -Uri "https://huggingface.co/SWHL/RapidOCR/resolve/main/PP-OCRv4/en/en_PP-OCRv4_rec_infer.onnx" -OutFile $recPath
}

# 3. Validate Bundled Browser
$chromePath = "bin/chrome-win64/chrome.exe"
$chromedriverPath = "bin/chromedriver.exe"

Write-Host "Validating bundled browser..."
if (-not (Test-Path $chromePath)) {
    Write-Host "ERROR: Bundled Chrome not found at $chromePath" -ForegroundColor Red
    Write-Host "Please download Chrome for Testing and extract to bin/chrome-win64/" -ForegroundColor Yellow
    exit 1
}
if (-not (Test-Path $chromedriverPath)) {
    Write-Host "ERROR: ChromeDriver not found at $chromedriverPath" -ForegroundColor Red
    Write-Host "Please download ChromeDriver and place at bin/chromedriver.exe" -ForegroundColor Yellow
    exit 1
}
Write-Host "Bundled browser validated: $chromePath" -ForegroundColor Green

# 4. Run PyInstaller
Write-Host "Running PyInstaller..."
pyinstaller --noconfirm TaxChecker.spec

if (Test-Path "dist/TaxChecker/TaxChecker.exe") {
    Write-Host "Build Successful! Output at dist/TaxChecker/TaxChecker.exe" -ForegroundColor Green
} else {
    Write-Host "Build Failed!" -ForegroundColor Red
}
