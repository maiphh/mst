# -*- mode: python ; coding: utf-8 -*-
"""
TaxChecker PyInstaller Spec File

Bundles:
- EasyOCR models (from easyocr_models/ directory)
- ChromeDriver (from chromedriver/ directory)
- Selenium Manager (for fallback)
"""
import os
import sys

# Get project root directory
spec_dir = os.getcwd()

# =============================================================================
# Collect Selenium Manager binaries (for fallback when chromedriver not bundled)
# =============================================================================
selenium_manager_datas = []
try:
    import selenium
    selenium_dir = os.path.dirname(selenium.__file__)
    
    if sys.platform == 'darwin':
        manager_path = ('webdriver/common/macos/selenium-manager', 'selenium/webdriver/common/macos')
    elif sys.platform == 'win32':
        manager_path = ('webdriver/common/windows/selenium-manager.exe', 'selenium/webdriver/common/windows')
    else:
        manager_path = ('webdriver/common/linux/selenium-manager', 'selenium/webdriver/common/linux')
    
    src = os.path.join(selenium_dir, manager_path[0])
    if os.path.exists(src):
        selenium_manager_datas.append((src, manager_path[1]))
        print(f"Bundling Selenium Manager: {src}")
except ImportError:
    print("WARNING: Selenium not found, skipping Selenium Manager bundling")

# =============================================================================
# Collect EasyOCR models from easyocr_models/ directory
# =============================================================================
easyocr_model_datas = []
model_dir = os.path.join(spec_dir, 'easyocr_models')
if os.path.isdir(model_dir):
    for f in os.listdir(model_dir):
        if f.endswith('.pth'):
            src = os.path.join(model_dir, f)
            easyocr_model_datas.append((src, 'easyocr_models'))
            print(f"Bundling EasyOCR model: {f}")
else:
    print("WARNING: easyocr_models/ directory not found")

# =============================================================================
# Collect ChromeDriver binaries from chromedriver/ directory
# =============================================================================
chromedriver_datas = []
chromedriver_dir = os.path.join(spec_dir, 'chromedriver')

if sys.platform == 'darwin':
    chromedriver_src = os.path.join(chromedriver_dir, 'macos', 'chromedriver')
elif sys.platform == 'win32':
    chromedriver_src = os.path.join(chromedriver_dir, 'windows', 'chromedriver.exe')
else:
    chromedriver_src = os.path.join(chromedriver_dir, 'linux', 'chromedriver')

if os.path.exists(chromedriver_src):
    chromedriver_datas.append((chromedriver_src, 'chromedriver'))
    print(f"Bundling ChromeDriver: {chromedriver_src}")
else:
    print(f"WARNING: ChromeDriver not found at {chromedriver_src}")

# =============================================================================
# PyInstaller Analysis
# =============================================================================
a = Analysis(
    ['gui_app_qt.py'],
    pathex=[],
    binaries=[],
    datas=selenium_manager_datas + easyocr_model_datas + chromedriver_datas,
    hiddenimports=[
        'selenium.webdriver.common.service',
        'easyocr.easyocr',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude unused heavy modules to reduce size
        'torch.utils.benchmark',
        'tkinter',
        'matplotlib',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TaxChecker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TaxChecker',
)

# macOS app bundle
app = BUNDLE(
    coll,
    name='TaxChecker.app',
    icon=None,
    bundle_identifier='com.taxchecker.app',
)
