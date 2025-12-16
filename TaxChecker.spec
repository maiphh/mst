# -*- mode: python ; coding: utf-8 -*-
"""
TaxChecker PyInstaller Spec File

Bundles:
- EasyOCR models (from easyocr_models/ directory)
- ChromeDriver (from chromedriver/ directory)
- Selenium Manager (for fallback)

Fixes Applied:
- UPX disabled to prevent DLL corruption on Windows
- "Trust Only System" SSL filter to prevent library conflicts
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

# =============================================================================
# "Trust Only System" SSL Filter
# Remove ALL SSL libraries except those from Python's official DLLs folder
# This prevents conflicts from cv2, PyQt6, urllib3, etc.
# =============================================================================
python_dll_dir = os.path.join(sys.base_prefix, 'DLLs') if sys.platform == 'win32' else ''
if sys.platform == 'darwin':
    # On macOS, trust libraries from Homebrew or system paths
    trusted_ssl_paths = ['/opt/homebrew', '/usr/local', '/usr/lib']
else:
    trusted_ssl_paths = [python_dll_dir]

print(f"Trusting SSL libraries only from: {trusted_ssl_paths}")

new_binaries = []
excluded_count = 0

for (name, path, typecode) in a.binaries:
    name_lower = name.lower()
    
    # Identify SSL libraries
    is_ssl = "libssl" in name_lower or "libcrypto" in name_lower
    
    if is_ssl and path:
        path_lower = path.lower()
        # Check if the file is from a trusted location
        is_trusted = any(trusted.lower() in path_lower for trusted in trusted_ssl_paths if trusted)
        
        if not is_trusted:
            print(f"EXCLUDING: {name} from {path}")
            excluded_count += 1
            continue
        else:
            print(f"KEEPING trusted SSL: {name} from {path}")
            
    new_binaries.append((name, path, typecode))

print(f"Total conflicting SSL libraries removed: {excluded_count}")

# Inject correct OpenSSL libraries on macOS from Homebrew
if sys.platform == 'darwin':
    homebrew_ssl = '/opt/homebrew/opt/openssl@3/lib'
    if os.path.exists(homebrew_ssl):
        libssl = os.path.join(homebrew_ssl, 'libssl.3.dylib')
        libcrypto = os.path.join(homebrew_ssl, 'libcrypto.3.dylib')
        if os.path.exists(libssl) and os.path.exists(libcrypto):
            print(f"INJECTING Homebrew OpenSSL: {libssl}")
            new_binaries.append(('libssl.3.dylib', libssl, 'BINARY'))
            new_binaries.append(('libcrypto.3.dylib', libcrypto, 'BINARY'))
        else:
            print(f"WARNING: Homebrew OpenSSL files not found")
    else:
        # Try Intel Mac path
        intel_ssl = '/usr/local/opt/openssl@3/lib'
        if os.path.exists(intel_ssl):
            libssl = os.path.join(intel_ssl, 'libssl.3.dylib')
            libcrypto = os.path.join(intel_ssl, 'libcrypto.3.dylib')
            if os.path.exists(libssl) and os.path.exists(libcrypto):
                print(f"INJECTING Intel Homebrew OpenSSL: {libssl}")
                new_binaries.append(('libssl.3.dylib', libssl, 'BINARY'))
                new_binaries.append(('libcrypto.3.dylib', libcrypto, 'BINARY'))

a.binaries = new_binaries

# =============================================================================
# Executable
# =============================================================================
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TaxChecker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # DISABLED: UPX can corrupt DLLs causing "Invalid memory location" on Windows
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
    upx=False,  # DISABLED: UPX can corrupt DLLs
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
