@echo off
setlocal

:: Step 0: check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Install Python 3.10+ from https://www.python.org/
    pause
    exit /b 1
)

:: Step 1: upgrade pip
python -m pip install --upgrade pip --quiet

:: Step 2: PyTorch CUDA 12.1
echo.
echo [1/3] Installing PyTorch CUDA 12.1 ...
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu121
if %errorlevel% neq 0 (
    echo [WARN] CUDA 12.1 failed. Trying CUDA 11.8 ...
    pip install torch torchvision --index-url https://download.pytorch.org/whl/cu118
)
if %errorlevel% neq 0 (
    echo [WARN] CUDA failed. Installing CPU-only version ...
    pip install torch torchvision
)

:: Step 3: other dependencies
echo.
echo [2/3] Installing dependencies ...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)

:: Step 4: GPU check (messages printed by Python)
echo.
echo [3/3] GPU check ...
python check_gpu.py

:: Step 5: download model?
echo.
python installer_menu.py
if %errorlevel% equ 1 (
    python download_model.py drct
)

echo.
python -c "print('[Done] Setup complete. Run: run.bat')"
echo.
pause
