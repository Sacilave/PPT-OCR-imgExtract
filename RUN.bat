@echo off
chcp 65001 > nul

echo ======================================
echo PPT Keyword Extractor - Auto Runner
echo.
echo Plugin made by Sacilave
echo Project: https://github.com/Sacilave/PPT-OCR-imgExtract
echo ======================================
echo.

:: Check if input folder exists
if not exist "input" (
    echo Creating input folder...
    mkdir input
    echo Please put PPT files in the input folder and run this script again
    pause
    exit
)

:: Check if input folder is empty
dir /b /a-d "input\*.*" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Input folder is empty!
    echo Please put PPT files in the input folder and run this script again
    pause
    exit
)

echo.
echo [1/3] Installing dependencies... This might take a while, take a break~
pip install pywin32
python -m pip install paddlepaddle -i https://mirror.baidu.com/pypi/simple
python -m pip install "paddleocr>=2.0.1" -i https://mirror.baidu.com/pypi/simple

echo.
echo [2/3] Converting PPT to images...
python "ppt_to_png.py"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PPT conversion failed!
    echo Please check ppt_conversion.log for details
    pause
    exit /b 1
)

echo.
echo [3/3] Extracting pages with keywords...
python "ocr_process.py"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] OCR processing failed!
    echo Please check ocr_process.log for details
    pause
    exit /b 1
)

echo.
echo ======================================
echo Done! Results are in the FinalOutput folder
echo ======================================
pause