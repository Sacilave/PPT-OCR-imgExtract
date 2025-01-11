@echo off
chcp 65001 > nul
echo ======================================
echo PPT关键词提取工具 - 自动运行脚本
echo ======================================

:: 检查input文件夹是否存在
if not exist "input" (
    echo 创建input文件夹...
    mkdir input
    echo 请将PPT文件放入input文件夹后再运行此脚本
    pause
    exit
)

:: 检查input文件夹是否为空
dir /b /a-d "input\*.*" >nul 2>&1
if errorlevel 1 (
    echo [错误] input文件夹为空！
    echo 请将PPT文件放入input文件夹后再运行此脚本
    pause
    exit
)

echo.
echo [1/3] 安装依赖...
pip install pywin32
python -m pip install paddlepaddle -i https://mirror.baidu.com/pypi/simple
python -m pip install "paddleocr>=2.0.1" -i https://mirror.baidu.com/pypi/simple

echo.
echo [2/3] 转换PPT为图片...
python ppt_to_png.py
if errorlevel 1 (
    echo [错误] PPT转换失败！
    echo 请查看 ppt_conversion.log 获取详细错误信息
    pause
    exit
)

echo.
echo [3/3] 提取关键词页面...
python ocr_process.py
if errorlevel 1 (
    echo [错误] OCR处理失败！
    echo 请查看 ocr_process.log 获取详细错误信息
    pause
    exit
)

echo.
echo ======================================
echo 处理完成！结果在 FinalOutput 文件夹中
echo ======================================
pause 