# PPT 关键词和信息提取输出图片

> PPT 提取包含指定关键词的页面，将 PPT 转换为图片。可用于提取 PPT 课件中的复习题、知识点、重点。
> PPT 划重点工具，PPT 文字识别, PPT 图片提取, 考试周神器, PPT内容查询

考试周要复习了，ppt 课件太多，一个一个翻太麻烦了，于是写了这个工具，分成两个脚本

一个将 ppt 转换为图片，一个将图片中的文字识别并提取包含关键词的页面，最终批量提取全部复习题

> 👍 使用者请直接跳到 [快速开始](#快速开始) 

## 目录
- [功能特点](#功能特点)
- [技术栈](#技术栈)
- [快速开始](#快速开始)
- [详细说明](#详细说明)
    - [安装步骤](#安装步骤)
    - [使用方法](#使用方法)
    - [自定义关键词](#自定义关键词)
- [项目结构](#项目结构)
- [常见问题](#常见问题)

## 功能特点

  - 批量处理多个PPT文件
  - 自动将PPT文件转换为高质量PNG图片
  - 智能识别图片中的中文文字
  - 支持自定义关键词检测
  - 自动提取包含关键词的页面
  - 详细的处理日志记录

## 技术栈

  - **PPT处理**: pywin32
  - **OCR引擎**: [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR)
    - 基于百度飞桨深度学习框架
    - 支持中文文字识别
    - 准确率高，速度快
    - 支持文本方向分类
    - 可离线使用

## 快速开始

> 💡 推荐使用自动运行脚本：直接双击运行 `run_all.bat`

1. **环境准备**
    ```bash
    # 安装所需依赖
    pip install pywin32
    python -m pip install paddlepaddle -i https://mirror.baidu.com/pypi/simple
    python -m pip install "paddleocr>=2.0.1" -i https://mirror.baidu.com/pypi/simple
    ```

2. **使用步骤**
    1. 将PPT文件放入 `input` 文件夹
    2. 将PPT转换为PNG图片，运行 
    ```python
    python ppt_to_png.py
    ```
    
    3. 将PNG图片中的文字识别并提取包含关键词的页面，运行 
    ```python
    python ocr_process.py
    ```
    4. 在 `FinalOutput` 文件夹中查看结果

## 详细说明

### 安装步骤

  1. 确保系统已安装Python 3.7+
  2. 确保系统已安装Microsoft PowerPoint
  3. 下载本项目代码
  4. 按照[快速开始](#快速开始)中的命令安装依赖

### 使用方法

  1. **准备阶段**
     - 创建 `input` 文件夹（如果不存在）
     - 将需要处理的PPT文件复制到 `input` 文件夹

  2. **转换阶段**
     - 运行PPT转换程序：
       ```bash
       python ppt_to_png.py
       ```
     - 等待转换完成（可在 `ppt_conversion.log` 中查看进度）

  3. **提取阶段**
     - 运行OCR处理程序：
       ```bash
       python ocr_process.py
       ```
     - 等待处理完成（可在 `ocr_process.log` 中查看进度）

  4. **查看结果**
     - `output` 文件夹：所有PPT页面的图片
     - `FinalOutput` 文件夹：包含关键词的图片，文件命名格式为：
       ```
       关键词-相对路径-页码.png
       ```
       例如：
       - `单选题-tech/techReport-21.png` 表示在 input/tech/techReport.pptx 的第21页找到了"单选题"
       - `判断题-exam/final/test-15.png` 表示在 input/exam/final/test.pptx 的第15页找到了"判断题"
     - 每个PPT文件夹下的 `ocr_results.json`：详细的识别结果

### 自定义关键词

  默认配置查找试题相关的关键词：
  ```python
  TARGET_KEYWORDS = ['单选题', '判断题', '填空题', '多选题', 
                    '简答题', '论述题', '计算题', '分析题', 
                    '应用题', '综合题']
  ```

  您可以根据需要修改。编辑 `ocr_process.py` 文件中的 `TARGET_KEYWORDS` 列表，例如：
  ```python
  # 搜索其他关键词
  TARGET_KEYWORDS = ['重要', '总结', '结论']  # 替换为您需要的关键词
  ```

## 项目结构

  ```
  project_root/
  ├── input/                # 输入文件夹
  ├── output/               # 所有转换后的图片
  ├── FinalOutput/          # 包含关键词的图片
  ├── ppt_to_png.py        # PPT转换程序
  ├── ocr_process.py       # OCR处理程序
  ├── ppt_conversion.log   # 转换日志
  └── ocr_process.log      # OCR处理日志
  ```

## 常见问题

  1. **PPT转换失败**
     - 确保已关闭所有打开的PPT文件
     - 检查PowerPoint是否正确安装
     - 查看 `ppt_conversion.log` 获取详细错误信息

  2. **OCR识别问题**
     - 确保依赖包安装正确
     - 检查图片质量
     - 查看 `ocr_process.log` 获取详细错误信息

  3. **性能问题**
     - 大型PPT文件处理可能需要较长时间
     - 确保系统有足够的内存和存储空间
