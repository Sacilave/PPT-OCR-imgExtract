from paddleocr import PaddleOCR
import os
import json
import logging
import shutil

# 定义要搜索的关键词列表
TARGET_KEYWORDS = ['单选题', '判断题', '填空题', '多选题', 
                   '简答题', '论述题', '计算题', '分析题', 
                   '应用题', '综合题']  

def process_images_with_ocr(output_dir):
    """对输出目录中的所有PNG图片进行OCR处理，并复制包含指定关键词的图片"""
    try:
        # 初始化OCR模型
        ocr = PaddleOCR(use_angle_cls=True, lang="ch")
        
        # 创建FinalOutput文件夹
        final_output_dir = os.path.join(os.path.dirname(output_dir), 'FinalOutput')
        os.makedirs(final_output_dir, exist_ok=True)
        
        # 获取input文件夹的绝对路径
        input_dir = os.path.join(os.path.dirname(output_dir), 'input')
        
        # 遍历output目录下的所有文件夹
        for folder_name in os.listdir(output_dir):
            folder_path = os.path.join(output_dir, folder_name)
            if not os.path.isdir(folder_path):
                continue
                
            # 存储OCR结果的字典
            ocr_results = {}
            
            # 获取相对路径
            relative_path = folder_name  # 这是相对于input文件夹的路径
            
            # 处理文件夹中的所有PNG图片
            for file_name in os.listdir(folder_path):
                if not file_name.endswith('.png'):
                    continue
                    
                image_path = os.path.join(folder_path, file_name)
                logging.info(f"Processing OCR for: {image_path}")
                
                # 获取页码
                page_num = file_name.replace('slide_', '').replace('.png', '')
                
                # 执行OCR
                result = ocr.ocr(image_path, cls=True)
                
                # 提取文本结果
                texts = []
                found_keywords = set()
                
                if result:
                    for line in result[0]:
                        text = line[1][0]  # 获取识别的文本
                        confidence = line[1][1]  # 获取置信度
                        texts.append({
                            'text': text,
                            'confidence': float(confidence)
                        })
                        # 检查是否包含任何目标关键词
                        for keyword in TARGET_KEYWORDS:
                            if keyword in text:
                                found_keywords.add(keyword)
                
                # 如果包含任何关键词，复制图片到FinalOutput
                if found_keywords:
                    for keyword in found_keywords:
                        # 构建新的文件名：关键字-相对路径-页码
                        new_filename = f"{keyword}-{relative_path}-{page_num}.png"
                        # 复制图片
                        dest_path = os.path.join(final_output_dir, new_filename)
                        shutil.copy2(image_path, dest_path)
                        logging.info(f"Found keyword '{keyword}', copied to: {dest_path}")
                
                # 保存OCR结果
                ocr_results[file_name] = {
                    'texts': texts,
                    'found_keywords': list(found_keywords),
                    'contains_target': len(found_keywords) > 0
                }
            
            # 将结果保存为JSON文件
            output_json = os.path.join(folder_path, 'ocr_results.json')
            with open(output_json, 'w', encoding='utf-8') as f:
                json.dump(ocr_results, f, ensure_ascii=False, indent=2)
            
            logging.info(f"OCR results saved to: {output_json}")
            
    except Exception as e:
        logging.error(f"OCR处理出错: {str(e)}")

def main():
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('ocr_process.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    # 获取output目录路径
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    
    if not os.path.exists(output_dir):
        logging.error("Output directory not found")
        return
        
    # 处理图片
    process_images_with_ocr(output_dir)
    
if __name__ == '__main__':
    main() 