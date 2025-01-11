import os
import win32com.client
import time
import logging
from pathlib import Path
import win32gui
import win32con
import pythoncom
import psutil
import shutil
import tempfile
import re
import unicodedata
import uuid

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ppt_conversion.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def sanitize_filename(filename):
    """清理文件名,移除不安全字符"""
    # 移除所有不可见字符
    filename = ''.join(char for char in filename if not unicodedata.category(char).startswith('C'))
    # 移除多余空格
    filename = ' '.join(filename.split())
    # 替换中文字符为拼音或使用UUID
    return str(uuid.uuid4()) + os.path.splitext(filename)[1]

def get_safe_temp_path(original_file):
    """获取安全的临时文件路径"""
    # 使用更简单的路径
    temp_dir = os.path.join(os.environ['TEMP'], 'PPTConversion')
    os.makedirs(temp_dir, exist_ok=True)
    # 使用简单的数字作为文件名
    temp_name = f"temp_{int(time.time())}{os.path.splitext(original_file)[1]}"
    return os.path.join(temp_dir, temp_name)

def verify_ppt_file(file_path):
    """验证PPT文件的有效性"""
    try:
        # 检查文件大小
        if os.path.getsize(file_path) < 100:
            return False
            
        # 读取文件头部进行验证
        with open(file_path, 'rb') as f:
            header = f.read(8)
            # PPT文件头部特征
            ppt_signatures = [
                b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1',  # OLE2
                b'PK\x03\x04',  # PPTX (ZIP)
            ]
            return any(header.startswith(sig) for sig in ppt_signatures)
    except Exception as e:
        logging.error(f"文件验证失败: {str(e)}")
        return False

def kill_powerpoint_processes():
    """强制结束所有PowerPoint进程"""
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'].lower() in ['powerpnt.exe', 'powerpoint.exe']:
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

def minimize_powerpoint_window():
    """最小化PowerPoint窗口"""
    def callback(hwnd, windows):
        if "PowerPoint" in win32gui.GetWindowText(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        return True
    win32gui.EnumWindows(callback, [])

def find_ppt_files(input_dir):
    """查找目录下所有ppt和pptx文件"""
    ppt_files = []
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith(('.ppt', '.pptx')):
                ppt_files.append(os.path.join(root, file))
    return ppt_files

def normalize_path(path):
    """规范化路径,处理空格等特殊字符"""
    return str(Path(path).resolve())

def verify_file_access(file_path):
    """验证文件是否可访问"""
    try:
        if not os.path.exists(file_path):
            logging.error(f"文件不存在: {file_path}")
            return False
            
        if not os.access(file_path, os.R_OK):
            logging.error(f"文件无法读取: {file_path}")
            return False
            
        if not verify_ppt_file(file_path):
            logging.error(f"不是有效的PPT文件: {file_path}")
            return False
            
        return True
    except Exception as e:
        logging.error(f"文件访问验证失败: {str(e)}")
        return False

def copy_to_temp(file_path):
    """复制文件到临时目录"""
    temp_file = None
    try:
        # 获取安全的临时文件路径
        temp_file = get_safe_temp_path(file_path)
        temp_dir = os.path.dirname(temp_file)
        
        # 确保源文件存在
        if not os.path.exists(file_path):
            raise Exception(f"源文件不存在: {file_path}")
            
        # 记录文件路径信息
        logging.info(f"原始文件: {file_path}")
        logging.info(f"临时文件: {temp_file}")
        
        # 复制文件
        shutil.copy2(file_path, temp_file)
        
        # 验证复制后的文件
        if not os.path.exists(temp_file):
            raise Exception("临时文件创建失败")
            
        if not verify_ppt_file(temp_file):
            raise Exception("文件验证失败")
            
        return temp_file
        
    except Exception as e:
        logging.error(f"复制文件到临时目录失败: {str(e)}")
        # 清理临时文件
        if temp_file and os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except:
                pass
        # 清理临时目录
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
        return None

def convert_single_slide(slide, output_path, index):
    """转换单个幻灯片,包含错误处理"""
    try:
        # 确保输出路径存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        # 导出幻灯片
        slide.Export(output_path, 'PNG')
        return True
    except Exception as e:
        logging.error(f"转换第 {index} 页时出错: {str(e)}")
        return False

def convert_ppt_to_png(ppt_path, max_retries=3):
    """将ppt文件转换为png图片"""
    temp_file = None
    temp_dir = None
    
    try:
        ppt_path = str(Path(ppt_path).resolve())
        
        # 验证文件是否可访问
        if not verify_file_access(ppt_path):
            return False
            
        # 复制到临时目录
        temp_file = copy_to_temp(ppt_path)
        if not temp_file:
            return False
            
        # 确保临时文件有完全访问权限
        os.chmod(temp_file, 0o777)
            
        temp_dir = os.path.dirname(temp_file)
        powerpoint = None
        retry_count = 0
        success = False
        
        # 初始化COM
        pythoncom.CoInitialize()
        
        try:
            while retry_count < max_retries:
                try:
                    # 确保没有遗留进程
                    kill_powerpoint_processes()
                    time.sleep(3)
                    
                    # 创建PowerPoint实例
                    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                    powerpoint.DisplayAlerts = False
                    powerpoint.Visible = True  # 设置为可见，方便调试
                    time.sleep(2)
                    
                    # 使用绝对路径打开文件
                    abs_temp_file = os.path.abspath(temp_file)
                    logging.info(f"尝试打开文件: {abs_temp_file}")
                    # 以只读方式打开
                    presentation = powerpoint.Presentations.Open(
                        abs_temp_file,
                        ReadOnly=True,
                        Untitled=False,
                        WithWindow=False
                    )
                    time.sleep(1)
                    
                    if not presentation:
                        raise Exception("文件打开失败")
                    
                    # 修改输出目录到根目录下的output文件夹
                    output_base_dir = os.path.join(os.path.dirname(os.path.dirname(ppt_path)), 'output')
                    # 使用PPT文件名作为子文件夹名
                    ppt_name = os.path.splitext(os.path.basename(ppt_path))[0]
                    output_dir = os.path.join(output_base_dir, ppt_name)
                    os.makedirs(output_dir, exist_ok=True)
                    
                    success_count = 0
                    total_slides = presentation.Slides.Count
                    
                    for i in range(total_slides):
                        try:
                            slide = presentation.Slides[i + 1]
                            output_path = str(Path(os.path.join(output_dir, f'slide_{i+1}.png')).resolve())
                            if convert_single_slide(slide, output_path, i+1):
                                success_count += 1
                                time.sleep(0.5)  # 等待导出完成
                        except Exception as e:
                            logging.error(f"处理第 {i+1}/{total_slides} 页时出错: {str(e)}")
                            continue
                    
                    # 关闭文件前等待
                    time.sleep(1)
                    presentation.Close()
                    time.sleep(1)  # 等待关闭完成
                    
                    logging.info(f"文件转换完成: 成功 {success_count}/{total_slides} 页")
                    success = True
                    break
                    
                except Exception as e:
                    retry_count += 1
                    logging.error(f"转换失败 (尝试 {retry_count}/{max_retries})")
                    logging.error(f"错误详情: {str(e)}")
                    time.sleep(2)
                    
        finally:
            if powerpoint:
                try:
                    powerpoint.Quit()
                    time.sleep(1)  # 等待PowerPoint完全退出
                except:
                    pass
            pythoncom.CoUninitialize()
            kill_powerpoint_processes()
            
            # 清理临时文件
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                    
        return success
        
    except Exception as e:
        logging.error(f"转换过程出现未预期的错误: {str(e)}")
        return False

def main():
    # 确保开始时没有PowerPoint进程残留
    kill_powerpoint_processes()
    
    # 获取当前目录下的input文件夹路径
    input_dir = normalize_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'input'))
    
    # 确保input文件夹存在
    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
        logging.info("已创建input文件夹")
        return
        
    # 查找所有PPT文件
    ppt_files = find_ppt_files(input_dir)
    
    if not ppt_files:
        logging.info("未找到PPT文件")
        return
        
    # 转换每个PPT文件
    success_count = 0
    for ppt_file in ppt_files:
        logging.info(f"开始转换: {ppt_file}")
        if convert_ppt_to_png(ppt_file):
            success_count += 1
        time.sleep(3)  # 增加等待时间,确保PowerPoint完全关闭
        
    logging.info(f"转换完成! 成功: {success_count}/{len(ppt_files)}")

if __name__ == '__main__':
    main() 