import os
from pathlib import Path
import sys
import subprocess
import warnings
from PIL import Image
import cv2
import numpy as np
import time
import psutil
import gc
from contextlib import contextmanager
import weakref
import logging
import logging.handlers
from datetime import datetime
import traceback
import json
from typing import Optional, Dict, Any, List, Tuple, Union
from functools import lru_cache
import concurrent.futures
from queue import Queue
import threading
from tqdm import tqdm
Image.MAX_IMAGE_PIXELS = None  # 禁用图片大小限制警告
import platform
import mmap
import tempfile
from collections import OrderedDict
import io

# 添加用户目录的 site-packages 到 Python 路径
user_site_packages = os.path.expanduser("~/.local/lib/python3/site-packages")
if os.path.exists(user_site_packages) and user_site_packages not in sys.path:
    sys.path.append(user_site_packages)

# Windows 用户目录
windows_site_packages = os.path.expanduser("~/AppData/Roaming/Python/Python312/site-packages")
if os.path.exists(windows_site_packages) and windows_site_packages not in sys.path:
    sys.path.append(windows_site_packages)

def install_package(package_name):
    """安装Python包"""
    print(f"正在安装 {package_name}...")
    try:
        # 先尝试使用清华源安装
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install",
                "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
                package_name
            ])
            print(f"✓ {package_name} 安装成功")
            return True
        except:
            # 如果清华源失败,使用默认源
            print(f"使用清华源安装失败,尝试默认源...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"✓ {package_name} 安装成功")
            return True
    except subprocess.CalledProcessError as e:
        print(f"× {package_name} 安装失败: {str(e)}")
        return False
    except Exception as e:
        print(f"× 安装过程出错: {str(e)}")
        return False

# 定义所需的依赖
dependencies = {
    "tqdm": "tqdm",
    "python-docx": "docx",
    "PyPDF2": "PyPDF2",
    "python-pptx": "pptx",
    "openpyxl": "openpyxl",
    "pywin32": "win32com",
    "Pillow": "PIL",
    "paddlepaddle": "paddle",  # 添加PaddlePaddle
    "paddleocr": "paddleocr",  # 添加PaddleOCR
    "PyMuPDF": "fitz",
    "pdf2image": "pdf2image"
}

# 检查并安装依赖
print("检查依赖...")
for package, module in dependencies.items():
    try:
        __import__(module)
        print(f"✓ {package} 已安装")
    except ImportError:
        print(f"× {package} 未安装,开始安装...")
        if not install_package(package):
            print(f"! 请手动安装 {package}")
            sys.exit(1)

# 现在可以安全地导入所需的模块
try:
    print("正在导入基础模块...")
    from tqdm import tqdm
    from docx import Document
    from PyPDF2 import PdfReader
    from pptx import Presentation
    from openpyxl import load_workbook
    import win32com.client
    import re
    import io
    from PIL import Image
    
    print("正在导入PaddleOCR...")
    from paddleocr import PaddleOCR  # 导入PaddleOCR
    print("PaddleOCR导入成功")
    
    import fitz
    from pdf2image import convert_from_path
    print("所有模块导入成功")
except ImportError as e:
    print(f"导入模块时出错: {str(e)}")
    print("请确保所有依赖都已正确安装")
    sys.exit(1)
except Exception as e:
    print(f"初始化过程中出现未知错误: {str(e)}")
    print("错误类型:", type(e).__name__)
    import traceback
    print("错误堆栈:")
    traceback.print_exc()
    sys.exit(1)

# 过滤PDF处理的警告
warnings.filterwarnings('ignore', category=UserWarning, module='PyPDF2')

class CustomError(Exception):
    """自定义错误基类"""
    def __init__(self, message: str, error_code: str, details: Optional[Dict[str, Any]] = None):
        self.message = message
        self.error_code = error_code
        self.details = details or {}
        super().__init__(self.message)

class OCRError(CustomError):
    """OCR相关错误"""
    pass

class ImageProcessError(CustomError):
    """图像处理相关错误"""
    pass

class FileProcessError(CustomError):
    """文件处理相关错误"""
    pass

class MemoryError(CustomError):
    """内存相关错误"""
    pass

class Logger:
    """日志管理器"""
    def __init__(self, log_dir: str = "logs"):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # 禁用paddle和ppocr的日志
        logging.getLogger("paddle").setLevel(logging.ERROR)
        logging.getLogger("ppocr").setLevel(logging.ERROR)
        
        # 禁用root logger
        logging.getLogger().handlers = []
        
        # 创建日志记录器
        self.logger = logging.getLogger("DocumentCompressor")
        self.logger.setLevel(logging.DEBUG)
        
        # 清除已存在的处理器
        self.logger.handlers = []
        
        # 创建统一的格式化器
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 文件处理器(按日期轮转)
        log_file = self.log_dir / f"compressor_{datetime.now().strftime('%Y%m%d')}.log"
        file_handler = logging.handlers.TimedRotatingFileHandler(
            log_file,
            when='midnight',
            interval=1,
            backupCount=30,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        
        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        
        # 添加处理器
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        # 设置不传播到父logger
        self.logger.propagate = False
        
        # 错误日志文件
        self.error_log = self.log_dir / "error_log.json"
        self.error_records = self.load_error_records()
        
        # 性能日志文件
        self.perf_log = self.log_dir / "performance_log.json"
        self.perf_records = self.load_perf_records()

    def load_error_records(self) -> List[Dict[str, Any]]:
        """加载错误记录"""
        if self.error_log.exists():
            try:
                with open(self.error_log, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return []
        return []

    def load_perf_records(self) -> List[Dict[str, Any]]:
        """加载性能记录"""
        if self.perf_log.exists():
            try:
                with open(self.perf_log, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return []
        return []

    def save_error_records(self):
        """保存错误记录"""
        with open(self.error_log, 'w', encoding='utf-8') as f:
            json.dump(self.error_records, f, ensure_ascii=False, indent=2)

    def save_perf_records(self):
        """保存性能记录"""
        with open(self.perf_log, 'w', encoding='utf-8') as f:
            json.dump(self.perf_records, f, ensure_ascii=False, indent=2)

    def log_error(self, error: Exception, context: Dict[str, Any] = None):
        """记录错误"""
        error_record = {
            'timestamp': datetime.now().isoformat(),
            'error_type': type(error).__name__,
            'error_message': str(error),
            'traceback': traceback.format_exc(),
            'context': context or {}
        }
        
        if isinstance(error, CustomError):
            error_record['error_code'] = error.error_code
            error_record['details'] = error.details
        
        self.error_records.append(error_record)
        self.save_error_records()
        
        # 同时写入日志文件
        self.logger.error(f"错误: {error_record['error_message']}")
        self.logger.debug(f"错误详情: {error_record['traceback']}")

    def log_performance(self, operation: str, start_time: float, end_time: float, 
                       memory_usage: float, context: Dict[str, Any] = None):
        """记录性能数据"""
        perf_record = {
            'timestamp': datetime.now().isoformat(),
            'operation': operation,
            'duration': end_time - start_time,
            'memory_usage': memory_usage,
            'context': context or {}
        }
        
        self.perf_records.append(perf_record)
        self.save_perf_records()
        
        # 同时写入日志文件
        self.logger.info(f"性能记录: {operation} - 耗时: {perf_record['duration']:.2f}秒, "
                        f"内存使用: {perf_record['memory_usage']:.2f}MB")

    def debug(self, message: str):
        """记录调试信息"""
        self.logger.debug(message)

    def info(self, message: str):
        """记录一般信息"""
        self.logger.info(message)

    def warning(self, message: str):
        """记录警告信息"""
        self.logger.warning(message)

    def error(self, message: str):
        """记录错误信息"""
        self.logger.error(message)

    def critical(self, message: str):
        """记录严重错误信息"""
        self.logger.critical(message)

class ErrorHandler:
    """错误处理器"""
    def __init__(self, logger: Logger):
        self.logger = logger
        self.error_count = 0
        self.max_retries = 3
        self.retry_delay = 1

    def handle_error(self, error: Exception, context: Dict[str, Any] = None) -> bool:
        """处理错误"""
        self.error_count += 1
        
        # 记录错误
        self.logger.log_error(error, context)
        
        # 根据错误类型处理
        if isinstance(error, OCRError):
            return self.handle_ocr_error(error)
        elif isinstance(error, ImageProcessError):
            return self.handle_image_error(error)
        elif isinstance(error, FileProcessError):
            return self.handle_file_error(error)
        elif isinstance(error, MemoryError):
            return self.handle_memory_error(error)
        else:
            return self.handle_unknown_error(error)

    def handle_ocr_error(self, error: OCRError) -> bool:
        """处理OCR错误"""
        self.logger.error(f"OCR错误: {error.message}")
        if error.error_code == "OCR_INIT_FAILED":
            return False  # 初始化失败，无法继续
        elif error.error_code == "OCR_RECOGNITION_FAILED":
            return True  # 识别失败，可以重试
        return False

    def handle_image_error(self, error: ImageProcessError) -> bool:
        """处理图像处理错误"""
        self.logger.error(f"图像处理错误: {error.message}")
        if error.error_code == "IMAGE_LOAD_FAILED":
            return False  # 图像加载失败，无法继续
        elif error.error_code == "IMAGE_PROCESS_FAILED":
            return True  # 处理失败，可以重试
        return False

    def handle_file_error(self, error: FileProcessError) -> bool:
        """处理文件错误"""
        self.logger.error(f"文件处理错误: {error.message}")
        if error.error_code == "FILE_NOT_FOUND":
            return False  # 文件不存在，无法继续
        elif error.error_code == "FILE_ACCESS_DENIED":
            return False  # 文件访问被拒绝，无法继续
        elif error.error_code == "FILE_CORRUPTED":
            return False  # 文件损坏，无法继续
        return True

    def handle_memory_error(self, error: MemoryError) -> bool:
        """处理内存错误"""
        self.logger.error(f"内存错误: {error.message}")
        if error.error_code == "MEMORY_LIMIT_EXCEEDED":
            # 等待内存释放
            time.sleep(self.retry_delay)
            return True
        return False

    def handle_unknown_error(self, error: Exception) -> bool:
        """处理未知错误"""
        self.logger.error(f"未知错误: {str(error)}")
        return False

    def should_retry(self) -> bool:
        """判断是否应该重试"""
        return self.error_count < self.max_retries

    def reset_error_count(self):
        """重置错误计数"""
        self.error_count = 0

class MemoryManager:
    """内存管理器"""
    def __init__(self, max_memory_mb: int = 4096):
        self.max_memory = max_memory_mb * 1024 * 1024  # 转换为字节
        self.current_memory = 0
        self.temp_files = set()
        self.lock = threading.Lock()
        self.monitoring = False
        self.memory_history = []
        self.last_gc_time = time.time()
        self.gc_interval = 60  # 垃圾回收间隔(秒)
        
    def monitor_memory(self, operation: str):
        """监控内存使用"""
        if self.monitoring:
            return
            
        self.monitoring = True
        try:
            start_memory = psutil.Process().memory_info().rss
            self.memory_history.append((time.time(), start_memory))
            
            yield
            
            end_memory = psutil.Process().memory_info().rss
            self.memory_history.append((time.time(), end_memory))
            
            # 计算内存使用峰值
            peak_memory = max(memory for _, memory in self.memory_history)
            self.current_memory = peak_memory
            
            # 如果内存使用超过阈值,触发垃圾回收
            if peak_memory > self.max_memory * 0.8:  # 80%阈值
                self.force_garbage_collection()
                
        finally:
            self.monitoring = False
            
    def force_garbage_collection(self):
        """强制垃圾回收"""
        current_time = time.time()
        if current_time - self.last_gc_time < self.gc_interval:
            return
            
        with self.lock:
            # 清理临时文件
            for temp_file in self.temp_files:
                try:
                    temp_file.unlink()
                except:
                    pass
            self.temp_files.clear()
            
            # 清理内存映射文件
            for mmf in self.memory_mapped_files.values():
                try:
                    mmf.close()
                except:
                    pass
            self.memory_mapped_files.clear()
            
            # 清理图片缓存
            self.image_cache.clear()
            
            # 强制垃圾回收
            gc.collect()
            
            # 更新最后垃圾回收时间
            self.last_gc_time = current_time
            
    def add_temp_file(self, file_path: Path):
        """添加临时文件"""
        with self.lock:
            self.temp_files.add(file_path)
            
    def remove_temp_file(self, file_path: Path):
        """移除临时文件"""
        with self.lock:
            self.temp_files.discard(file_path)
            
    def get_memory_usage(self) -> int:
        """获取当前内存使用量"""
        return self.current_memory
        
    def get_memory_history(self) -> List[Tuple[float, int]]:
        """获取内存使用历史"""
        return self.memory_history.copy()

class ResourceManager:
    """资源管理器"""
    def __init__(self):
        self.resources = weakref.WeakValueDictionary()
        self.lock = threading.Lock()
        
    def register(self, resource_id: str, resource: Any):
        """注册资源"""
        with self.lock:
            self.resources[resource_id] = resource
        
    def unregister(self, resource_id: str):
        """注销资源"""
        with self.lock:
            self.resources.pop(resource_id, None)
        
    def get_resource(self, resource_id: str) -> Optional[Any]:
        """获取资源"""
        return self.resources.get(resource_id)
        
    def cleanup(self):
        """清理所有资源"""
        with self.lock:
            self.resources.clear()

class ImageBatchProcessor:
    """图像批处理器"""
    def __init__(self, batch_size=5):
        self.batch_size = batch_size
        self.current_batch = []
        self.memory_manager = MemoryManager()

    def add_image(self, image):
        """添加图像到批处理"""
        self.current_batch.append(image)
        if len(self.current_batch) >= self.batch_size:
            self.process_batch()

    def process_batch(self):
        """处理当前批次"""
        if not self.current_batch:
            return

        with self.memory_manager.monitor_memory("批处理"):
            # 处理当前批次
            for image in self.current_batch:
                try:
                    # 处理图像
                    yield image
                finally:
                    # 释放图像内存
                    if hasattr(image, 'close'):
                        image.close()
                    del image

            # 清理内存
            self.memory_manager.force_garbage_collection()
            
            # 如果内存使用过高，等待一段时间
            if self.memory_manager.is_memory_critical():
                print("\n警告: 内存使用过高，等待释放...")
                time.sleep(2)
                self.memory_manager.force_garbage_collection()

        # 清空当前批次
        self.current_batch = []

    def finish(self):
        """完成所有处理"""
        if self.current_batch:
            self.process_batch()
        
        # 显示内存使用统计
        print("\n内存使用统计:")
        print(f"  - 初始内存: {self.memory_manager.initial_memory:.2f} MB")
        print(f"  - 峰值内存: {self.memory_manager.peak_memory:.2f} MB")
        print(f"  - 最终内存: {self.memory_manager.get_memory_usage():.2f} MB")
        print(f"  - 内存增长: {self.memory_manager.get_memory_usage() - self.memory_manager.initial_memory:.2f} MB")

class OCRConfig:
    """OCR配置类"""
    def __init__(self):
        # 检测参数
        self.det_db_thresh = 0.1  # 降低检测阈值
        self.det_db_box_thresh = 0.1  # 降低检测框阈值
        self.det_db_unclip_ratio = 2.0  # 增加文本框扩张比例
        self.det_limit_side_len = 2000  # 降低最大边长限制
        
        # 文本处理参数
        self.min_confidence = 0.5  # 降低最小置信度
        self.y_tolerance = 20  # 增加y坐标容差
        
        # 段落处理参数
        self.angle_threshold = 3  # 角度阈值(度)
        self.paragraph_merge_threshold = 0.5  # 段落合并阈值
        self.line_spacing_threshold = 1.5  # 行间距阈值
        self.paragraph_spacing_threshold = 2.0  # 段落间距阈值

    def to_dict(self):
        """转换为字典格式"""
        return {
            'det_db_thresh': self.det_db_thresh,
            'det_db_box_thresh': self.det_db_box_thresh,
            'det_db_unclip_ratio': self.det_db_unclip_ratio,
            'det_limit_side_len': self.det_limit_side_len,
            'min_confidence': self.min_confidence,
            'y_tolerance': self.y_tolerance,
            'angle_threshold': self.angle_threshold,
            'paragraph_merge_threshold': self.paragraph_merge_threshold,
            'line_spacing_threshold': self.line_spacing_threshold,
            'paragraph_spacing_threshold': self.paragraph_spacing_threshold
        }

class ImageSegment:
    """图片分段类"""
    def __init__(self, image: Image.Image, start_y: int, end_y: int, overlap: int = 100):
        self.image = image
        self.start_y = start_y
        self.end_y = end_y
        self.overlap = overlap
        self.text = ""
        self.confidence = 0.0
        self.error = None
        self.retry_count = 0
        self.max_retries = 3

    def process(self, ocr_engine):
        """处理分段"""
        try:
            # OCR处理
            result = ocr_engine.ocr(np.array(self.image), cls=True)
            if not result or not result[0]:
                return False
            
            # 提取文本和置信度
            texts = []
            confidences = []
            for line in result[0]:
                if not line[1]:
                    continue
                text, confidence = line[1]
                texts.append(text)
                confidences.append(confidence)
            
            if not texts:
                return False
            
            self.text = '\n'.join(texts)
            self.confidence = sum(confidences) / len(confidences)
            return True
            
        except Exception as e:
            self.error = str(e)
            return False
            
    def should_retry(self):
        """判断是否需要重试"""
        return self.retry_count < self.max_retries

    def increment_retry(self):
        """增加重试次数"""
        self.retry_count += 1

class LongImageProcessor:
    """长图片处理器"""
    def __init__(self, image: Image.Image, ocr: PaddleOCR, ocr_config: OCRConfig, original_file: Path):
        self.image = image
        self.ocr = ocr
        self.ocr_config = ocr_config
        self.original_file = original_file
        self.logger = logging.getLogger(__name__)
        self.overlap = 100  # 重叠区域大小
        
        # 禁用ppocr和paddle的非错误日志
        logging.getLogger("ppocr").setLevel(logging.ERROR)
        logging.getLogger("paddle").setLevel(logging.ERROR)
        
    def process(self) -> str:
        """处理长图片"""
        # 创建临时目录
        temp_dir = Path("temp") / self.original_file.stem
        temp_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            # 分析图片布局
            text_regions = self.analyze_layout()
            
            # 创建图片分段
            segments = self.create_segments(text_regions)
            
            # 处理每个分段
            processed_texts = self.process_segments(segments, temp_dir)
            
            # 合并文本
            final_text = self.merge_text(processed_texts)
            
            return final_text
            
        except Exception as e:
            self.logger.error(f"处理长图片时出错: {str(e)}")
            return ""
            
        finally:
            # 清理临时文件
            try:
                import shutil
                shutil.rmtree(temp_dir)
            except Exception as e:
                self.logger.error(f"清理临时文件失败: {str(e)}")
            
    def analyze_layout(self) -> List[Tuple[int, int, int, int]]:
        """分析图片布局,返回文本区域列表"""
        try:
            # 将图片转换为OpenCV格式
            image = cv2.cvtColor(np.array(self.image), cv2.COLOR_RGB2BGR)
            
            # 检查图片尺寸
            height, width = image.shape[:2]
            if height > 30000 or width > 30000:
                # 计算缩放比例
                scale = min(30000 / height, 30000 / width)
                new_width = int(width * scale)
                new_height = int(height * scale)
                # 缩放图片
                image = cv2.resize(image, (new_width, new_height))
            
            # 使用PaddleOCR检测文本区域
            result = self.ocr.ocr(image, cls=True)
            if not result or not result[0]:
                return []
                
            # 提取文本区域坐标
            text_regions = []
            for line in result[0]:
                if not line:
                    continue
                points = line[0]
                x_coords = [p[0] for p in points]
                y_coords = [p[1] for p in points]
                x1, y1 = min(x_coords), min(y_coords)
                x2, y2 = max(x_coords), max(y_coords)
                text_regions.append((x1, y1, x2, y2))
            
            return text_regions
            
        except Exception as e:
            self.logger.error(f"分析布局时出错: {str(e)}")
            return []
            
    def create_segments(self, text_regions: List[Tuple[int, int, int, int]]) -> List[ImageSegment]:
        """根据文本区域创建图片分段"""
        try:
            segments = []
            image_width, image_height = self.image.size
        
            if not text_regions:
                # 如果没有检测到文本区域,使用固定分段高度
                segment_height = 3000  # 默认分段高度
                for y in range(0, image_height, segment_height):
                    # 添加重叠区域
                    start_y = max(0, y - self.overlap)
                    end_y = min(image_height, y + segment_height + self.overlap)
                    
                    # 创建分段
                    segment = self.image.crop((0, start_y, image_width, end_y))
                    segments.append(ImageSegment(segment, start_y, end_y, self.overlap))
            else:
                # 根据文本区域创建分段
                current_y = 0
                for region in text_regions:
                    x1, y1, x2, y2 = region
                    
                    # 如果当前区域太大,需要分割
                    if y2 - y1 > 3000:
                        while y1 < y2:
                            segment_end = min(y1 + 3000, y2)
                            # 添加重叠区域
                            segment_start = max(0, y1 - self.overlap)
                            segment_end = min(image_height, segment_end + self.overlap)
                            
                            # 创建分段
                            segment = self.image.crop((0, segment_start, image_width, segment_end))
                            segments.append(ImageSegment(segment, segment_start, segment_end, self.overlap))
                            y1 = segment_end - self.overlap
                    else:
                        # 添加重叠区域
                        segment_start = max(0, y1 - self.overlap)
                        segment_end = min(image_height, y2 + self.overlap)
                        
                        # 创建分段
                        segment = self.image.crop((0, segment_start, image_width, segment_end))
                        segments.append(ImageSegment(segment, segment_start, segment_end, self.overlap))
                    
                    current_y = y2
            
            return segments
            
        except Exception as e:
            self.logger.error(f"创建分段时出错: {str(e)}")
            return []
            
    def process_segments(self, segments: List[ImageSegment], temp_dir: Path) -> List[str]:
        """处理每个分段"""
        processed_texts = []
        temp_files = []  # 用于跟踪临时文件
        
        # 创建进度条
        with tqdm(total=len(segments), 
                 desc="处理长图片分段", 
                 leave=True, 
                 position=1,
                 unit="段",
                 bar_format="{desc} |{bar}| {n_fmt}/{total_fmt}段 "
                           "[已用时:{elapsed}剩余:{remaining}, "
                           "处理速度:{rate_fmt}]") as pbar:
            for i, segment in enumerate(segments, 1):
                temp_file = None
                try:
                    # 从临时文件池获取临时文件
                    temp_file = self.temp_file_pool.get_temp_file(prefix=f"{self.original_file.stem}_segment_{i}_")
                    temp_files.append(temp_file)
                    
                    # 保存分段图片
                    segment.image.save(temp_file)
                    
                    # 预处理分段
                    processed_image = self.preprocess_segment(segment.image)
                    if processed_image is None:
                        continue
                        
                    # OCR识别
                    result = self.ocr.ocr(processed_image, cls=True)
                    if not result or not result[0]:
                        continue
                        
                    # 提取文本
                    text = ""
                    for line in result[0]:
                        if not line:
                            continue
                        text += line[1][0] + "\n"
                        
                    if text.strip():
                        processed_texts.append(text)
                        
                except Exception as e:
                    self.logger.error(f"处理分段 {i} 时出错: {str(e)}")
                    continue
                    
                finally:
                    # 归还临时文件到池中
                    if temp_file:
                        self.temp_file_pool.return_temp_file(temp_file)
                        temp_files.remove(temp_file)
                    
                    # 更新进度条
                    pbar.update(1)
        
        return processed_texts
        
    def preprocess_segment(self, segment: Image.Image) -> Optional[np.ndarray]:
        """预处理分段图片"""
        try:
            # 转换为OpenCV格式
            image = cv2.cvtColor(np.array(segment), cv2.COLOR_RGB2BGR)
            
            # 调整大小
            height, width = image.shape[:2]
            if height > 2000:
                scale = 2000 / height
                image = cv2.resize(image, None, fx=scale, fy=scale)
                
            # 图像增强
            image = cv2.fastNlMeansDenoisingColored(image, None, 10, 10, 7, 21)
            
            return image
            
        except Exception as e:
            self.logger.error(f"预处理分段时出错: {str(e)}")
            return None
            
    def merge_text(self, texts: List[str]) -> str:
        """合并处理后的文本"""
        if not texts:
            return ""
            
        # 合并文本并去重
        merged_text = "\n".join(texts)
        lines = merged_text.split("\n")
        unique_lines = []
        seen = set()
        
        for line in lines:
            line = line.strip()
            if line and line not in seen:
                seen.add(line)
                unique_lines.append(line)
                
        return "\n".join(unique_lines)

class TempFilePool:
    """临时文件池"""
    def __init__(self, max_size: int = 10):
        self.pool = OrderedDict()  # 使用OrderedDict实现LRU
        self.max_size = max_size
        self.lock = threading.Lock()
        
    def get_temp_file(self, prefix: str = "temp") -> Path:
        """获取临时文件"""
        with self.lock:
            # 检查是否有可用的临时文件
            if self.pool:
                # 获取最旧的临时文件
                temp_file = next(iter(self.pool.values()))
                # 从池中移除
                self.pool.popitem(last=False)
                return temp_file
                
            # 创建新的临时文件
            temp_file = Path(tempfile.mktemp(prefix=prefix))
            return temp_file
            
    def return_temp_file(self, temp_file: Path):
        """归还临时文件到池中"""
        with self.lock:
            if len(self.pool) >= self.max_size:
                # 如果池已满,删除最旧的临时文件
                oldest_file = next(iter(self.pool.values()))
                try:
                    oldest_file.unlink()
                except:
                    pass
                self.pool.popitem(last=False)
            
            # 添加到池中
            self.pool[temp_file] = temp_file
        
    def cleanup(self):
        """清理所有临时文件"""
        with self.lock:
            for temp_file in self.pool.values():
                try:
                    temp_file.unlink()
                except:
                    pass
            self.pool.clear()

class MemoryMappedFile:
    """内存映射文件管理器"""
    def __init__(self, file_path: Union[str, Path], chunk_size: int = 1024*1024):
        self.file_path = Path(file_path)
        self.chunk_size = chunk_size
        self.mmap = None
        self.current_chunk = 0
        self.total_chunks = 0
        
    def __enter__(self):
        """打开内存映射文件"""
        self.file = open(self.file_path, 'rb')
        self.mmap = mmap.mmap(self.file.fileno(), 0, access=mmap.ACCESS_READ)
        self.total_chunks = (self.mmap.size() + self.chunk_size - 1) // self.chunk_size
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """关闭内存映射文件"""
        if self.mmap:
            self.mmap.close()
        self.file.close()
        
    def read_chunk(self) -> Optional[bytes]:
        """读取下一个数据块"""
        if not self.mmap or self.current_chunk >= self.total_chunks:
            return None
            
        start = self.current_chunk * self.chunk_size
        end = min(start + self.chunk_size, self.mmap.size())
        data = self.mmap[start:end]
        self.current_chunk += 1
        return data
        
    def reset(self):
        """重置读取位置"""
        self.current_chunk = 0

class ImageCache:
    """优化的图片缓存池"""
    def __init__(self, max_size: int = 100, max_memory_mb: int = 1024):
        self.cache = OrderedDict()  # 使用OrderedDict实现LRU
        self.max_size = max_size
        self.max_memory = max_memory_mb * 1024 * 1024  # 转换为字节
        self.current_memory = 0
        self.lock = threading.Lock()
        self.access_times = {}
        
    def get(self, key: str) -> Optional[np.ndarray]:
        """获取缓存的图像"""
        with self.lock:
            if key in self.cache:
                # 更新访问时间
                self.access_times[key] = time.time()
                # 移动到最新位置
                self.cache.move_to_end(key)
                return self.cache[key]
            return None
            
    def put(self, key: str, image: np.ndarray):
        """添加图像到缓存"""
        with self.lock:
            # 计算图像大小
            image_size = image.nbytes
            
            # 如果缓存已满,删除最旧的图像
            while (len(self.cache) >= self.max_size or 
                   self.current_memory + image_size > self.max_memory):
                if not self.cache:
                    break
                    
                # 删除最旧的图像
                oldest_key = next(iter(self.cache))
                oldest_image = self.cache[oldest_key]
                self.current_memory -= oldest_image.nbytes
                del self.cache[oldest_key]
                del self.access_times[oldest_key]
            
            # 添加新图像
            self.cache[key] = image
            self.access_times[key] = time.time()
            self.current_memory += image_size
            
    def clear(self):
        """清空缓存"""
        with self.lock:
            self.cache.clear()
            self.access_times.clear()
            self.current_memory = 0
            
    def get_memory_usage(self) -> int:
        """获取当前内存使用量"""
        return self.current_memory

class ImageProcessor:
    """图像处理器"""
    def __init__(self, num_workers: int = 4):
        self.num_workers = num_workers
        self.thread_pool = concurrent.futures.ThreadPoolExecutor(max_workers=num_workers)
        self.process_queue = Queue()
        self.result_queue = Queue()
        self.cache = ImageCache()
        self.lock = threading.Lock()

    def preprocess_image(self, image_input):
        """图像预处理"""
        try:
            # 确保图像是PIL Image对象
            if isinstance(image_input, str):
                image = Image.open(image_input)
            elif isinstance(image_input, Image.Image):
                image = image_input
            else:
                image = Image.fromarray(image_input)
            
            # 转换为RGB模式
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # 调整图像大小
            max_dimension = 2000
            ratio = min(max_dimension / image.width, max_dimension / image.height)
            if ratio < 1:
                new_size = (int(image.width * ratio), int(image.height * ratio))
                image = image.resize(new_size, Image.Resampling.LANCZOS)
            
            # 转换为numpy数组
            img_array = np.array(image)
            
            # 转换为BGR格式(OpenCV使用BGR)
            img_bgr = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
            
            # 图像增强
            img_bgr = cv2.fastNlMeansDenoisingColored(img_bgr, None, 10, 10, 7, 21)
            
            # 验证预处理结果
            if img_bgr is None or img_bgr.size == 0:
                print("图像预处理失败: 结果为空")
                return None
                
            return img_bgr
            
        except Exception as e:
            print(f"图像预处理失败: {str(e)}")
            return None

    def process_batch(self, images: List[Tuple[str, np.ndarray]]) -> List[Tuple[str, np.ndarray]]:
        """批量处理图像"""
        futures = []
        for image_key, image in images:
            # 添加到缓存
            self.cache.put(image_key, image)
            # 提交处理任务
            future = self.thread_pool.submit(self.preprocess_image, image_key)
            futures.append((image_key, future))
        
        # 收集结果
        results = []
        for image_key, future in futures:
            try:
                processed = future.result()
                if processed is not None:
                    results.append((image_key, processed))
            except Exception as e:
                print(f"处理图像 {image_key} 时出错: {str(e)}")
        
        return results

    def shutdown(self):
        """关闭线程池"""
        self.thread_pool.shutdown(wait=True)
        self.cache.clear()

class ParallelProcessor:
    """文件处理器"""
    def __init__(self, num_workers: int = 4, ocr_engine=None, ocr_config=None):
        self.num_workers = num_workers
        self.image_processor = ImageProcessor(num_workers=num_workers)
        self.batch_size = 5
        self.processing_queue = Queue()
        self.result_queue = Queue()
        self.logger = Logger()
        self.ocr = ocr_engine
        self.ocr_config = ocr_config

    def ocr_image(self, image):
        """OCR识别图片"""
        try:
            # 确保图片是numpy数组格式
            if isinstance(image, Image.Image):
                image = np.array(image)
            
            # 调用PaddleOCR进行识别
            result = self.ocr.ocr(image, cls=True)
            if not result or not result[0]:
                return ""
            
            # 提取文本
            texts = []
            for line in result[0]:
                if not line or not line[1]:
                    continue
                text, confidence = line[1]
                # 根据置信度过滤
                if confidence >= self.ocr_config.min_confidence:
                    texts.append(text)
            
            # 合并文本
            return "\n".join(texts)
                    
        except Exception as e:
            self.logger.error(f"OCR识别失败: {str(e)}")
            self.logger.error(f"错误详情: {traceback.format_exc()}")
            return ""

    def process_files(self, files: List[Path]) -> List[Tuple[Path, bool]]:
        """串行处理文件"""
        results = []
        
        # 创建总进度条
        with tqdm(total=len(files), desc="处理文件", leave=True,
                 mininterval=0.5, maxinterval=1.0) as pbar:
            # 串行处理每个文件
            for file_path in files:
                try:
                    success = self.process_single_file(file_path)
                    results.append((file_path, success))
                    if success:
                        self.logger.info(f"✓ {file_path.name} 处理成功")
                    else:
                        self.logger.error(f"× {file_path.name} 处理失败")
                except Exception as e:
                    self.logger.error(f"处理文件 {file_path.name} 时出错: {str(e)}")
                    results.append((file_path, False))
                
                # 更新进度条
                pbar.update(1)
        
        return results

    def process_single_file(self, file_path: Path) -> bool:
        """处理单个文件"""
        try:
            self.logger.info(f"开始处理文件: {file_path.name}")
            
            # 检查文件是否存在
            if not file_path.exists():
                self.logger.error(f"文件不存在: {file_path}")
                return False
                
            # 检查文件是否可访问
            try:
                with open(file_path, 'rb') as f:
                    pass
            except PermissionError:
                self.logger.error(f"文件被其他程序占用: {file_path}")
                return False
                
            # 获取文件大小
            file_size = file_path.stat().st_size
            self.logger.info(f"文件大小: {file_size / 1024:.2f} KB")
            
            # 根据文件类型选择处理方法
            if file_path.suffix.lower() in {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif'}:
                # 处理图片文件
                try:
                    # 读取图片
                    image = Image.open(file_path)
                    
                    # 获取原始图片信息
                    original_size = image.size
                    self.logger.debug(f"原始图片尺寸: {original_size}")
                    
                    # 如果是长截图,使用长图片处理器
                    if original_size[1] > original_size[0] * 3:
                        self.logger.info("检测到长截图,使用长图片处理器...")
                        
                        # 创建长图片处理器
                        processor = LongImageProcessor(image, self.ocr, self.ocr_config, file_path)
                        
                        # 处理图片
                        text = processor.process()
                        
                        if not text.strip():
                            self.logger.error(f"OCR未能识别出文本: {file_path}")
                            return False
                    else:
                        # 预处理图片
                        processed_image = self.image_processor.preprocess_image(image)
                        if processed_image is None:
                            self.logger.error(f"图片预处理失败: {file_path}")
                            return False
                        
                        # 将处理后的图片转换回PIL格式
                        processed_image = Image.fromarray(cv2.cvtColor(processed_image, cv2.COLOR_BGR2RGB))
                        
                        # OCR处理
                        text = self.ocr_image(processed_image)
                        if not text.strip():
                            self.logger.error(f"OCR未能识别出文本: {file_path}")
                            return False
                        
                        # 保存结果
                        output_file = Path("output") / f"{file_path.stem}_compressed.txt"
                        with open(output_file, 'w', encoding='utf-8') as f:
                            f.write(text)
                        
                        self.logger.info(f"✓ {file_path.name} 处理成功")
                        self.logger.info(f"  - 原始大小: {file_size / 1024:.2f} KB")
                        self.logger.info(f"  - 压缩大小: {output_file.stat().st_size / 1024:.2f} KB")
                        self.logger.info(f"  - 压缩比: {output_file.stat().st_size / file_size:.2%}")
                        
                        return True
                        
                except Exception as e:
                    self.logger.error(f"处理图片文件时出错: {str(e)}")
                    self.logger.error(f"错误详情: {traceback.format_exc()}")
                    return False
                
            elif file_path.suffix.lower() == '.pdf':
                # 处理PDF文件
                try:
                    return self.process_pdf(file_path)
                except Exception as e:
                    self.logger.error(f"处理PDF文件失败: {str(e)}")
                    self.logger.error(f"错误详情: {traceback.format_exc()}")
                    return False
                
            else:
                self.logger.error(f"不支持的文件类型: {file_path.suffix}")
                return False
                
        except Exception as e:
            self.logger.error(f"处理文件时出错: {str(e)}")
            self.logger.error(f"错误详情: {traceback.format_exc()}")
            return False

    def process_pdf(self, file_path: Path) -> bool:
        """处理PDF文件"""
        try:
            self.logger.info(f"开始处理PDF文件: {file_path.name}")
            
            # 禁用ppocr的非错误日志
            logging.getLogger("ppocr").setLevel(logging.ERROR)
            # 禁用paddle的非错误日志
            logging.getLogger("paddle").setLevel(logging.ERROR)
            
            # 打开PDF文件
            pdf_document = fitz.open(file_path)
            
            # 创建临时目录用于保存提取的图片
            temp_dir = Path("temp") / file_path.stem
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            all_texts = []
            
            # 获取文件大小(MB)
            file_size_mb = file_path.stat().st_size / (1024 * 1024)
            
            # 创建PDF页面进度条
            with tqdm(total=len(pdf_document), 
                     desc=f"处理PDF: {file_path.name}", 
                     position=0, 
                     leave=True,
                     unit="页",
                     bar_format="{desc} |{bar}| {n_fmt}/{total_fmt}页 "
                               "[已用时:{elapsed}剩余:{remaining}, "
                               "处理速度:{rate_fmt}]") as pdf_pbar:
                # 处理每一页
                for page_num in range(len(pdf_document)):
                    try:
                        # 获取页面
                        page = pdf_document[page_num]
                        
                        # 提取图片
                        image_list = page.get_images()
                        
                        # 如果页面包含图片,创建图片处理进度条
                        if image_list:
                            with tqdm(total=len(image_list), 
                                    desc=f"第{page_num + 1}页图片", 
                                    position=1, 
                                    leave=False,
                                    unit="张",
                                    bar_format="{desc} |{bar}| {n_fmt}/{total_fmt}张 "
                                              "[已用时:{elapsed}剩余:{remaining}, "
                                              "处理速度:{rate_fmt}]") as img_pbar:
                                for img_index, img_info in enumerate(image_list):
                                    try:
                                        # 获取图片
                                        base_image = pdf_document.extract_image(img_info[0])
                                        image_bytes = base_image["image"]
                                        
                                        # 转换为PIL图片
                                        image = Image.open(io.BytesIO(image_bytes))
                                        
                                        # 检查是否为长图片
                                        if image.height > image.width * 3:
                                            self.logger.info(f"检测到长图片: 第{page_num + 1}页, 图片{img_index + 1}")
                                            
                                            # 保存图片到临时文件
                                            temp_file = temp_dir / f"page_{page_num + 1}_img_{img_index + 1}.png"
                                            image.save(temp_file)
                                            
                                            # 使用长图片处理器处理
                                            processor = LongImageProcessor(image, self.ocr, self.ocr_config, temp_file)
                                            text = processor.process()
                                            
                                            if text.strip():
                                                all_texts.append(text)
                                        else:
                                            # 使用普通OCR处理
                                            text = self.ocr_image(image)
                                            if text.strip():
                                                all_texts.append(text)
                                                
                                        # 更新图片进度条
                                        img_pbar.update(1)
                                            
                                    except Exception as e:
                                        self.logger.error(f"处理图片失败: 第{page_num + 1}页, 图片{img_index + 1}")
                                        self.logger.error(str(e))
                                        continue
                        
                        # 提取页面文本
                        page_text = page.get_text()
                        if page_text.strip():
                            all_texts.append(page_text)
                            
                    except Exception as e:
                        self.logger.error(f"处理PDF页面失败: 第{page_num + 1}页")
                        self.logger.error(str(e))
                        continue
                        
                    finally:
                        # 更新PDF页面进度条
                        pdf_pbar.update(1)
                        
                        # 更新进度条描述,显示处理速度
                        elapsed_time = time.time() - pdf_pbar.start_t
                        if elapsed_time > 0:
                            mb_per_sec = file_size_mb * (page_num + 1) / (len(pdf_document) * elapsed_time)
                            pdf_pbar.set_postfix({"处理速度": f"{mb_per_sec:.2f}MB/s"})
            
            # 合并所有文本
            final_text = "\n\n".join(all_texts)
            
            # 保存结果
            output_file = Path("output") / f"{file_path.stem}_compressed.txt"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(final_text)
            
            self.logger.info(f"✓ {file_path.name} 处理成功")
            self.logger.info(f"  - 原始大小: {file_path.stat().st_size / 1024:.2f} KB")
            self.logger.info(f"  - 压缩大小: {output_file.stat().st_size / 1024:.2f} KB")
            self.logger.info(f"  - 压缩比: {output_file.stat().st_size / file_path.stat().st_size:.2%}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"处理PDF文件失败: {str(e)}")
            self.logger.error(f"错误详情: {traceback.format_exc()}")
            return False
            
        finally:
            # 清理临时文件
            if 'temp_dir' in locals():
                try:
                    import shutil
                    shutil.rmtree(temp_dir)
                except:
                    pass

    def shutdown(self):
        """关闭处理器"""
        self.image_processor.shutdown()

class OCRManager:
    """OCR引擎管理器(单例模式)"""
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(OCRManager, cls).__new__(cls)
        return cls._instance
        
    def __init__(self):
        if not self._initialized:
            self._initialized = True
            self.logger = logging.getLogger(__name__)
            self.init_ocr_engine()
            
    def init_ocr_engine(self):
        """初始化OCR引擎"""
        try:
            # 检查是否有GPU
            import torch
            use_gpu = torch.cuda.is_available()
            self.logger.info(f"GPU可用: {use_gpu}")
            
            # OCR参数配置
            self.ocr_config = {
                # 通用配置
                'use_gpu': use_gpu,
                'use_mp': True,
                'total_process_num': os.cpu_count(),
                'enable_mkldnn': True,  # 启用Intel加速
                
                # 检测模型配置
                'det_model_dir': 'models/ch_PP-OCRv4_det_infer',  # 使用v4检测模型
                'det_limit_side_len': 960,  # 限制最长边,加快处理
                'det_db_thresh': 0.3,  # 降低检测阈值,提高速度
                'det_db_box_thresh': 0.5,  # 框选阈值
                'det_db_unclip_ratio': 1.6,  # 文本框扩张比例
                
                # 识别模型配置
                'rec_model_dir': 'models/ch_PP-OCRv4_rec_infer',  # 使用v4识别模型
                'rec_batch_num': 6,  # 批量识别数量
                'rec_img_shape': "3, 48, 320",  # 限制识别图片大小
                
                # 角度分类配置
                'use_angle_cls': False,  # 关闭角度分类,提高速度
                'cls_model_dir': None,
                
                # 其他优化
                'show_log': False,  # 关闭paddleocr的日志
            }
            
            # 初始化OCR引擎
            self.ocr = PaddleOCR(**self.ocr_config)
            self.logger.info("OCR引擎初始化成功")
            
        except Exception as e:
            self.logger.error(f"OCR引擎初始化失败: {str(e)}")
            self.logger.error(f"错误详情: {traceback.format_exc()}")
            raise
            
    def get_ocr_engine(self) -> PaddleOCR:
        """获取OCR引擎实例"""
        return self.ocr
        
    def get_config(self) -> dict:
        """获取OCR配置"""
        return self.ocr_config

class OCRConfig:
    """OCR配置"""
    def __init__(self):
        self.min_confidence = 0.5  # 最小置信度
        self.max_workers = os.cpu_count()  # 最大工作进程数
        self.temp_dir = Path("temp")  # 临时目录
        self.output_dir = Path("output")  # 输出目录
        
        # 创建必要的目录
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

class DocumentCompressor:
    """文档压缩器"""
    def __init__(self):
        self.logger = Logger()
        self.config = OCRConfig()
        
        # 使用OCR管理器获取引擎实例
        ocr_manager = OCRManager()
        self.ocr = ocr_manager.get_ocr_engine()
        
        # 创建处理器
        self.processor = ParallelProcessor(
            num_workers=self.config.max_workers,
            ocr_engine=self.ocr,
            ocr_config=self.config
        )
        
        # 初始化其他属性
        self.input_dir = Path("raw")
        self.output_dir = Path("output")
        self.temp_dir = Path("temp")
        
        # 创建必要的目录
        for dir_path in [self.input_dir, self.output_dir, self.temp_dir]:
            dir_path.mkdir(parents=True, exist_ok=True)
            self.logger.debug(f"创建目录: {dir_path}")
        
        # 支持的文件类型
        self.supported_extensions = {
            '.docx', '.doc',  # Word文档
            '.pdf',          # PDF文档
            '.pptx', '.ppt', # PPT文档
            '.xlsx', '.xls', # Excel文档
            '.txt',          # 纯文本
            '.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif'  # 图片文件
        }
        
        # 初始化统计信息
        self.stats = {
            'processed': 0,
            'failed': 0,
            'total_size_before': 0,
            'total_size_after': 0,
            'unsupported_files': [],
            'processing_times': {}
        }
        
        # 初始化资源管理器
        self.resource_manager = ResourceManager()
        
        # 初始化内存管理器
        self.memory_manager = MemoryManager(max_memory_mb=4096)
        
        # 初始化批处理器
        self.batch_processor = ImageBatchProcessor()
        
        self.logger.info("文档压缩工具初始化完成")
        self.logger.info(f"输入目录: {self.input_dir.absolute()}")
        self.logger.info(f"输出目录: {self.output_dir.absolute()}")
        self.logger.info(f"支持的文件类型: {', '.join(self.supported_extensions)}")

        # 初始化并行处理器
        self.parallel_processor = ParallelProcessor(num_workers=4, ocr_engine=self.ocr, ocr_config=self.config)
        
        # 初始化图像处理器
        self.image_processor = ImageProcessor(num_workers=4)
        
        # 初始化缓存
        self.image_cache = ImageCache(max_size=100)
        
        # 初始化临时文件池
        self.temp_file_pool = TempFilePool(max_size=10)
        
        # 初始化内存映射管理器
        self.memory_mapped_files = {}

    def update_ocr_config(self, **kwargs):
        """更新OCR配置"""
        for key, value in kwargs.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
        
        # 重新初始化PaddleOCR
        try:
            self.ocr = PaddleOCR(**self.config.to_dict())
            self.logger.info("OCR配置更新成功")
        except Exception as e:
            print(f"OCR配置更新失败: {str(e)}")
            raise

    def compress_text(self, text):
        """简化的文本压缩"""
        if not text:
            return ""
            
        # 分割成行并去重
        lines = text.split('\n')
        unique_lines = []
        seen = set()
        
        for line in lines:
            line = line.strip()
            if line and line not in seen:
                unique_lines.append(line)
                seen.add(line)
        
        # 合并处理后的行
        return '\n'.join(unique_lines).strip()

    def log(self, message):
        """添加日志到缓存"""
        self.log_buffer.append(message)

    def clear_log(self):
        """清空日志缓存"""
        self.log_buffer = []

    def print_log(self):
        """打印日志缓存"""
        for message in self.log_buffer:
            print(message)

    def check_files(self):
        """检查目录中的所有文件，返回支持和不支持的文件列表"""
        supported_files = []
        unsupported_files = []
        temp_files = []
        
        for file_path in self.input_dir.glob('*.*'):
            # 过滤临时文件
            if file_path.name.startswith('~$'):
                temp_files.append(file_path)
                continue
            
            # 检查文件扩展名
            if file_path.suffix.lower() in self.supported_extensions:
                # 检查文件是否可访问
                try:
                    with open(file_path, 'rb'):
                        supported_files.append(file_path)
                except PermissionError:
                    self.log(f"警告: 文件 {file_path.name} 被其他程序占用")
                    continue
            else:
                unsupported_files.append(file_path)
                self.stats['unsupported_files'].append(file_path.name)
        
        # 显示临时文件信息
        if temp_files:
            self.log("\n发现以下临时文件,将被跳过:")
            for temp_file in temp_files:
                self.log(f"- {temp_file.name}")
        
        return supported_files, unsupported_files

    def preprocess_image(self, image_input):
        """图像预处理"""
        try:
            # 确保图像是PIL Image对象
            if isinstance(image_input, str):
                image = Image.open(image_input)
            elif isinstance(image_input, Image.Image):
                image = image_input
            else:
                image = Image.fromarray(image_input)
            
            # 转换为RGB模式
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # 调整图像大小
            max_dimension = 2000
            ratio = min(max_dimension / image.width, max_dimension / image.height)
            if ratio < 1:
                new_size = (int(image.width * ratio), int(image.height * ratio))
                image = image.resize(new_size, Image.Resampling.LANCZOS)
            
            # 转换为numpy数组
            img_array = np.array(image)
            
            # 转换为BGR格式(OpenCV使用BGR)
            img_bgr = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
            
            # 图像增强
            img_bgr = cv2.fastNlMeansDenoisingColored(img_bgr, None, 10, 10, 7, 21)
            
            # 验证预处理结果
            if img_bgr is None or img_bgr.size == 0:
                print("图像预处理失败: 结果为空")
                return None
                
            return img_bgr
            
        except Exception as e:
            print(f"图像预处理失败: {str(e)}")
            return None

    def split_image(self, image, max_height=1000):  # 降低最大高度
        """将长图片分割成多个部分"""
        try:
            if image.height <= max_height:
                return [image]
            
            parts = []
            current_y = 0
            
            while current_y < image.height:
                # 计算当前部分的高度
                part_height = min(max_height, image.height - current_y)
                
                # 增加重叠区域
                overlap = 100  # 增加到100像素
                if current_y > 0:
                    current_y -= overlap
                
                # 裁剪图片
                part = image.crop((0, current_y, image.width, current_y + part_height))
                parts.append(part)
                
                current_y += part_height
                        
            return parts
        except Exception as e:
            self.log(f"图片分割失败: {str(e)}")
            return [image]

    def process_image(self, file_path):
        """处理图片文件"""
        start_time = time.time()
        try:
            self.logger.info(f"开始处理图片: {file_path.name}")
            
            # 使用内存映射读取图片
            with MemoryMappedFile(file_path) as mmf:
                # 读取图片数据
                image_data = mmf.read_chunk()
                if not image_data:
                    raise ImageProcessError(
                        "图片读取失败",
                        "IMAGE_LOAD_FAILED",
                        {"file": str(file_path)}
                    )
                
                # 转换为PIL图片
                image = Image.open(io.BytesIO(image_data))
                
                # 获取原始图片信息
                original_size = image.size
                self.logger.debug(f"原始图片尺寸: {original_size}")
                
                # 如果是长截图,使用长图片处理器
                if original_size[1] > original_size[0] * 3:
                    self.logger.info("检测到长截图,使用长图片处理器...")
                    
                    # 使用临时文件池获取临时文件
                    temp_file = self.temp_file_pool.get_temp_file(prefix=f"{file_path.stem}_")
                    try:
                        # 保存原始图片
                        image.save(temp_file)
                        
                        # 创建长图片处理器
                        processor = LongImageProcessor(image, self.ocr, self.config, file_path)
                        
                        # 处理图片
                        with self.memory_manager.monitor_memory("长图片处理"):
                            text = processor.process()
                    finally:
                        # 归还临时文件到池中
                        self.temp_file_pool.return_temp_file(temp_file)
                
                if not text.strip():
                    raise OCRError(
                        "OCR未能识别出文本",
                        "OCR_RECOGNITION_FAILED",
                        {"file": str(file_path)}
                    )
                    
                    # 简化文本压缩
                    self.logger.info("压缩识别出的文本...")
                compressed_text = self.compress_text(text)
                
                # 保存结果
                output_file = self.output_dir / f"{file_path.stem}_compressed.txt"
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(compressed_text)
                
                # 更新统计信息
                self.stats['total_size_before'] += file_path.stat().st_size
                self.stats['total_size_after'] += output_file.stat().st_size
                self.stats['processed'] += 1
                    
                # 记录性能数据
                end_time = time.time()
                self.logger.log_performance(
                    "长图片处理",
                    start_time,
                    end_time,
                    self.memory_manager.get_memory_usage(),
                    {
                        "file": str(file_path),
                        "original_size": original_size,
                        "compressed_size": output_file.stat().st_size
                    }
                )
                
                # 显示处理结果
                self.logger.info(f"✓ {file_path.name} 处理成功")
                self.logger.info(f"  - 原始大小: {file_path.stat().st_size / 1024:.2f} KB")
                self.logger.info(f"  - 压缩大小: {output_file.stat().st_size / 1024:.2f} KB")
                self.logger.info(f"  - 压缩比: {output_file.stat().st_size / file_path.stat().st_size:.2%}")
                
                return True
                
        except Exception as e:
            self.logger.error(f"图片处理失败: {str(e)}")
            self.logger.error(f"错误详情: {traceback.format_exc()}")
            
            # 处理错误
            if not self.error_handler.handle_error(e, {"file": str(file_path)}):
                self.stats['failed'] += 1
                return False
            
            # 如果错误可以重试
            if self.error_handler.should_retry():
                self.logger.info("准备重试处理...")
                time.sleep(self.error_handler.retry_delay)
                return self.process_image(file_path)
            
            return False
            
        finally:
            # 清理资源
            self.memory_manager.force_garbage_collection()
            self.error_handler.reset_error_count()

    def process_all_files(self):
        """处理所有文件"""
        # 检查目录
        if not self.input_dir.exists():
            self.logger.info(f"创建输入目录: {self.input_dir}")
            self.input_dir.mkdir(parents=True)
            self.logger.info("请将要处理的文件放入 raw 目录中，然后重新运行程序")
            return

        if not self.output_dir.exists():
            self.output_dir.mkdir(parents=True)

        # 检查文件
        supported_files, unsupported_files = self.check_files()

        if not supported_files:
            self.logger.warning("未找到可处理的文件！")
            self.logger.info("请将文件放入 raw 目录后重试")
            return

        self.logger.info(f"找到 {len(supported_files)} 个文件待处理")
        
        try:
            # 使用并行处理器处理文件
            results = self.parallel_processor.process_files(supported_files)
            
            # 统计结果
            for file_path, success in results:
                if success:
                    self.stats['processed'] += 1
                else:
                    self.stats['failed'] += 1
            
            # 显示统计信息
            self.print_stats()
                    
        except KeyboardInterrupt:
            self.logger.warning("程序已中断")
        finally:
            # 关闭并行处理器
            self.parallel_processor.shutdown()
            
            # 清理资源
            self.memory_manager.cleanup_temp_files()
            self.image_cache.clear()
            
            # 显示最终内存使用情况
            self.logger.info("\n最终内存使用情况:")
            self.logger.info(f"  - 初始内存: {self.memory_manager.initial_memory:.2f} MB")
            self.logger.info(f"  - 峰值内存: {self.memory_manager.peak_memory:.2f} MB")
            self.logger.info(f"  - 最终内存: {self.memory_manager.get_memory_usage():.2f} MB")
            self.logger.info(f"  - 内存增长: {self.memory_manager.get_memory_usage() - self.memory_manager.initial_memory:.2f} MB")
    
    def print_stats(self):
        """打印处理统计信息"""
        print("\n处理统计:")
        print(f"✓ 成功处理: {self.stats['processed']} 个文件")
        if self.stats['failed'] > 0:
            print(f"× 处理失败: {self.stats['failed']} 个文件")
        if self.stats['unsupported_files']:
            print(f"- 不支持的文件: {len(self.stats['unsupported_files'])} 个")
        print(f"- 原始大小: {self.stats['total_size_before'] / 1024:.2f} KB")
        print(f"- 压缩大小: {self.stats['total_size_after'] / 1024:.2f} KB")
        if self.stats['total_size_before'] > 0:
            ratio = self.stats['total_size_after'] / self.stats['total_size_before']
            print(f"- 总压缩比: {ratio:.2%}")
            print(f"- 节省空间: {(self.stats['total_size_before'] - self.stats['total_size_after']) / 1024:.2f} KB")

    def __del__(self):
        """清理资源"""
        try:
            # 清理临时文件池
            self.temp_file_pool.cleanup()
            
            # 清理图片缓存
            self.image_cache.clear()
            
            # 关闭所有内存映射文件
            for mmf in self.memory_mapped_files.values():
                mmf.close()
                
            # 关闭并行处理器
            self.parallel_processor.shutdown()
        except:
            pass

def check_dependencies():
    """检查所有必要的依赖是否已安装,如果没有则自动安装"""
    missing_deps = []
    installed_deps = []
    manual_deps = []  # 需要手动安装的依赖
    
    # 依赖列表
    dependencies = {
        "python-docx": "docx",
        "PyPDF2": "PyPDF2",
        "python-pptx": "pptx",
        "openpyxl": "openpyxl",
        "PyMuPDF": "fitz",
        "pdf2image": "pdf2image",
        "Pillow": "PIL",
        "paddlepaddle": "paddle",
        "paddleocr": "paddleocr",
        "pywin32": "win32com",
        "tqdm": "tqdm",
        "scikit-image": "skimage",
        "scipy": "scipy"
    }
    
    print("\n检查Python包依赖...")
    
    # 检查每个依赖
    for package, module in dependencies.items():
        try:
            __import__(module)
            print(f"✓ {package} 已安装")
            installed_deps.append(package)
        except ImportError:
            print(f"× {package} 未安装")
            missing_deps.append(package)
    
    # 如果有缺失的依赖,尝试安装
    if missing_deps:
        print(f"\n发现 {len(missing_deps)} 个缺失的依赖,开始自动安装...")
        
        # 先升级pip
        try:
            print("\n升级pip...")
            subprocess.check_call([
                sys.executable, "-m", "pip", "install",
                "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
                "--upgrade", "pip"
            ])
        except:
            print("警告: pip升级失败,继续安装依赖...")
        
        # 安装缺失的依赖
        for package in missing_deps:
            if install_package(package):
                installed_deps.append(package)
            else:
                manual_deps.append(package)
        
        if manual_deps:
            print("\n以下依赖需要手动安装:")
            for package in manual_deps:
                print(f"pip install {package}")
            
            if "pdf2image" in manual_deps:
                print("\npdf2image还需要安装poppler:")
                print("1. 下载地址: https://github.com/oschwartz10612/poppler-windows/releases/")
                print("2. 解压到任意目录")
                print("3. 将bin目录添加到系统环境变量PATH")
    
    return manual_deps  # 返回需要手动安装的依赖

def main():
    print("文档压缩工具 v1.0")
    print("=" * 50)
    
    # 检查依赖
    manual_deps = check_dependencies()
    if manual_deps:
        print(f"\n以下依赖需要手动安装,请按照上述说明进行安装:")
        for dep in manual_deps:
            print(f"- {dep}")
        print("\n安装完成后重新运行程序")
        input("\n按回车键退出...")
        sys.exit(1)
    
    try:
        compressor = DocumentCompressor()
        print("\n支持格式:", ", ".join(sorted(compressor.supported_extensions)))
        print("\n使用说明:")
        print("1. 将需要处理的文件放在 raw 目录下")
        print("2. 运行程序进行处理")
        print("3. 处理后的文件保存在 output 目录下")
        print("4. 按 Ctrl+C 可随时中断处理")
        print("=" * 50)
        
        compressor.process_all_files()
        
    except KeyboardInterrupt:
        print("\n程序已中断")
    except Exception as e:
        print(f"\n程序出错: {str(e)}")
    finally:
        print("\n处理完成!")

if __name__ == "__main__":
    main()

print(platform.architecture())  # 检查Python位数
