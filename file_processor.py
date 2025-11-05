#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件类型检测和文本提取模块
支持多种文件格式的智能检测和文本提取，包括PDF、DOCX、图片OCR等
"""

import os
import logging
import importlib.util
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any, Union

# 配置日志
logger = logging.getLogger(__name__)

# =============================================================================
# 依赖检查和导入
# =============================================================================

class DependencyManager:
    """依赖管理器"""
    
    def __init__(self):
        self.available_modules = {}
        self._check_dependencies()
    
    def _check_dependencies(self):
        """检查可选依赖"""
        # PDF处理 - 使用pdfminer.six + PyPDF2备用
        if importlib.util.find_spec('pdfminer'):
            self.available_modules['pdf'] = ['pdfminer.six']
        elif importlib.util.find_spec('PyPDF2'):
            self.available_modules['pdf'] = ['PyPDF2']
        else:
            self.available_modules['pdf'] = []
        
        # DOCX处理
        if importlib.util.find_spec('docx'):
            self.available_modules['docx'] = ['python-docx']
        else:
            self.available_modules['docx'] = []
        
        # 文件类型检测
        if importlib.util.find_spec('magic'):
            self.available_modules['magic'] = ['python-magic']
        else:
            self.available_modules['magic'] = []
        
        # OCR功能 - 使用科大讯飞
        # 检查requests库和xunfei_ocr模块
        has_requests = importlib.util.find_spec('requests') is not None
        has_xunfei_module = importlib.util.find_spec('xunfei_ocr') is not None
        
        if has_requests and has_xunfei_module:
            self.available_modules['ocr'] = ['xunfei']
        else:
            self.available_modules['ocr'] = []
            if not has_requests:
                logger.warning("OCR功能需要requests库，请安装: pip install requests")
            if not has_xunfei_module:
                logger.warning("未找到xunfei_ocr模块")
        
        self._log_dependencies()
    
    def _log_dependencies(self):
        """记录依赖状态"""
        for module, deps in self.available_modules.items():
            if deps:
                logger.info(f"✅ {module.upper()} 支持: {', '.join(deps)}")
            else:
                logger.warning(f"⚠️  {module.upper()} 不可用")
    
    def is_available(self, module: str) -> bool:
        """检查模块是否可用"""
        return bool(self.available_modules.get(module, []))


# 全局依赖管理器
deps = DependencyManager()


# =============================================================================
# 文件类型检测
# =============================================================================

class FileTypeDetector:
    """文件类型检测器"""
    
    # 支持的MIME类型映射
    SUPPORTED_TYPES = {
        'text': ['text/plain', 'text/html', 'text/xml', 'application/xml'],
        'pdf': ['application/pdf'],
        'docx': ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
        'doc': ['application/msword'],
        'image': ['image/jpeg', 'image/png', 'image/tiff', 'image/bmp', 'image/gif']
    }
    
    @staticmethod
    def detect_file_type(file_path: str) -> Tuple[str, str]:
        """
        检测文件类型
        
        Args:
            file_path: 文件路径
            
        Returns:
            Tuple[str, str]: (文件类型, MIME类型)
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        # 首先尝试使用python-magic
        if deps.is_available('magic'):
            try:
                import magic
                mime_type = magic.from_file(file_path, mime=True)
                return FileTypeDetector._categorize_mime_type(mime_type), mime_type
            except Exception as e:
                logger.warning(f"Magic检测失败: {e}")
        
        # 回退到基于扩展名的检测
        return FileTypeDetector._detect_by_extension(file_path)
    
    @staticmethod
    def _detect_by_extension(file_path: str) -> Tuple[str, str]:
        """基于文件扩展名检测类型"""
        ext = Path(file_path).suffix.lower()
        
        # 扩展名映射
        extension_map = {
            '.txt': ('text', 'text/plain'),
            '.md': ('text', 'text/markdown'),
            '.html': ('text', 'text/html'),
            '.xml': ('text', 'application/xml'),
            '.pdf': ('pdf', 'application/pdf'),
            '.docx': ('docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            '.doc': ('doc', 'application/msword'),
            '.jpg': ('image', 'image/jpeg'),
            '.jpeg': ('image', 'image/jpeg'),
            '.png': ('image', 'image/png'),
            '.tiff': ('image', 'image/tiff'),
            '.bmp': ('image', 'image/bmp'),
            '.gif': ('image', 'image/gif'),
        }
        
        return extension_map.get(ext, ('unknown', 'application/octet-stream'))
    
    @staticmethod
    def _categorize_mime_type(mime_type: str) -> str:
        """根据MIME类型分类"""
        for category, mime_list in FileTypeDetector.SUPPORTED_TYPES.items():
            if mime_type in mime_list:
                return category
        return 'unknown'


# =============================================================================
# 文本提取器
# =============================================================================

class TextExtractor:
    """文本提取器基类"""
    
    def extract(self, file_path: str) -> List[str]:
        """提取文本的抽象方法"""
        raise NotImplementedError


class PlainTextExtractor(TextExtractor):
    """纯文本提取器"""
    
    def extract(self, file_path: str) -> List[str]:
        """提取纯文本文件内容"""
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin1']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read().strip()
                    if content:
                        logger.info(f"成功使用 {encoding} 编码读取文件")
                        return [content]
            except UnicodeDecodeError:
                continue
            except Exception as e:
                logger.error(f"读取文件失败: {e}")
                break
        
        raise ValueError(f"无法读取文件 {file_path}，尝试了所有编码")


class PDFExtractor(TextExtractor):
    """PDF文本提取器 - 支持纯文本提取或直接OCR"""
    
    def __init__(self, enable_ocr: bool = True, use_gpu: bool = False):
        """
        初始化PDF提取器
        
        Args:
            enable_ocr: 是否使用OCR模式（True=直接OCR, False=仅纯文本提取）
            use_gpu: 是否使用GPU加速（仅在OCR模式时生效）
        """
        self.enable_ocr = enable_ocr
        self.use_gpu = use_gpu
        self.ocr_extractor = None  # 延迟初始化，只在需要时创建
        logger.debug(f"PDFExtractor初始化: enable_ocr={enable_ocr}, use_gpu={use_gpu}")
        
        if enable_ocr:
            logger.info("📋 PDF处理模式: 智能识别（文本型直接提取，扫描版使用OCR）")
        else:
            logger.info("📋 PDF处理模式: 仅纯文本提取")
    
    def extract(self, file_path: str) -> List[str]:
        """提取PDF文件内容"""
        # 先尝试直接提取文本（文本型PDF）
        try:
            texts = self._try_extract_text_pdf(file_path)
            if texts:
                return texts
        except Exception as e:
            logger.info(f"文本型PDF提取失败，将尝试OCR: {e}")
        
        # 如果直接提取失败且启用了OCR，使用OCR处理
        if self.enable_ocr:
            logger.info(f"🔍 使用OCR模式处理PDF: {Path(file_path).name}")
            return self._extract_with_ocr(file_path)
        else:
            raise ValueError("无法提取PDF文本，请启用OCR功能处理扫描版PDF")
    
    def _try_extract_text_pdf(self, file_path: str) -> Optional[List[str]]:
        """尝试提取文本型PDF"""
        if not deps.is_available('pdf'):
            return None
        
        logger.info(f"📄 尝试文本型PDF提取: {Path(file_path).name}")
        print(f"🔍 检测PDF类型...")
        
        # 优先使用pdfminer.six
        if 'pdfminer.six' in deps.available_modules['pdf']:
            try:
                from pdfminer.high_level import extract_text
                from pdfminer.layout import LAParams
                
                laparams = LAParams(
                    line_overlap=0.5,
                    char_margin=2.0,
                    line_margin=0.5,
                    word_margin=0.3,
                    boxes_flow=0.5,
                    detect_vertical=True,
                    all_texts=True
                )
                
                full_text = extract_text(file_path, laparams=laparams)
                text_stripped = full_text.strip() if full_text else ""
                
                # 判断是否为文本型PDF（每页平均100字符以上）
                if text_stripped and len(text_stripped) > 100:
                    logger.info(f"✅ 检测到文本型PDF，共{len(text_stripped)}字符")
                    print(f"✅ 检测到文本型PDF（包含可复制文字）")
                    print(f"   共{len(text_stripped):,}字符")
                    print(f"   使用直接提取模式（快速）")
                    
                    # 文本后处理
                    import re
                    full_text = re.sub(r'([a-z])([A-Z])', r'\1 \2', full_text)
                    full_text = re.sub(r'([A-Z]{2,})([A-Z][a-z])', r'\1 \2', full_text)
                    full_text = re.sub(r'(\d)([A-Za-z])', r'\1 \2', full_text)
                    full_text = re.sub(r'([A-Za-z])(\d)', r'\1 \2', full_text)
                    full_text = re.sub(r'\s+', ' ', full_text)
                    
                    return [full_text.strip()]
                else:
                    logger.info(f"检测到扫描版PDF，需要使用OCR")
                    print(f"⚠️  检测到扫描版PDF（图片型）")
                    print(f"   将使用OCR识别（较慢）")
                    return None
                    
            except Exception as e:
                logger.warning(f"pdfminer提取失败: {e}")
                return None
        
        return None
    
    def _get_pdf_page_count(self, file_path: str) -> int:
        """获取PDF页数"""
        try:
            import PyPDF2
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                return len(reader.pages)
        except Exception as e:
            logger.warning(f"无法获取PDF页数: {e}")
            return 0
    
    def _extract_with_ocr(self, file_path: str) -> List[str]:
        """使用OCR处理扫描版PDF"""
        try:
            logger.info(f"使用OCR处理PDF: {Path(file_path).name}")
            
            # 懒加载：只在需要OCR时才初始化科大讯飞OCR
            if self.ocr_extractor is None:
                if not self.enable_ocr:
                    raise ValueError("OCR功能未启用")
                
                if not deps.is_available('ocr'):
                    raise ValueError("OCR库不可用，请确保已安装requests库和配置讯飞API")
                
                logger.info("🔄 正在初始化科大讯飞OCR引擎...")
                try:
                    from xunfei_ocr import XunfeiOCRExtractor
                    from config import XUNFEI_OCR_CONFIG
                    
                    app_id = XUNFEI_OCR_CONFIG.get('app_id')
                    secret = XUNFEI_OCR_CONFIG.get('secret')
                    
                    if not app_id or not secret or app_id == 'your-xunfei-app-id':
                        raise ValueError("请在config.py中配置科大讯飞的 app_id 和 secret")
                    
                    self.ocr_extractor = XunfeiOCRExtractor(app_id=app_id, secret=secret)
                    logger.info("✅ 科大讯飞OCR初始化完成")
                except Exception as init_error:
                    logger.error(f"❌ 科大讯飞OCR初始化失败: {init_error}")
                    raise ValueError(f"无法初始化OCR引擎: {init_error}")
            
            # 使用OCR提取器处理
            try:
                result = self.ocr_extractor.extract(file_path)
                print(f"✅ OCR成功处理PDF\n")
                logger.info(f"✅ OCR成功处理PDF")
                return result
            except KeyboardInterrupt:
                logger.warning("用户中断OCR处理")
                raise
            except Exception as ocr_error:
                logger.error(f"OCR处理失败: {type(ocr_error).__name__}: {ocr_error}")
                raise RuntimeError(f"OCR处理失败: {ocr_error}")
            
        except KeyboardInterrupt:
            raise
        except Exception as e:
            logger.error(f"OCR处理PDF失败: {e}")
            raise ValueError(f"OCR处理失败: {e}")


class DOCXExtractor(TextExtractor):
    """DOCX文档提取器"""
    
    def extract(self, file_path: str) -> List[str]:
        """提取DOCX文件内容"""
        if not deps.is_available('docx'):
            raise RuntimeError("DOCX处理库不可用，请安装 python-docx")
        
        try:
            from docx import Document
            
            doc = Document(file_path)
            paragraphs = []
            
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    paragraphs.append(text)
            
            if not paragraphs:
                raise ValueError("DOCX文件为空或无文本内容")
            
            # 将段落合并为完整文档
            full_text = '\n\n'.join(paragraphs)
            logger.info(f"成功提取DOCX文档，共 {len(paragraphs)} 个段落")
            
            return [f"[DOCX文档]\n{full_text}"]
            
        except Exception as e:
            logger.error(f"DOCX提取失败: {e}")
            raise ValueError(f"DOCX文件处理失败: {e}")


class ImageExtractorWrapper(TextExtractor):
    """图片OCR提取器包装类"""
    
    def __init__(self, use_gpu: bool = False):
        """初始化图片提取器包装器"""
        self.use_gpu = use_gpu
        self.ocr_extractor = None
    
    def extract(self, file_path: str) -> List[str]:
        """提取图片文本"""
        logger.warning("科大讯飞OCR当前仅支持PDF文件")
        raise ValueError(
            f"科大讯飞OCR暂不支持单独的图片文件。\n"
            f"请将图片转换为PDF格式后再处理。\n"
            f"文件: {Path(file_path).name}"
        )


# =============================================================================
# 主文件处理器
# =============================================================================

class FileProcessor:
    """统一文件处理器"""
    
    def __init__(self, use_gpu: bool = False, enable_ocr: bool = True):
        """
        初始化文件处理器
        
        Args:
            use_gpu: 是否使用GPU加速OCR（科大讯飞是云端API，此参数无效）
            enable_ocr: 是否启用OCR功能（用于扫描版PDF）
        """
        self.use_gpu = use_gpu
        self.enable_ocr = enable_ocr
        self.extractors = self._init_extractors()
    
    def _init_extractors(self) -> Dict[str, TextExtractor]:
        """初始化提取器"""
        extractors = {
            'text': PlainTextExtractor(),
        }
        
        # PDF提取器（支持懒加载OCR）
        if deps.is_available('pdf'):
            # 只有在用户启用OCR且OCR库可用时才启用OCR
            ocr_available = deps.is_available('ocr')
            enable_ocr = self.enable_ocr and ocr_available
            
            logger.info(f"📄 初始化PDF提取器: OCR={'启用' if enable_ocr else '禁用'}")
            extractors['pdf'] = PDFExtractor(enable_ocr=enable_ocr, use_gpu=self.use_gpu)
            
            if enable_ocr:
                logger.info("✅ PDF提取器已启用OCR功能（懒加载模式）")
            elif self.enable_ocr and not ocr_available:
                logger.warning("⚠️  用户启用了OCR但科大讯飞OCR未配置，OCR功能不可用")
            else:
                logger.info("ℹ️  PDF提取器OCR功能已禁用（用户选择）")
        
        # 图片OCR提取器（懒加载）
        if self.enable_ocr and deps.is_available('ocr'):
            extractors['image'] = ImageExtractorWrapper(use_gpu=self.use_gpu)
        
        if deps.is_available('docx'):
            extractors['docx'] = DOCXExtractor()
            extractors['doc'] = DOCXExtractor()
        
        return extractors
    
    def process_file(self, file_path: str) -> Tuple[str, List[str]]:
        """
        处理文件并提取文本
        
        Args:
            file_path: 文件路径
            
        Returns:
            Tuple[str, List[str]]: (文件类型, 提取的文本列表)
        """
        try:
            # 检测文件类型
            file_type, mime_type = FileTypeDetector.detect_file_type(file_path)
            logger.info(f"检测到文件类型: {file_type} ({mime_type})")
            
            # 获取对应的提取器
            extractor = self.extractors.get(file_type)
            if not extractor:
                raise ValueError(f"不支持的文件类型: {file_type}")
            
            # 提取文本
            texts = extractor.extract(file_path)
            
            # 验证结果
            if not texts or not any(text.strip() for text in texts):
                raise ValueError("文件中未提取到有效文本")
            
            logger.info(f"成功处理文件 {Path(file_path).name}: {len(texts)} 个文本块")
            return file_type, texts
            
        except Exception as e:
            logger.error(f"文件处理失败 {file_path}: {e}")
            raise
    
    def get_supported_formats(self) -> List[str]:
        """获取支持的文件格式列表"""
        formats = ['txt', 'md', 'html', 'xml']
        
        if 'pdf' in self.extractors:
            formats.append('pdf')
        
        if 'docx' in self.extractors:
            formats.extend(['docx', 'doc'])
        
        if 'image' in self.extractors:
            formats.extend(['jpg', 'jpeg', 'png', 'tiff', 'bmp', 'gif'])
        
        return formats
    
    def get_processor_info(self) -> Dict[str, Any]:
        """获取处理器信息"""
        return {
            "supported_formats": self.get_supported_formats(),
            "available_extractors": list(self.extractors.keys()),
            "dependencies": deps.available_modules,
            "ocr_enabled": 'image' in self.extractors,
        }
    
    def save_extracted_text(self, page_texts: List[str], output_path: str):
        """
        保存提取的文本到文件
        
        Args:
            page_texts: 文本列表
            output_path: 输出文件路径
        """
        try:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                for text in page_texts:
                    f.write(text)
                    f.write("\n\n")
            
            logger.info(f"文本已保存到: {output_path}")
            
        except Exception as e:
            logger.error(f"保存文本失败: {e}")
            raise


# =============================================================================
# 工具函数
# =============================================================================

def create_file_processor(enable_ocr: bool = True) -> FileProcessor:
    """
    创建文件处理器实例
    
    Args:
        enable_ocr: 是否启用OCR功能
        
    Returns:
        FileProcessor: 文件处理器实例
    """
    return FileProcessor(enable_ocr=enable_ocr)


def get_file_info(file_path: str) -> Dict[str, Any]:
    """
    获取文件信息
    
    Args:
        file_path: 文件路径
        
    Returns:
        Dict: 文件信息
    """
    path = Path(file_path)
    
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    file_type, mime_type = FileTypeDetector.detect_file_type(file_path)
    
    return {
        "name": path.name,
        "size": path.stat().st_size,
        "extension": path.suffix,
        "type": file_type,
        "mime_type": mime_type,
        "is_supported": file_type in ['text', 'pdf', 'docx', 'doc', 'image'],
    }


if __name__ == "__main__":
    # 简单测试
    processor = create_file_processor()
    info = processor.get_processor_info()
    
    print("文件处理器信息:")
    print(f"支持格式: {info['supported_formats']}")
    print(f"可用提取器: {info['available_extractors']}")
    print(f"OCR功能: {'启用' if info['ocr_enabled'] else '禁用'}")
