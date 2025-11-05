#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF文本提取工具（期刊版）
将PDF文件提取为文本，保存到extracted_texts文件夹
支持文本型PDF（快速）和扫描版PDF（科大讯飞OCR）
"""

import sys
import argparse
from pathlib import Path
from typing import List, Dict, Any, Optional
import logging

# 导入模块
try:
    from file_processor import FileProcessor, get_file_info
except ImportError as e:
    print(f"❌ 导入模块失败: {e}")
    sys.exit(1)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class BatchFileExtractor:
    """批量文件提取器"""
    
    def __init__(self, enable_ocr: bool = True):
        """
        初始化批量提取器
        
        Args:
            enable_ocr: 是否启用OCR功能（用于扫描版PDF）
        """
        self.processor = FileProcessor(use_gpu=False, enable_ocr=enable_ocr)
        self.stats = {
            "total_files": 0,
            "success_count": 0,
            "failed_count": 0,
            "skipped_count": 0,
            "total_size": 0,
            "extracted_chars": 0
        }
    
    def extract_all_files(self, 
                         source_dir: str = "pdf",
                         output_dir: str = "extracted_texts",
                         overwrite: bool = False,
                         file_types: Optional[List[str]] = None) -> Dict[str, Any]:
        """
        提取所有支持的文件
        
        Args:
            source_dir: 源文件目录
            output_dir: 输出目录
            overwrite: 是否覆盖已存在的文件
            file_types: 要处理的文件类型列表，None表示处理所有支持的类型
            
        Returns:
            Dict: 处理结果统计
        """
        source_path = Path(source_dir)
        output_path = Path(output_dir)
        
        # 验证和创建目录
        if not self._setup_directories(source_path, output_path):
            return self.stats
        
        # 扫描文件
        files_to_process = self._scan_files(source_path, file_types)
        if not files_to_process:
            return self.stats
        
        # 显示处理计划
        self._display_processing_plan(files_to_process)
        
        # 批量处理文件
        self._process_files_batch(files_to_process, output_path, overwrite)
        
        # 显示最终统计
        self._display_final_stats()
        
        return self.stats
    
    def extract_single_file(self, 
                           file_path: str,
                           output_dir: str = "extracted_texts",
                           custom_name: Optional[str] = None) -> bool:
        """
        提取单个文件
        
        Args:
            file_path: 文件路径
            output_dir: 输出目录
            custom_name: 自定义输出文件名
            
        Returns:
            bool: 是否成功
        """
        try:
            source_path = Path(file_path)
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            
            if not source_path.exists():
                logger.error(f"文件不存在: {file_path}")
                return False
            
            # 处理文件
            return self._process_single_file(source_path, output_path, custom_name)
            
        except Exception as e:
            logger.error(f"处理文件失败 {file_path}: {e}")
            return False
    
    def _setup_directories(self, source_dir: Path, output_dir: Path) -> bool:
        """设置和验证目录"""
        if not source_dir.exists():
            logger.error(f"源目录不存在: {source_dir}")
            print(f"❌ '{source_dir}' 文件夹不存在")
            return False
        
        try:
            output_dir.mkdir(exist_ok=True)
            logger.info(f"输出目录: {output_dir}")
        except Exception as e:
            logger.error(f"创建输出目录失败: {e}")
            return False
        
        return True
    
    def _scan_files(self, source_dir: Path, file_types: Optional[List[str]]) -> List[Path]:
        """扫描要处理的文件"""
        supported_formats = self.processor.get_supported_formats()
        
        if file_types:
            target_types = [t.lower().lstrip('.') for t in file_types]
            supported_formats = [f for f in supported_formats if f in target_types]
        
        files_to_process = []
        
        for fmt in supported_formats:
            pattern = f"*.{fmt}"
            found_files = list(source_dir.glob(pattern))
            files_to_process.extend(found_files)
        
        files_to_process.sort(key=lambda x: x.name.lower())
        
        self.stats["total_files"] = len(files_to_process)
        
        if not files_to_process:
            print(f"❌ 在 '{source_dir}' 中未找到支持的文件")
            print(f"支持的格式: {', '.join(supported_formats)}")
        
        return files_to_process
    
    def _display_processing_plan(self, files: List[Path]):
        """显示处理计划"""
        print(f"\n📚 找到 {len(files)} 个文件待处理")
        print("=" * 70)
        
        type_stats = {}
        total_size = 0
        
        for file_path in files:
            try:
                file_info = get_file_info(str(file_path))
                file_type = file_info['type']
                file_size = file_info['size']
                
                if file_type not in type_stats:
                    type_stats[file_type] = {"count": 0, "size": 0}
                
                type_stats[file_type]["count"] += 1
                type_stats[file_type]["size"] += file_size
                total_size += file_size
                
            except Exception as e:
                logger.warning(f"获取文件信息失败 {file_path.name}: {e}")
        
        print("文件类型统计:")
        for file_type, stats in type_stats.items():
            size_mb = stats["size"] / (1024 * 1024)
            print(f"  {file_type.upper()}: {stats['count']} 个文件, {size_mb:.1f} MB")
        
        self.stats["total_size"] = total_size
        print(f"\n总大小: {total_size / (1024 * 1024):.1f} MB")
        print("-" * 70)
    
    def _process_files_batch(self, files: List[Path], output_dir: Path, overwrite: bool):
        """批量处理文件"""
        for i, file_path in enumerate(files, 1):
            print(f"\n📄 处理文件 {i}/{len(files)}: {file_path.name}")
            
            try:
                output_file = output_dir / f"{file_path.stem}.txt"
                if output_file.exists() and not overwrite:
                    print(f"   ⏭️  跳过（文件已存在）: {output_file.name}")
                    self.stats["skipped_count"] += 1
                    continue
                
                success = self._process_single_file(file_path, output_dir)
                
                if success:
                    self.stats["success_count"] += 1
                else:
                    self.stats["failed_count"] += 1
                    
            except KeyboardInterrupt:
                print("\n⏹️  用户中断处理")
                break
            except Exception as e:
                logger.error(f"处理文件异常 {file_path.name}: {e}")
                self.stats["failed_count"] += 1
    
    def _process_single_file(self, 
                           file_path: Path, 
                           output_dir: Path,
                           custom_name: Optional[str] = None) -> bool:
        """处理单个文件"""
        try:
            file_size = file_path.stat().st_size
            print(f"   文件大小: {file_size / 1024:.1f} KB")
            
            # 提取文本
            file_type, text_list = self.processor.process_file(str(file_path))
            
            if not text_list or not any(text.strip() for text in text_list):
                print(f"   ⚠️  文件为空或无法提取文本")
                return False
            
            # 合并所有文本
            full_text = self._combine_texts(text_list, file_path.name)
            
            # 保存文件
            output_name = custom_name or f"{file_path.stem}.txt"
            output_file = output_dir / output_name
            
            success = self._save_extracted_text(full_text, output_file, file_path)
            
            if success:
                char_count = len(full_text)
                self.stats["extracted_chars"] += char_count
                print(f"   ✅ 成功提取 {char_count:,} 字符")
                print(f"   💾 保存至: {output_file.name}")
            
            return success
            
        except Exception as e:
            print(f"   ❌ 处理失败: {e}")
            logger.error(f"处理文件失败 {file_path}: {e}")
            return False
    
    def _combine_texts(self, text_list: List[str], filename: str) -> str:
        """合并文本列表"""
        if len(text_list) == 1:
            return text_list[0]
        
        combined_parts = []
        for i, text in enumerate(text_list, 1):
            if text.strip():
                if not text.startswith('[页面') and not text.startswith('[文件'):
                    combined_parts.append(f"[页面 {i}]\n{text.strip()}")
                else:
                    combined_parts.append(text.strip())
        
        return '\n\n'.join(combined_parts)
    
    def _save_extracted_text(self, text: str, output_file: Path, source_file: Path) -> bool:
        """保存提取的文本"""
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(text)
            
            return True
            
        except Exception as e:
            logger.error(f"保存文件失败 {output_file}: {e}")
            return False
    
    def _display_final_stats(self):
        """显示最终统计信息"""
        print(f"\n" + "=" * 70)
        print("📊 处理完成统计")
        print("=" * 70)
        print(f"总文件数: {self.stats['total_files']}")
        print(f"成功处理: {self.stats['success_count']} ✅")
        print(f"处理失败: {self.stats['failed_count']} ❌")
        print(f"跳过文件: {self.stats['skipped_count']} ⏭️")
        
        if self.stats['success_count'] > 0:
            success_rate = (self.stats['success_count'] / self.stats['total_files']) * 100
            print(f"成功率: {success_rate:.1f}%")
            
            total_mb = self.stats['total_size'] / (1024 * 1024)
            chars_k = self.stats['extracted_chars'] / 1000
            print(f"处理数据: {total_mb:.1f} MB → {chars_k:.1f}K 字符")
        
        print("=" * 70)
    
    def get_processor_info(self) -> Dict[str, Any]:
        """获取处理器信息"""
        return self.processor.get_processor_info()


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description="期刊PDF文本提取工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python extract_pdf_texts.py                    # 处理pdf文件夹中的所有文件
  python extract_pdf_texts.py --types pdf        # 只处理PDF文件
  python extract_pdf_texts.py --file sample.pdf  # 处理单个文件
  python extract_pdf_texts.py --overwrite        # 覆盖已存在的文件
        """
    )
    
    parser.add_argument("--source", "-s", default="pdf",
                       help="源文件目录 (默认: pdf)")
    parser.add_argument("--output", "-o", default="extracted_texts",
                       help="输出目录 (默认: extracted_texts)")
    parser.add_argument("--types", "-t", nargs="+", 
                       help="要处理的文件类型 (如: pdf)")
    parser.add_argument("--file", "-f", help="处理单个文件")
    parser.add_argument("--overwrite", action="store_true",
                       help="覆盖已存在的文件")
    parser.add_argument("--disable-ocr", action="store_true",
                       help="禁用OCR功能（仅处理文本型PDF）")
    parser.add_argument("--info", action="store_true", help="显示处理器信息")
    
    args = parser.parse_args()
    
    # 创建提取器
    enable_ocr = not args.disable_ocr
    extractor = BatchFileExtractor(enable_ocr=enable_ocr)
    
    # 显示处理器信息
    if args.info:
        info = extractor.get_processor_info()
        print("📋 文件处理器信息:")
        print(f"支持格式: {', '.join(info['supported_formats'])}")
        print(f"可用提取器: {', '.join(info['available_extractors'])}")
        print(f"OCR功能: {'启用' if info['ocr_enabled'] else '禁用'}")
        return
    
    print("🎉 期刊PDF文本提取工具")
    print("=" * 50)
    
    try:
        if args.file:
            # 处理单个文件
            success = extractor.extract_single_file(
                args.file, 
                args.output,
                None
            )
            if success:
                print("✅ 文件处理完成")
            else:
                print("❌ 文件处理失败")
        else:
            # 批量处理
            stats = extractor.extract_all_files(
                source_dir=args.source,
                output_dir=args.output,
                overwrite=args.overwrite,
                file_types=args.types
            )
            
            if stats['success_count'] > 0:
                print(f"\n🎉 处理完成！成功提取 {stats['success_count']} 个文件")
            else:
                print("\n❌ 没有成功处理任何文件")
                
    except KeyboardInterrupt:
        print("\n⏹️  用户中断程序")
    except Exception as e:
        logger.error(f"程序执行失败: {e}")
        print(f"❌ 程序执行失败: {e}")


if __name__ == "__main__":
    main()

