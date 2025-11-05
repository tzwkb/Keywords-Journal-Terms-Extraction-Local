#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
科大讯飞PDF OCR API封装
支持PDF文件的OCR识别
"""

import hashlib
import hmac
import base64
import time
import requests
import logging
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)


class XunfeiOCR:
    """科大讯飞PDF OCR客户端"""
    
    def __init__(self, app_id: str, secret: str):
        """
        初始化科大讯飞OCR客户端
        
        Args:
            app_id: 讯飞开放平台的 appId
            secret: 讯飞开放平台的 secret
        """
        self.app_id = app_id
        self.secret = secret
        self.base_url = "https://iocr.xfyun.cn/ocrzdq/v1/pdfOcr"
        
        if not app_id or not secret:
            raise ValueError("科大讯飞OCR需要配置 app_id 和 secret")
        
        logger.info(f"科大讯飞OCR客户端初始化成功 (AppID: {app_id[:8]}***)")
    
    def _get_signature(self) -> tuple:
        """
        生成API签名
        
        Returns:
            tuple: (timestamp, signature)
        """
        timestamp = str(int(time.time()))
        auth = hashlib.md5((self.app_id + timestamp).encode('utf-8')).hexdigest()
        signature = hmac.new(
            self.secret.encode('utf-8'), 
            auth.encode('utf-8'), 
            hashlib.sha1
        ).digest()
        signature = base64.b64encode(signature).decode('utf-8')
        return timestamp, signature
    
    def start_ocr_task(self, pdf_path: str, export_format: str = "txt") -> str:
        """
        启动OCR识别任务
        
        Args:
            pdf_path: PDF文件路径
            export_format: 导出格式 (txt, word, markdown, json)
            
        Returns:
            str: 任务ID
        """
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
        
        timestamp, signature = self._get_signature()
        
        headers = {
            'appId': self.app_id,
            'timestamp': timestamp,
            'signature': signature
        }
        
        logger.info(f"上传PDF文件到科大讯飞OCR: {Path(pdf_path).name}")
        
        try:
            with open(pdf_path, 'rb') as f:
                files = {'file': (Path(pdf_path).name, f, 'application/pdf')}
                data = {'exportFormat': export_format}
                
                response = requests.post(
                    f"{self.base_url}/start",
                    headers=headers,
                    files=files,
                    data=data,
                    timeout=60
                )
            
            result = response.json()
            
            if result.get('code') == 0:
                task_id = result['data']['taskId']
                logger.info(f"OCR任务创建成功, taskId: {task_id}")
                return task_id
            else:
                error_msg = result.get('desc', '未知错误')
                raise RuntimeError(f"启动OCR任务失败: {error_msg}")
                
        except requests.RequestException as e:
            logger.error(f"网络请求失败: {e}")
            raise RuntimeError(f"无法连接到科大讯飞OCR服务: {e}")
    
    def get_task_result(self, task_id: str, max_wait_time: int = 300) -> str:
        """
        查询OCR任务结果并下载
        
        Args:
            task_id: 任务ID
            max_wait_time: 最大等待时间（秒）
            
        Returns:
            str: 提取的文本内容
        """
        timestamp, signature = self._get_signature()
        
        headers = {
            'appId': self.app_id,
            'timestamp': timestamp,
            'signature': signature
        }
        
        start_time = time.time()
        check_interval = 5  # 每5秒查询一次
        
        logger.info("等待OCR任务完成...")
        
        while True:
            # 检查是否超时
            if time.time() - start_time > max_wait_time:
                raise TimeoutError(f"OCR任务超时（{max_wait_time}秒）")
            
            try:
                response = requests.get(
                    f"{self.base_url}/getResult",
                    headers=headers,
                    params={'taskId': task_id},
                    timeout=30
                )
                
                result = response.json()
                
                if result.get('code') != 0:
                    error_msg = result.get('desc', '未知错误')
                    raise RuntimeError(f"查询任务状态失败: {error_msg}")
                
                status = result['data']['status']
                
                if status == '3':  # 任务完成
                    logger.info("OCR任务完成，正在下载结果...")
                    download_url = result['data']['downloadUrl']
                    return self._download_result(download_url)
                
                elif status == '2':  # 处理中
                    elapsed = int(time.time() - start_time)
                    logger.info(f"OCR处理中... (已等待 {elapsed}秒)")
                    time.sleep(check_interval)
                
                elif status == '4':  # 任务失败
                    error_desc = result['data'].get('desc', '未知错误')
                    raise RuntimeError(f"OCR任务失败: {error_desc}")
                
                else:
                    logger.warning(f"未知任务状态: {status}")
                    time.sleep(check_interval)
                    
            except requests.RequestException as e:
                logger.error(f"查询任务状态时网络错误: {e}")
                time.sleep(check_interval)
    
    def _download_result(self, download_url: str) -> str:
        """
        下载OCR结果文本
        
        Args:
            download_url: 结果下载链接
            
        Returns:
            str: 文本内容
        """
        try:
            response = requests.get(download_url, timeout=60)
            response.raise_for_status()
            
            # 尝试使用UTF-8解码，如果失败则尝试GBK
            try:
                text = response.content.decode('utf-8')
            except UnicodeDecodeError:
                text = response.content.decode('gbk', errors='ignore')
            
            logger.info(f"OCR结果下载成功，文本长度: {len(text)} 字符")
            return text
            
        except requests.RequestException as e:
            logger.error(f"下载OCR结果失败: {e}")
            raise RuntimeError(f"无法下载OCR结果: {e}")
    
    def ocr_pdf(self, pdf_path: str, export_format: str = "txt") -> str:
        """
        一站式PDF OCR处理（启动任务 + 等待 + 获取结果）
        
        Args:
            pdf_path: PDF文件路径
            export_format: 导出格式 (txt, word, markdown, json)
            
        Returns:
            str: 提取的文本内容
        """
        logger.info(f"开始OCR处理: {Path(pdf_path).name}")
        
        # 启动任务
        task_id = self.start_ocr_task(pdf_path, export_format)
        
        # 等待并获取结果
        text = self.get_task_result(task_id)
        
        logger.info(f"OCR处理完成: {Path(pdf_path).name}")
        return text


class XunfeiOCRExtractor:
    """科大讯飞OCR文本提取器（适配file_processor接口）"""
    
    def __init__(self, app_id: str, secret: str):
        """
        初始化科大讯飞OCR提取器
        
        Args:
            app_id: 讯飞开放平台的 appId
            secret: 讯飞开放平台的 secret
        """
        try:
            self.ocr = XunfeiOCR(app_id=app_id, secret=secret)
            print("✅ 科大讯飞OCR引擎初始化完成")
            logger.info("科大讯飞OCR提取器初始化成功")
        except Exception as e:
            logger.error(f"科大讯飞OCR初始化失败: {e}")
            raise RuntimeError(f"科大讯飞OCR初始化失败: {e}")
    
    def extract(self, file_path: str) -> List[str]:
        """
        从PDF提取文本（file_processor统一接口）
        
        Args:
            file_path: 文件路径
            
        Returns:
            List[str]: 提取的文本列表
        """
        try:
            file_ext = Path(file_path).suffix.lower()
            file_name = Path(file_path).name
            
            # 检查文件类型
            if file_ext == '.pdf':
                logger.info(f"使用科大讯飞OCR处理PDF: {file_name}")
                print(f"📄 正在使用科大讯飞OCR处理PDF...")
                print("⏳ 这可能需要几分钟，请耐心等待...")
                
                # 调用讯飞OCR
                text = self.ocr.ocr_pdf(file_path, export_format="txt")
                
                if not text or not text.strip():
                    raise ValueError("PDF中未检测到有效文本")
                
                logger.info(f"OCR成功，提取{len(text)}字符")
                print(f"✅ OCR完成，提取 {len(text):,} 字符")
                
                # 返回格式化的文本
                return [f"[扫描版PDF - {Path(file_path).name}]\n{text.strip()}"]
            else:
                raise ValueError(
                    f"科大讯飞OCR仅支持PDF文件。\n"
                    f"请将图片转换为PDF格式后再处理。\n"
                    f"文件: {file_name}"
                )
                
        except Exception as e:
            logger.error(f"科大讯飞OCR提取失败: {e}")
            raise ValueError(f"OCR处理失败: {e}")


def test_xunfei_ocr():
    """测试科大讯飞OCR功能"""
    import os
    
    # 从环境变量读取配置
    app_id = os.getenv("XUNFEI_APP_ID")
    secret = os.getenv("XUNFEI_SECRET")
    
    if not app_id or not secret:
        print("请设置环境变量: XUNFEI_APP_ID 和 XUNFEI_SECRET")
        return
    
    # 创建提取器
    extractor = XunfeiOCRExtractor(app_id=app_id, secret=secret)
    
    # 测试文件
    test_pdf = "test.pdf"
    
    if not Path(test_pdf).exists():
        print(f"测试文件不存在: {test_pdf}")
        return
    
    # 执行OCR
    try:
        texts = extractor.extract(test_pdf)
        print(f"\n提取结果（前500字符）:\n{texts[0][:500]}")
    except Exception as e:
        print(f"OCR失败: {e}")


if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.INFO)
    
    # 运行测试
    test_xunfei_ocr()

