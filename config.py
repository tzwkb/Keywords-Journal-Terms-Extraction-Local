#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
期刊术语提取工具配置文件
"""

import os

# =============================================================================
# OpenAI API 配置
# =============================================================================

# OpenAI API密钥和端点
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "")

# 使用的模型
DEFAULT_MODEL = "gpt-4o"

# =============================================================================
# 模型配置辅助函数
# =============================================================================

def get_token_param_name(model: str) -> str:
    """
    根据模型名称返回正确的token参数名
    
    Args:
        model: 模型名称
        
    Returns:
        str: token参数名（"max_tokens" 或 "max_completion_tokens"）
    """
    # o1系列模型使用max_completion_tokens
    if "o1" in model.lower():
        return "max_completion_tokens"
    # 其他模型使用max_tokens
    return "max_tokens"

# =============================================================================
# OCR配置 - 科大讯飞
# =============================================================================

# 科大讯飞OCR配置
XUNFEI_OCR_CONFIG = {
    # 从环境变量读取，或在这里直接设置（不推荐）
    "app_id": os.getenv("XUNFEI_APP_ID", "your-xunfei-app-id"),  # 讯飞开放平台AppID
    "secret": os.getenv("XUNFEI_SECRET", "your-xunfei-secret"),  # 讯飞开放平台Secret
    
    # OCR任务配置
    "export_format": "txt",  # 导出格式：txt, word, markdown, json
    "max_wait_time": 300,    # 最大等待时间（秒）
    "check_interval": 5,     # 状态检查间隔（秒）
    
    # 是否启用OCR功能
    "enabled": True,  # True=启用科大讯飞OCR, False=禁用
}

# =============================================================================
# 输出配置
# =============================================================================

# 输出目录（None表示使用PDF所在目录）
OUTPUT_DIR = None

# Excel样式配置
EXCEL_STYLES = {
    "keywords_header_color": "366092",  # 关键词表头颜色（蓝色）
    "abstract_header_color": "2E7D32",  # 摘要表头颜色（绿色）
    "header_font_color": "FFFFFF",      # 表头字体颜色（白色）
}

# =============================================================================
# 关键词和摘要识别配置
# =============================================================================

# 关键词识别的正则表达式模式
KEYWORDS_PATTERNS = {
    "zh": [
        # 更宽松的匹配，允许关键词标题和内容之间有多行空白或其他内容
        r'关键词\s*[:：]+.*?[\n\s]+([\u4e00-\u9fff；;，,、\s]+?)(?=\n\s*(?:Abstract|ABSTRACT|Key\s*words|中图分类号|DOI|文献标识码|$))',
        r'关键字\s*[:：]+.*?[\n\s]+([\u4e00-\u9fff；;，,、\s]+?)(?=\n\s*(?:Abstract|ABSTRACT|Key\s*words|中图分类号|DOI|文献标识码|$))',
        # 保留原有的紧密匹配作为备选
        r'关键词\s*[:：]?\s*(.*?)(?=\n\s*(?:Abstract|ABSTRACT|摘要|引言|Introduction|1\s+|$))',
    ],
    "en": [
        # 更宽松的匹配，允许跨行
        r'Key\s*words?\s*[:：]+\s*[\n\s]*([a-zA-Z\-；;，,、\s]+?)(?=\n\s*(?:摘要|Introduction|引言|1\s+|中图分类号|基金项目|$))',
        r'KEYWORDS?\s*[:：]+\s*[\n\s]*([a-zA-Z\-；;，,、\s]+?)(?=\n\s*(?:摘要|Introduction|引言|1\s+|中图分类号|基金项目|$))',
        # 保留原有的紧密匹配作为备选
        r'Key\s*words?\s*[:：]?\s*(.*?)(?=\n\s*(?:摘要|Introduction|引言|1\s+|$))',
    ]
}

# 摘要识别的正则表达式模式
ABSTRACT_PATTERNS = {
    "zh": [
        # 允许摘要标题和内容之间有空行
        r'摘\s*要\s*[:：]+\s*(.*?)(?=\n\s*(?:关键词|Key\s*words?|Abstract|ABSTRACT|中图分类号|DOI|引言|Introduction|1\s+|$))',
        r'摘要\s*[:：]+\s*(.*?)(?=\n\s*(?:关键词|Key\s*words?|Abstract|ABSTRACT|中图分类号|DOI|引言|Introduction|1\s+|$))',
    ],
    "en": [
        # 允许Abstract标题和内容之间有空行
        r'Abstract\s*[:：]+\s*(.*?)(?=\n\s*(?:Key\s*words?|KEYWORDS?|摘要|基金项目|收稿日期|Introduction|引言|1\s+|$))',
        r'ABSTRACT\s*[:：]+\s*(.*?)(?=\n\s*(?:Key\s*words?|KEYWORDS?|摘要|基金项目|收稿日期|Introduction|引言|1\s+|$))',
    ]
}

# 关键词分隔符
KEYWORD_SEPARATORS = [';', '；', ',', '，', '\n', '  ']  # 按优先级排序

# =============================================================================
# GPT提取配置
# =============================================================================

# GPT处理参数
GPT_CONFIG = {
    "model": DEFAULT_MODEL,
    "temperature": 0,        # 完全确定性输出
    "max_tokens": 4096,      # 最大输出token数
    "timeout": 60,           # 超时时间（秒）
}

# GPT系统提示词
GPT_SYSTEM_PROMPT = """You are a technical terminology extraction specialist for academic papers. Your task is to extract bilingual technical term pairs from academic abstracts."""

# GPT用户提示词模板（针对单篇文章）
GPT_USER_PROMPT_TEMPLATE = """Please extract technical term pairs from the following abstract.

For each technical term you identify, provide both its English and Chinese versions.

EXTRACTION RULES:
1. Focus on domain-specific technical terms, concepts, and specialized vocabulary
2. EVERY term must have BOTH English and Chinese versions
3. If abstract is in Chinese, provide English translation
4. If abstract is in English, provide Chinese translation
5. Extract 8-12 most important technical terms from THIS SINGLE abstract
6. Avoid generic words unless they have specific technical meaning
7. Prioritize aerospace, aviation, and engineering terminology
8. Include method names, material names, key concepts, and technical processes

OUTPUT FORMAT (strict JSON):
{{
  "terms": [
    {{
      "en_term": "English technical term",
      "zh_term": "对应的中文术语"
    }}
  ]
}}

ABSTRACT TEXT:
{abstract_text}

Please extract the technical terms now:"""

# =============================================================================
# 术语处理配置
# =============================================================================

TERM_PROCESSING = {
    "case_sensitive_matching": False,  # 术语匹配是否区分大小写
    "remove_duplicates": True,          # 是否移除重复术语
}

# =============================================================================
# 日志配置
# =============================================================================

LOG_CONFIG = {
    "level": "INFO",
    "format": "%(asctime)s - %(levelname)s - %(message)s",
    "file": "journal_extractor.log",
}

