#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GPT处理器 - 使用OpenAI API提取术语对
"""

import os
import sys
import json
import logging
import time
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)


class GPTProcessor:
    """GPT处理器 - 使用OpenAI API进行术语提取"""
    
    def __init__(
        self, 
        api_key: str,
        base_url: str = "https://api.openai.com/v1",
        enable_checkpoint: bool = False
    ):
        """
        初始化GPT处理器
        
        Args:
            api_key: OpenAI API密钥
            base_url: API端点URL
            enable_checkpoint: 是否启用检查点（保存中间结果）
        """
        self.api_key = api_key
        self.base_url = base_url
        self.enable_checkpoint = enable_checkpoint
        self.client = None
        
        self._init_client()
    
    def _init_client(self):
        """初始化OpenAI客户端"""
        try:
            from openai import OpenAI
            
            logger.info("正在初始化OpenAI客户端...")
            
            self.client = OpenAI(
                api_key=self.api_key,
                base_url=self.base_url
            )
            
            logger.info("OpenAI客户端初始化成功")
            
        except ImportError as e:
            logger.error(f"openai库未安装: {e}")
            raise ImportError(
                "openai库未安装，请运行: pip install openai"
            )
        except Exception as e:
            logger.error(f"OpenAI客户端初始化失败: {e}")
            raise
    
    def process_single_text(
        self,
        text: str,
        custom_id: str = "default",
        system_prompt: str = None,
        user_prompt_template: str = None,
        model: str = "gpt-4o",
        temperature: float = 0,
        max_tokens: int = 4096
    ) -> Dict[str, Any]:
        """
        处理单个文本并提取术语
        
        Args:
            text: 要处理的文本
            custom_id: 自定义ID
            system_prompt: 系统提示词
            user_prompt_template: 用户提示词模板
            model: 使用的模型
            temperature: 温度参数
            max_tokens: 最大token数
            
        Returns:
            Dict[str, Any]: 包含提取结果的字典
        """
        logger.info(f"开始处理文本，ID: {custom_id}")
        
        try:
            # 构建消息
            messages = []
            
            if system_prompt:
                messages.append({
                    "role": "system",
                    "content": system_prompt
                })
            
            # 用户消息
            user_content = user_prompt_template if user_prompt_template else text
            messages.append({
                "role": "user",
                "content": user_content
            })
            
            # 调用API
            logger.info(f"调用OpenAI API - 模型: {model}")
            
            response = self.client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens,
                response_format={"type": "json_object"}  # 强制JSON输出
            )
            
            # 解析响应
            content = response.choices[0].message.content
            
            logger.info(f"API调用成功，返回内容长度: {len(content)}")
            
            # 解析JSON
            try:
                extracted_terms = json.loads(content)
            except json.JSONDecodeError as e:
                logger.error(f"JSON解析失败: {e}")
                logger.error(f"原始内容: {content}")
                extracted_terms = {}
            
            # 构建结果
            result = {
                "custom_id": custom_id,
                "extracted_terms": extracted_terms,
                "model": model,
                "usage": {
                    "prompt_tokens": response.usage.prompt_tokens,
                    "completion_tokens": response.usage.completion_tokens,
                    "total_tokens": response.usage.total_tokens
                }
            }
            
            logger.info(f"处理完成 - Token使用: {result['usage']['total_tokens']}")
            
            return result
            
        except Exception as e:
            logger.error(f"处理文本失败: {e}")
            raise
    
    def process_batch(
        self,
        texts: list,
        system_prompt: str = None,
        user_prompt_template: str = None,
        model: str = "gpt-4o",
        temperature: float = 0,
        max_tokens: int = 4096,
        delay: float = 1.0
    ) -> list:
        """
        批量处理文本
        
        Args:
            texts: 文本列表
            system_prompt: 系统提示词
            user_prompt_template: 用户提示词模板
            model: 使用的模型
            temperature: 温度参数
            max_tokens: 最大token数
            delay: 每次请求之间的延迟（秒）
            
        Returns:
            list: 结果列表
        """
        logger.info(f"开始批量处理，共{len(texts)}个文本")
        
        results = []
        
        for idx, text in enumerate(texts):
            try:
                logger.info(f"处理进度: {idx + 1}/{len(texts)}")
                
                result = self.process_single_text(
                    text=text,
                    custom_id=f"batch_{idx}",
                    system_prompt=system_prompt,
                    user_prompt_template=user_prompt_template,
                    model=model,
                    temperature=temperature,
                    max_tokens=max_tokens
                )
                
                results.append(result)
                
                # 延迟以避免速率限制
                if idx < len(texts) - 1 and delay > 0:
                    time.sleep(delay)
                    
            except Exception as e:
                logger.error(f"处理第{idx + 1}个文本失败: {e}")
                results.append({
                    "custom_id": f"batch_{idx}",
                    "error": str(e)
                })
        
        logger.info(f"批量处理完成，成功: {len([r for r in results if 'error' not in r])}/{len(texts)}")
        
        return results


if __name__ == "__main__":
    # 测试代码
    import os
    
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # 从环境变量读取API密钥
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("错误: 请设置OPENAI_API_KEY环境变量")
        sys.exit(1)
    
    # 创建处理器
    processor = GPTProcessor(
        api_key=api_key,
        base_url=os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
    )
    
    # 测试文本
    test_text = """
    【中文摘要】
    本文研究了高超声速飞行器的气动特性。通过数值模拟和风洞实验，分析了马赫数对升阻比的影响。
    
    【English Abstract】
    This paper studies the aerodynamic characteristics of hypersonic vehicles. Through numerical simulation 
    and wind tunnel experiments, the influence of Mach number on lift-to-drag ratio was analyzed.
    """
    
    # 系统提示词
    system_prompt = "You are a technical terminology extraction specialist."
    
    # 用户提示词
    user_prompt = f"""Extract technical term pairs from the following abstract:

{test_text}

Return JSON format:
{{
  "terms": [
    {{"en_term": "English term", "zh_term": "中文术语"}}
  ]
}}
"""
    
    # 处理
    result = processor.process_single_text(
        text=test_text,
        system_prompt=system_prompt,
        user_prompt_template=user_prompt,
        model="gpt-4o-mini",  # 使用便宜的模型进行测试
        temperature=0,
        max_tokens=1000
    )
    
    print("\n处理结果:")
    print(json.dumps(result, ensure_ascii=False, indent=2))

