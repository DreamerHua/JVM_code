"""
工具函数模块：提供文本预处理、表情符号处理和格式化打印等通用功能
"""

import re
import emoji
import pandas as pd

def preprocess_text(text):
    """预处理文本内容"""
    if pd.isna(text):
        return ""
    
    text = str(text)
    # 保留emoji表情和基本字符
    text = ''.join([char for char in text if char in emoji.EMOJI_DATA or char.isalnum() or char.isspace()])
    # 处理连续标点
    text = re.sub(r'[!！]{2,}', '！', text)
    text = re.sub(r'[?？]{2,}', '？', text)
    return text.strip()


def clean_emoji(text):
    """清理文本中的表情符号"""
    if pd.isna(text):
        return ""
    clean_text = re.sub(r'\[[^\]]+\]', '', str(text))
    clean_text = re.sub(r'[^\w\s]', '', clean_text)
    return clean_text.strip()


def format_print(message, is_title=False):
    """格式化打印信息"""
    if is_title:
        print(f"\n{'='*20} {message} {'='*20}")
    else:
        print(f">>> {message}") 