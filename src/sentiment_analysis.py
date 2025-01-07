"""
文件功能：对excel中的评论进行情感分析，并输出情感得分
作者：huasheng
日期：2025-01-04
"""

import pandas as pd
from snownlp import SnowNLP
import warnings
warnings.filterwarnings('ignore')

def analyze_sentiment(text):
    """
    对输入的文本进行情感分析，返回情感得分
    得分范围：0-1，越接近1表示情感越正面
    """
    try:
        # 处理空值情况
        if pd.isna(text):
            return None
        
        s = SnowNLP(str(text))
        return s.sentiments
    except Exception as e:
        print(f"处理文本时出错: {text}")
        print(f"错误信息: {str(e)}")
        return None

def process_excel(input_file, comment_column, output_file=None):
    """
    处理Excel文件中的评论数据
    
    参数:
    input_file: 输入的Excel文件路径
    comment_column: 包含评论的列名
    output_file: 输出的Excel文件路径（可选）
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        
        # 检查评论列是否存在
        if comment_column not in df.columns:
            raise ValueError(f"未找到列名 '{comment_column}'")
        
        # 对评论进行情感分析
        df['情感得分'] = df[comment_column].apply(analyze_sentiment)
        
        # 如果指定了输出文件，则保存结果
        if output_file:
            df.to_excel(output_file, index=False)
            print(f"结果已保存至: {output_file}")
        
        return df
        
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return None

if __name__ == "__main__":
    # 使用示例
    input_file = "../处理后的评论汇总.xlsx"  # 输入文件名
    comment_column = "评论内容"    # 评论列的列名
    output_file = "comments_with_sentiment.xlsx"  # 输出文件名
    
    result = process_excel(input_file, comment_column, output_file)
    
    if result is not None:
        print("\n数据分析完成！")
        print(f"总计处理评论数: {len(result)}")
        print(f"平均情感得分: {result['情感得分'].mean():.3f}") 