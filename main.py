import os
import yaml
from src.extract_comments import process_folder
from src.process_comments import process_comments_data
from src.sentiment_analysis import process_excel
from src.sentiment_analysis_compare import compare_models

def load_config():
    """加载配置文件"""
    with open('config/config.yaml', 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

def main():
    """主程序入口"""
    # 加载配置
    config = load_config()
    
    print("=== 开始评论数据分析流程 ===")
    
    # 步骤1: 提取评论
    print("\n[步骤1] 提取并汇总评论数据...")
    raw_comments = process_folder(config['input_folder'])
    if raw_comments is not None:
        raw_comments_file = config['raw_comments_file']
        raw_comments.to_excel(raw_comments_file, index=False)
        print(f"原始评论已保存至：{raw_comments_file}")
    else:
        print("评论提取失败，程序终止")
        return
    
    # 步骤2: 处理评论数据
    print("\n[步骤2] 处理评论数据...")
    processed_df = process_comments_data(
        raw_comments_file,
        config['processed_comments_file']
    )
    if processed_df is None:
        print("评论处理失败，程序终止")
        return
    
    # 步骤3: 情感分析
    print("\n[步骤3] 进行情感分析...")
    sentiment_result = process_excel(
        config['processed_comments_file'],
        "评论内容",
        config['sentiment_output_file']
    )
    
    # 步骤4: 模型对比（可选）
    if config.get('run_model_comparison', False):
        print("\n[步骤4] 进行模型对比分析...")
        results, stats = compare_models(
            config['processed_comments_file'],
            text_column='评论内容',
            sample_size=config.get('comparison_sample_size', 100)
        )
    
    print("\n=== 所有处理完成 ===")

if __name__ == "__main__":
    main() 