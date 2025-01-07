# extract_comments.py
# 文件功能：提取excel文件中的评论，并输出评论内容、一级评论ID/评论类型和宣传片内容
# 作者：huasheng
# 日期：2025-01-04

import pandas as pd
import os
import re

def extract_comments(excel_path):
    """
    从Excel文件中提取评论内容、一级评论ID/评论类型和宣传片内容
    
    参数:
    excel_path: Excel文件的路径，文件名将被用作宣传片内容
    
    返回：
    DataFrame: 包含以下列的数据框:
        - 评论内容: 评论的具体内容
        - 一级评论ID/评论类型: 根据Excel文件列名自动选择
        - 宣传片内容: 从Excel文件名中提取
        
    异常:
    ValueError: 当Excel文件缺少必要的列组合时抛出
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(
            excel_path,
            dtype={'评论时间': 'object'}  # 先将日期列读取为对象类型
        )
        df['评论时间'] = pd.to_datetime(df['评论时间'], errors='coerce')
        df['评论时间'] = df['评论时间'].dt.strftime('%Y-%m-%d')  # 格式化为年-月-日

        # 定义两种可能的列名组合
        column_sets = [
            ['评论内容', '一级评论ID', '评论时间', 'IP地址'],
            ['评论内容', '评论类型', '评论时间', 'IP地址']
        ]
        
        # 检查是否存在任一组列名
        found_columns = None
        for cols in column_sets:
            if all(col in df.columns for col in cols):
                found_columns = cols
                break
                
        if found_columns is None:
            raise ValueError("Excel文件中缺少必要的列组合：\n"
                           "需要 '评论内容'、'一级评论ID'(或'评论类型')、'评论时间'和'IP地址'")
        
        # 提取所需列
        result_df = df[found_columns].copy()
        
        # 删除空值行
        result_df = result_df.dropna(subset=['评论内容'])
        
        # 从文件路径中提取文件名(不含扩展名)
        file_name = excel_path.split('\\')[-1]  # 获取文件名
        file_name_without_ext = file_name.rsplit('.', 1)[0]  # 移除扩展名
        
        # 添加宣传片内容列
        result_df['宣传片内容'] = file_name_without_ext
        
        return result_df
        
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return None


def convert_excel_date(df, date_column='评论时间'):
    """
    将Excel日期序列号转换为正确的日期格式
    
    参数:
    df: DataFrame
    date_column: 需要转换的日期列名
    
    返回:
    DataFrame: 日期格式已转换的数据框
    """
    try:
        # 将Excel序列号转换为datetime格式
        df[date_column] = pd.to_datetime('1899-12-30') + pd.to_timedelta(df[date_column], 'D')
        
        # 如果需要特定的显示格式，可以使用strftime
        df[date_column] = df[date_column].dt.strftime('%Y-%m-%d')
        
        return df
    except Exception as e:
        print(f"转换日期格式时出错: {str(e)}")
        return df
    

def add_main_comment_flag(df):
    """
    为数据框添加是否主评论列
    
    参数:
    df: 包含评论相关列的DataFrame
    
    返回:
    DataFrame: 添加是否主评论列后的数据框
    """
    def is_main_comment(row):
        if '一级评论ID' in df.columns:
            comment_id = str(row['一级评论ID'])
            hex_pattern = re.compile(r'^[0-9a-fA-F]+$')
            if hex_pattern.match(comment_id):
                return 0
        
        if '评论类型' in df.columns:
            comment_type = str(row['评论类型'])
            if comment_type == '子评论':
                return 0
        
        return 1
    
    df['是否主评论'] = df.apply(is_main_comment, axis=1)

    # 删除一级评论ID列、评论类型列
    if '一级评论ID' in df.columns:
        df = df.drop(columns=['一级评论ID'])
    if '评论类型' in df.columns:
        df = df.drop(columns=['评论类型'])

    return df


def process_folder(folder_path):
    """
    处理文件夹中的所有Excel文件并合并结果
    
    参数:
    folder_path: 包含Excel文件的文件夹路径
    
    返回:
    DataFrame: 合并后的数据框，包含所有文件的评论数据
        - 如果成功处理至少一个文件，返回合并后的DataFrame
        - 如果没有成功处理任何文件，返回None
    """
    # 存储所有数据框的列表
    all_dataframes = []
    
    # 获取文件夹中所有Excel文件
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    # 处理每个Excel文件
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        print(f"\n处理文件: {file}")
        df = extract_comments(file_path)
        # df = convert_excel_date(df, date_column='评论时间')
        if df is not None:
            df = add_main_comment_flag(df) # 添加是否主评论列
            all_dataframes.append(df)
            print(f"成功提取评论数：{len(df)}")
    
    # 合并所有数据框
    if all_dataframes:
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        print(f"\n所有文件处理完成！")
        print(f"总评论数：{len(combined_df)}")
        return combined_df
    else:
        print("没有成功处理任何文件")
        return None
    


if __name__ == "__main__":
    # 使用示例
    folder_path = "D:\\下载new\\出国申请材料\\12-AI论文\\数据采集"  # 替换为你的文件夹路径
    
    # 处理文件夹中的所有Excel文件
    combined_comments = process_folder(folder_path)
    
    if combined_comments is not None:
        # 保存合并后的结果
        comments_file = "所有评论汇总.xlsx"
        combined_comments.to_excel(comments_file, index=False)
        print(f"\n汇总结果已保存至：{comments_file}")
