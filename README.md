# 评论数据分析工具

## 项目简介
这是一个用于处理和分析视频评论数据的综合工具，主要功能包括评论数据提取、数据处理和情感分析。该工具专门设计用于处理旅游景点宣传视频的评论数据，支持多维度分析。

## 项目架构 

## 模块功能说明

### 1. 评论提取模块 (extract_comments.py)
- 从多个Excel文件中提取评论数据
- 识别主评论和子评论
- 统一处理评论格式

### 2. 评论处理模块 (process_comments.py)
- 添加视频元数据信息
- 生成各种ID映射
- 计算评论特征（字数、时间差等）
- 判断评论属性（本地评论、视频是否AI生成等）

### 3. 情感分析模块 (sentiment_analysis.py)
- 评论情感倾向分析
- 情感得分计算
- 结果统计和输出

### 4. 模型对比模块 (sentiment_analysis_compare.py)
- 多个情感分析模型的对比
- 模型性能统计
- 结果可视化


## 依赖包 
- 见requirements.txt

## 使用方法
1. 安装依赖包：
    ```bash
    pip install -r requirements.txt
    ```

2. 配置参数：
- 修改 `config/config.yaml` 文件中的相关配置
- 设置输入输出路径
- 配置模型参数

3. 运行程序：
- 在命令行中运行 `python main.py` 即可执行整个流程
- 根据需要选择是否运行模型对比

## 输出说明
程序会依次生成以下文件：
- 所有评论汇总.xlsx：原始评论数据汇总
- 处理后的评论汇总.xlsx：添加特征后的评论数据
- comments_with_sentiment.xlsx：情感分析结果
- 模型对比结果.xlsx：（可选）多模型分析结果

## 注意事项
1. 确保输入Excel文件格式符合要求
2. 运行前检查配置文件config.yaml中的路径设置
3. 如果run_model_comparison配置为true，先使用小规模数据测试

## TODO List
- [ ] 修复微博情感模型
- [ ] 修复SKEP模型
- [ ] 修复PaddleNLP模型

