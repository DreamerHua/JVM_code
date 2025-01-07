import hanlp.pretrained
import pandas as pd
import torch
from transformers import AutoModelForSequenceClassification, AutoTokenizer
# from paddlenlp import Taskflow
from snownlp import SnowNLP
import hanlp
import emoji
import re
from tqdm import tqdm

class SentimentAnalyzer:
    """情感分析器类，整合多个模型"""
    
    def __init__(self):
        self.models = {}
        self.results = {}
    
    
    def preprocess_text(self, text):
        """预处理文本"""
        if pd.isna(text):
            return ""
        
        text = str(text)
        # 保留emoji表情和基本字符
        text = ''.join([char for char in text if char in emoji.EMOJI_DATA or char.isalnum() or char.isspace()])
        # 处理连续标点
        text = re.sub(r'[!！]{2,}', '！', text)
        text = re.sub(r'[?？]{2,}', '？', text)
        return text.strip()
    
    
    def init_all_models(self):
        """初始化所有模型"""
        print("正在初始化模型...")
        
        # 1. 微博情感模型
        # print("初始化微博模型...")
        # self.init_weibo_model()   # TODO: FIX HERE
        
        # 2. BERT-WWM模型
        print("初始化BERT-WWM模型...")
        self.init_bert_wwm_model()
        
        # 3. SKEP模型
        # print("初始化SKEP模型...")
        # self.init_skep_model()   # TODO: FIX HERE
        
        # 4. PaddleNLP情感分析模型
        # print("初始化PaddleNLP模型...")
        # self.init_paddle_model()   # TODO: FIX HERE
        
        # 5. HanLP情感分析模型
        print("初始化HanLP模型...")
        self.init_hanlp_model()
        
        # 6. SnowNLP模型
        # SnowNLP不需要预初始化
        
        print("所有模型初始化完成！")
    
    
    def init_weibo_model(self):
        """初始化微博情感分析模型"""
        model_name = "uer/roberta-base-finetuned-weibo-sentiment"
        tokenizer = AutoTokenizer.from_pretrained(model_name)
        model = AutoModelForSequenceClassification.from_pretrained(model_name)
        if torch.cuda.is_available():
            model = model.cuda()
        self.models['weibo'] = (model, tokenizer)
    
    
    def init_bert_wwm_model(self):
        """初始化BERT-WWM模型"""
        model_name = "hfl/chinese-bert-wwm-ext"
        tokenizer = AutoTokenizer.from_pretrained(model_name)
        model = AutoModelForSequenceClassification.from_pretrained(model_name)
        if torch.cuda.is_available():
            model = model.cuda()
        self.models['bert_wwm'] = (model, tokenizer)
    
    
    # def init_skep_model(self):
    #     """初始化SKEP模型"""
    #     self.models['skep'] = Taskflow("sentiment_analysis", model="skep_ernie_1.0_large_ch")
    
    
    # def init_paddle_model(self):
    #     """初始化PaddleNLP模型"""
    #     self.models['paddle'] = Taskflow('sentiment_analysis', model='finetune-sentiment')
    
    
    def init_hanlp_model(self):
        """初始化HanLP模型"""
        try:
            # 使用更简单的情感分析模型
            self.models['hanlp'] = hanlp.load('CHNSENTICORP_BERT_BASE_ZH')
        except:
            try:
                # 备选方案：使用 HanLP 2.1 内置的情感分析模型
                HanLP = hanlp.pipeline().append(hanlp.utils.rules.split_sentence, output_key='sentences')\
                                      .append(hanlp.load('CHNSENTICORP_ALBERT_BASE'))
                self.models['hanlp'] = HanLP
            except Exception as e:
                print(f"HanLP模型加载失败: {str(e)}")
                print("将跳过HanLP模型的情感分析")
                # self.models['hanlp'] = None
    
    
    def analyze_with_transformer(self, text, model_key):
        """使用Transformer模型进行分析"""
        try:
            model, tokenizer = self.models[model_key]
            inputs = tokenizer(
                self.preprocess_text(text),
                return_tensors="pt",
                padding=True,
                truncation=True,
                max_length=512
            )
            
            if torch.cuda.is_available():
                inputs = {k: v.cuda() for k, v in inputs.items()}
            
            with torch.no_grad():
                outputs = model(**inputs)
            
            probs = torch.nn.functional.softmax(outputs.logits, dim=-1)
            return probs[0][1].item()
            
        except Exception as e:
            print(f"{model_key}处理文本时出错: {text}")
            print(f"错误信息: {str(e)}")
            return None
    
    
    def analyze_with_skep(self, text):
        """使用SKEP模型进行分析"""
        try:
            result = self.models['skep'](self.preprocess_text(text))
            return result[0]['score']
        except Exception as e:
            print(f"SKEP处理文本时出错: {text}")
            print(f"错误信息: {str(e)}")
            return None
    
    
    def analyze_with_paddle(self, text):
        """使用PaddleNLP模型进行分析"""
        try:
            result = self.models['paddle'](self.preprocess_text(text))
            return result[0]['probability']
        except Exception as e:
            print(f"PaddleNLP处理文本时出错: {text}")
            print(f"错误信息: {str(e)}")
            return None
    
    
    def analyze_with_hanlp(self, text):
        """使用HanLP模型进行分析"""
        try:
            if self.models['hanlp'] is None:
                return None
            
            result = self.models['hanlp'](self.preprocess_text(text))
            
            # 处理不同类型的返回结果
            if isinstance(result, list):
                if isinstance(result[0], tuple):
                    # 处理 (label, score) 格式
                    label, score = result[0]
                    return score if label == 'positive' else 1 - score
                else:
                    # 处理句子列表格式
                    return sum(r['positive'] for r in result) / len(result)
            return result['positive'] if 'positive' in result else 0.5
            
        except Exception as e:
            print(f"HanLP处理文本时出错: {text}")
            print(f"错误信息: {str(e)}")
            return None
    
    
    def analyze_with_snownlp(self, text):
        """使用SnowNLP进行分析"""
        try:
            s = SnowNLP(self.preprocess_text(text))
            return s.sentiments
        except Exception as e:
            print(f"SnowNLP处理文本时出错: {text}")
            print(f"错误信息: {str(e)}")
            return None
    
    
    def analyze_text(self, text):
        """使用所有模型分析文本"""
        results = {'评论内容': text}
        
        if 'weibo' in self.models:
            results['微博模型'] = self.analyze_with_transformer(text, 'weibo')
        
        if 'bert_wwm' in self.models:
            results['BERT-WWM'] = self.analyze_with_transformer(text, 'bert_wwm')
        
        if 'skep' in self.models:
            results['SKEP'] = self.analyze_with_skep(text)
        
        if 'paddle' in self.models:
            results['PaddleNLP'] = self.analyze_with_paddle(text)
        
        if 'hanlp' in self.models:
            results['HanLP'] = self.analyze_with_hanlp(text)
        
        # SnowNLP不需要预加载模型
        results['SnowNLP'] = self.analyze_with_snownlp(text)
        
        return results


def compare_models(input_file, text_column='评论内容', sample_size=None):
    """
    比较多个模型的情感分析结果
    
    参数:
    input_file: 输入文件路径
    text_column: 文本评论对应的列名
    sample_size: 采样数量，如果不指定则处理所有数据
    """
    # 读取数据
    df = pd.read_excel(input_file)
    if sample_size:
        df = df.sample(n=min(sample_size, len(df)))
    
    # 初始化分析器
    analyzer = SentimentAnalyzer()
    analyzer.init_all_models()
    
    # 分析文本
    print("开始分析文本...")
    results = []
    for text in tqdm(df[text_column]):
        result = analyzer.analyze_text(text)
        results.append(result)
    
    # 转换结果为DataFrame
    results_df = pd.DataFrame(results)
    
    # 计算统计信息
    model_columns = list(results[0].keys()) if results else []
    stats = results_df[model_columns].agg(['mean', 'std', 'min', 'max'])
    
    # 保存结果
    output_file = '模型对比结果.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        results_df.to_excel(writer, sheet_name='详细结果', index=False)
        stats.to_excel(writer, sheet_name='统计信息')
    
    print(f"\n结果已保存至: {output_file}")
    print("\n模型统计信息:")
    print(stats)
    
    return results_df, stats


if __name__ == "__main__":
    # 使用示例
    input_file = "处理后的评论汇总2.xlsx"
    results, stats = compare_models(
        input_file=input_file,
        text_column='评论内容',
        sample_size=100  # 可以先用100条数据测试
    )