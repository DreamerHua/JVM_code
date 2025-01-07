"""
主要功能：处理评论汇总文件，添加新的属性列
输入：汇总的评论文件，文件格式为xlsx。该文件放置在代码所在目录下
输出：处理后的评论汇总文件，文件格式为xlsx。输出位置为代码所在目录下
"""

import re
import pandas as pd


def add_video_id(df):
    """
    为数据框添加宣传片ID列，通过对宣传片内容进行枚举映射实现
    
    参数:
    df: 包含宣传片内容列的DataFrame
    
    返回:
    DataFrame: 添加宣传片ID列后的数据框
    """
    unique_contents = df['宣传片内容'].unique()
    content_to_id = {content: idx + 1 for idx, content in enumerate(unique_contents)}
    df['宣传片ID'] = df['宣传片内容'].map(content_to_id)

    # 打印映射关系供参考
    for spot_id, spot_name in enumerate(df['宣传片内容'].unique(), start=1):
        print(f"  宣传片ID {spot_id}: {spot_name}")
    print(f"宣传片数量：{len(content_to_id)}")
    return df


def add_ai_generated_flag(df):
    """
    为数据框添加是否AI生成列
    
    参数:
    df: 包含宣传片内容列的DataFrame
    
    返回:
    DataFrame: 添加是否AI生成列后的数据框
    """
    def determine_video_type(content):
        if 'AI生成' in str(content):
            return 1
        elif '人生成' in str(content):
            return 0
        else:
            return None
    
    df['是否AI生成'] = df['宣传片内容'].apply(determine_video_type)
    return df


def add_video_metadata(df, metadata_dict):
    """
    为数据框添加视频元数据信息(发布时间、视频链接、景区所在地和景区类型)
    
    参数:
    df: 包含宣传片内容列的DataFrame
    metadata_dict: 字典，key为宣传片内容，value为包含元数据的元组
                  例如: {'景区A': ('2024-01-01', 'http://example.com/video1', '江苏', '自然景观')}
    
    返回:
    DataFrame: 添加视频发布时间、视频链接、景区所在地和景区类型列后的数据框
    """
    # 添加新列并设置默认值
    df['视频发布时间'] = None
    df['视频链接'] = None
    df['景区所在地'] = None  # 新增景区所在地列
    df['景区类型'] = None    # 新增景区类型列
    
    # 根据宣传片内容映射元数据
    for content, (publish_time, video_url, location, spot_type) in metadata_dict.items():
        mask = df['宣传片内容'] == content
        df.loc[mask, '视频发布时间'] = publish_time
        df.loc[mask, '视频链接'] = video_url
        df.loc[mask, '景区所在地'] = location  # 设置景区所在地
        df.loc[mask, '景区类型'] = spot_type   # 设置景区类型
        
    return df


def get_video_metadata():
    """
    维护视频元数据信息的函数
    
    返回:
    dict: 包含视频元数据的字典，结构为:
        {
            '宣传片内容': (视频发布时间, 视频链接, 景区所在地, 景区类型)
        }
    """
    # 手动输入视频元数据信息。
    metadata = {
        '《HYPER AI》小红书个体号-云南风光-AI生成': ('2024-03-29', 'https://www.xiaohongshu.com/explore/660675a0000000001a00e578?xsec_token=AB2Rkh2xSAGzwL8AZ_qY-sLKDLPCrvIQuD9iK-ECWXBcU=&xsec_source=pc_user', '云南', '自然景观'),
        '《Longhoo文旅》城市宣传号-南京风光-AI生成': ('2024-04-10', 'https://www.xiaohongshu.com/explore/6616477b000000001b008dd8?xsec_token=ABakRUPmIxE7OCKpxF7ErQ4jlVuSe64aFtPi6j2SYmyMA=&xsec_source=pc_search&source=web_search_result_notes', '江苏', '自然景观'),
        '《凌凌张～》小红书个体号-云南风光-人生成': ('2024-11-02', 'https://www.xiaohongshu.com/explore/67258b4b0000000019014ddf?xsec_token=ABVfHmXpUpgIcIdHWZ9wHUVbge9sRIkeOn9LVgsPPWbso=&xsec_source=pc_search&source=web_search_result_notes', '云南', '自然景观'),
        '《哩好南京HOKU》城市宣传号-南京风光-人生成': ('2024-07-16', 'https://www.xiaohongshu.com/explore/6695ec040000000025016881?xsec_token=ABjsy_R8olVtQaD-pm7cwwLrW19YF-blruSRurXD_3G0Q=&xsec_source=pc_search&source=web_explore_feed', '江苏', '自然景观'),
        '《山西省文化和旅游厅》小红书官号-城市建筑宣传-AI生成': ('2023-05-16', 'https://www.xiaohongshu.com/explore/64634f8700000000110133b1?note_flow_source=wechat&xsec_token=CBpqJPvwsmZEhks6nHxwQGBgf15AWm9Cw4x0Z74at1eFQ=', '山西', '人文景观'),
        '《山西省文化和旅游厅》小红书官号-城市建筑宣传-人生成': ('2024-11-02', 'https://www.xiaohongshu.com/explore/6724913400000000190179dc?note_flow_source=wechat&xsec_token=CBGhC3tqbKmq9-Jy8zN6PxeNwL_dGeMpFB2lZEGzl9k0M=', '山西', '人文景观'),
        '《山西省文化和旅游厅》小红书官号-城市文化宣传-AI生成': ('2024-08-22', 'https://www.xiaohongshu.com/explore/66c6d679000000001d038084?note_flow_source=wechat&xsec_token=CBKuJafhCZCrjGJY0f2bPN39tUEe5g_8wAjM8M8LFaiFE=', '山西', '人文景观'),
        '《山西省文化和旅游厅》小红书官号-城市文化宣传-人生成': ('2024-10-12', 'https://www.xiaohongshu.com/explore/670a36b200000000240144a1?xsec_token=ABbidCEp2FeN4RCB641kazCY6S3Uw9fR9KXUHEg3LTlu0=&xsec_source=pc_user', '山西', '人文景观'),
        '《本溪文旅》抖音官号-城市风光宣传-AI生成': ('2024-07-11', 'https://www.douyin.com/user/MS4wLjABAAAASTbxU0XV3jZgW_bXCseyDaWMPmcLpyAmT6_rYnN6lyU?from_tab_name=main&modal_id=7390364835101330703&relation=0&vid=7438961235174903074', '辽宁', '自然景观'),
        '《本溪文旅》抖音官号-城市风光宣传-人生成': ('2024-11-19', 'https://www.douyin.com/video/7438961235174903074?modeFrom=userPost&secUid=MS4wLjABAAAASTbxU0XV3jZgW_bXCseyDaWMPmcLpyAmT6_rYnN6lyU', '辽宁', '自然景观')
    }
    
    return metadata


def add_location_id(df):
    """
    为数据框添加景区所在地ID列，通过对景区所在地进行枚举映射实现
    
    参数:
    df: 包含景区所在地列的DataFrame
    
    返回:
    DataFrame: 添加景区所在地ID列后的数据框
    """
    # 获取所有唯一的景区所在地
    unique_locations = df['景区所在地'].unique()
    
    # 创建映射字典,从1开始编号
    location_to_id = {location: idx for idx, location in enumerate(unique_locations, start=1)}
    
    # 添加新列并根据映射填充ID
    df['景区所在地ID'] = df['景区所在地'].map(location_to_id)
    
    # 打印映射关系供参考
    print("\n景区所在地ID映射关系：")
    for location, loc_id in location_to_id.items():
        print(f"  所在地ID {loc_id}: {location}")
        
    return df


def add_spot_type_id(df):
    """
    为数据框添加景区类型ID列，通过对景区类型进行枚举映射实现
    
    参数:
    df: 包含景区类型列的DataFrame
    
    返回:
    DataFrame: 添加景区类型ID列后的数据框
    """
    # 获取所有唯一的景区类型
    unique_types = df['景区类型'].unique()
    
    # 创建映射字典,从1开始编号
    type_to_id = {spot_type: idx for idx, spot_type in enumerate(unique_types, start=1)}
    
    # 添加新列并根据映射填充ID
    df['景区类型ID'] = df['景区类型'].map(type_to_id)
    
    # 打印映射关系供参考
    print("\n景区类型ID映射关系：")
    for spot_type, type_id in type_to_id.items():
        print(f"  类型ID {type_id}: {spot_type}")
        
    return df


def add_time_diff(df):
    """
    计算评论时间与视频发布时间的时间差(天数)
    
    参数:
    df: 包含评论时间和视频发布时间列的DataFrame
    
    返回:
    DataFrame: 添加评论时间差列后的数据框
    """
    try:
        # 将日期字符串转换为datetime对象
        df['评论时间'] = pd.to_datetime(df['评论时间'])
        df['视频发布时间'] = pd.to_datetime(df['视频发布时间'])
        
        # 计算时间差(天数)
        df['评论时间差'] = (df['评论时间'] - df['视频发布时间']).dt.days + 1
        
        # 检查异常情况(评论时间早于发布时间)
        mask = df['评论时间差'] <= 0
        df.loc[mask, '评论时间差'] = -1
        
        # 将日期重新转换回字符串格式
        df['评论时间'] = df['评论时间'].dt.strftime('%Y-%m-%d')
        df['视频发布时间'] = df['视频发布时间'].dt.strftime('%Y-%m-%d')
        
        print("\n评论时间差统计:")
        print(f"  异常评论数(评论时间早于发布时间): {mask.sum()}")
        print(f"  正常评论的平均时间差: {df.loc[~mask, '评论时间差'].mean():.1f}天")
        
        return df
        
    except Exception as e:
        print(f"计算时间差时出错: {str(e)}")
        return df


def add_comment_length(df):
    """
    计算去除表情后的评论字数、表情认为是1个字符的评论字数
    
    参数:
    df: 包含评论内容列的DataFrame
    
    返回:
    DataFrame: 添加评论字数列后的数据框
    """
    def count_text_length(text):
        # 使用正则表达式去除表情符号（格式为[xxx]或[xxxR]）
        clean_text = re.sub(r'\[[^\]]+\]', '', str(text))
        # 去除其他特殊字符，只保留文字和空格
        clean_text = re.sub(r'[^\w\s]', '', clean_text)
        return len(clean_text.strip())
    
    def count_text_length_with_emojis(text):
        # 使用正则表达式匹配表情符号（格式为[xxx]或[xxxR]）
        clean_text = re.sub(r'\[[^\]]+\]', 'a', str(text))  # 将表情符号替换为'a'，占1个字符
        # 去除其他特殊字符，只保留文字和空格
        clean_text = re.sub(r'[^\w\s]', '', clean_text)
        return len(clean_text.strip())
    
    df['评论字数'] = df['评论内容'].apply(count_text_length)
    df['评论字数(加表情)'] = df['评论内容'].apply(count_text_length_with_emojis)
    return df


def add_ip_address(df, ip_address_file):
    """
    从IP地址文件读取信息并拼接到原始数据框
    
    参数:
    df: DataFrame - 原始数据框
    
    返回:
    DataFrame: 添加IP地址列后的数据框
    """
    try:
        # 读取IP地址文件
        ip_df = pd.read_excel(ip_address_file)
        
        # 确保两个数据框有相同的长度用于合并
        if len(df) != len(ip_df):
            print("数据框长度不一致，无法合并")
            return df
        df['IP地址'] = ip_df['IP地址']
        
        print(f"IP地址添加完成，共添加 {ip_df.shape[0]} 条记录")
        return df
        
    except FileNotFoundError:
        print("未找到IP地址文件(ip_addresses.xlsx)，跳过IP地址添加")
        df['IP地址'] = None
        return df
    except Exception as e:
        print(f"添加IP地址时出错: {str(e)}")
        df['IP地址'] = None
        return df


def add_local_comment_flag(df):
    """
    添加是否为本地评论的标记
    如果IP地址与景区所在地一致，标记为1；否则标记为0
    
    参数:
    df: DataFrame - 包含IP地址和景区所在地列的数据框
    
    返回:
    DataFrame: 添加是否本地评论标记后的数据框
    """
    # 添加新列，默认值为0
    df['是否本地评论'] = 0
    
    # 确保数据清洗：去除字符串前后空格
    df['IP地址'] = df['IP地址'].astype(str).str.strip()
    df['景区所在地'] = df['景区所在地'].astype(str).str.strip()
    
    # 如果IP地址包含景区所在地信息，则标记为本地评论
    mask = df.apply(lambda row: row['景区所在地'] in row['IP地址'], axis=1)
    df.loc[mask, '是否本地评论'] = 1
    
    # 打印统计信息
    local_count = df['是否本地评论'].sum()
    total_count = len(df)
    print(f"\n本地评论统计:")
    print(f"  本地评论数: {local_count}")
    print(f"  本地评论占比: {(local_count/total_count*100):.1f}%")
    
    return df


def process_comments_data(input_file, output_file, ip_address_file):
    """
    处理评论汇总文件，添加新的属性列
    
    参数:
    input_file: 输入的Excel文件路径
    output_file: 输出的Excel文件路径
    
    返回:
    DataFrame: 处理后的数据框，包含新增的属性列
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(input_file)
        
        # 1. 添加视频元数据信息
        metadata_dict = get_video_metadata()
        df = add_video_metadata(df, metadata_dict)

        # 2. 处理所有需要枚举映射到ID的列
        df = add_video_id(df)  # 宣传片ID映射
        df = add_location_id(df)  # 景区所在地ID映射
        df = add_spot_type_id(df)  # 景区类型ID映射
        
        # 3. 拼接缓存好的IP地址
        df = add_ip_address(df, ip_address_file)

        # 4. 处理其他属性列
        df = add_ai_generated_flag(df)  # 是否AI生成标记
        df = add_time_diff(df)  # 计算评论时间差
        df = add_comment_length(df)  # 计算去除表情后的评论字数
        df = add_local_comment_flag(df)  # 是否是本地人评论

        # 保存处理后的结果
        df.to_excel(output_file, index=False)
        
        print(f"数据处理完成！")
        print(f"总评论数：{len(df)}")
        print(f"结果已保存至：{output_file}")
        
        return df
        
    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return None


if __name__ == "__main__":
    raw_comments_file = "所有评论汇总.xlsx"
    processed_comments_file = "处理后的评论汇总.xlsx"
    processed_df = process_comments_data(raw_comments_file, processed_comments_file)
    print(f"处理后的评论汇总已保存至：{processed_comments_file}")

