#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文选题查重系统
读取"原始数据"文件夹中的Excel文件，查重"课题名称"，输出到"查重结果"文件夹
"""

import os
import sys
import glob
import re
from pathlib import Path
from datetime import datetime
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import jieba

# 设置输出编码为UTF-8
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class ThesisChecker:
    def __init__(self, data_dir="原始数据", output_dir="查重结果", threshold=0.75):
        self.data_dir = Path(data_dir)
        self.output_dir = Path(output_dir)
        self.threshold = threshold
        self.output_dir.mkdir(exist_ok=True)

    def extract_class_from_filename(self, filename):
        """从文件名中提取班级名称，班级名必须包含数字"""
        # 移除扩展名
        name = Path(filename).stem

        # 优先匹配标准班级格式：汽服2401B、新能2401D、汽服2401ZB等
        # 格式：专业简称(汽服/新能) + 年份(2位) + 班号(2位) + 类型(B/D/ZB等)
        standard_pattern = r'((?:汽服|新能)\d{4}[A-Z]+)'
        match = re.search(standard_pattern, name)
        if match:
            return match.group(1)

        # 备选模式：其他可能的班级名格式（必须包含数字）
        patterns = [
            r'(\d+班)',  # 1班, 2021班
            r'(班\s*\d+)',  # 班1, 班 2021
            r'([^\d\s]*\d+[^\d\s]*班)',  # 计算机1班, 软工2021班
            r'(班级[^\d\s]*\d+)',  # 班级1, 班级2021
            r'(\d+\s*级)',  # 2021级, 2021 级
        ]

        for pattern in patterns:
            match = re.search(pattern, name)
            if match:
                class_name = match.group(1).strip()
                # 确保包含数字
                if re.search(r'\d', class_name):
                    return class_name

        return ''  # 未找到班级名时返回空字符串

    def read_excel_files(self):
        """读取原始数据文件夹中的所有Excel文件"""
        excel_files = list(self.data_dir.glob("*.xls")) + list(self.data_dir.glob("*.xlsx"))
        if not excel_files:
            raise FileNotFoundError(f"在 {self.data_dir} 文件夹中未找到Excel文件")

        all_data = []
        for file in excel_files:
            try:
                # 尝试不同的header位置
                df = None
                for header_row in [None, 0, 1, 2]:
                    try:
                        df = pd.read_excel(file, header=header_row)
                        # 检查是否包含关键列（同时检查课题名称和学生姓名）
                        cols = df.columns.tolist()
                        has_title = any('课题' in str(c) or '题目' in str(c) or '标题' in str(c) for c in cols)
                        has_student = any(c == '学生姓名' for c in cols)
                        if has_title and has_student:
                            break
                    except:
                        continue

                if df is not None and len(df) > 0:
                    df['source_file'] = file.name
                    df['source_class'] = self.extract_class_from_filename(file.name)
                    all_data.append(df)
                    class_info = f" [班级: {df['source_class'].iloc[0]}]" if df['source_class'].iloc[0] else ""
                    print(f"✓ 已读取: {file.name}{class_info} ({len(df)} 条记录)")
            except Exception as e:
                print(f"✗ 读取失败: {file.name} - {e}")

        if not all_data:
            raise ValueError("未能成功读取任何Excel文件")

        # 合并所有数据
        combined_df = pd.concat(all_data, ignore_index=True)
        print(f"\n总共读取 {len(combined_df)} 条记录\n")
        return combined_df

    def is_valid_title(self, title):
        """判断题目是否有效（非空、非待定等无意义内容）"""
        if pd.isna(title) or title == '' or str(title).strip() == '':
            return False

        title_str = str(title).strip()

        # 检查无意义的关键词（使用词边界匹配，避免误匹配正常词汇）
        # 通过正则表达式确保匹配的是完整词语或独立状态
        import re

        # 定义无效模式：完全匹配、前后有空格或特殊符号、纯英文/数字状态
        invalid_patterns = [
            r'^待定$',          # 完全匹配"待定"
            r'^暂定$',          # 完全匹配"暂定"
            r'^待填$',          # 完全匹配"待填"
            r'^\s*无\s*$',      # 仅包含"无"
            r'^空题?$',         # "空"或"空题"
            r'^待选$',          # 完全匹配"待选"
            r'^未定$',          # 完全匹配"未定"
            r'^nan$',           # 仅"nan"
            r'^None$',          # 仅"None"
            r'^入伍$',          # 入伍（已参军）
        ]

        for pattern in invalid_patterns:
            if re.search(pattern, title_str, re.IGNORECASE):
                return False

        # 额外检查：去除书名号《》后，如果内容过短也可能是无效
        content = title_str.replace('《', '').replace('》', '').strip()
        if len(content) <= 2:
            return False

        # 检查是否为纯符号或无意义字符
        if not re.search(r'[\u4e00-\u9fa5a-zA-Z]', content):  # 不包含中文或英文
            return False

        return True

    def normalize_columns(self, df):
        """标准化列名，提取所需信息"""
        # 查找课题名称列
        title_col = None
        for col in df.columns:
            if '课题' in str(col):
                title_col = col
                break

        # 查找学生姓名列（精确匹配"学生姓名"）
        student_col = None
        for col in df.columns:
            if col == '学生姓名':
                student_col = col
                break

        # 查找指导教师列（精确匹配"指导教师姓名"）
        teacher_col = None
        for col in df.columns:
            if col == '指导教师姓名':
                teacher_col = col
                break

        # 查找班级/专业列（精确匹配"所属专业"）
        class_col = None
        for col in df.columns:
            if col == '所属专业':
                class_col = col
                break

        # 创建标准化的数据
        result = pd.DataFrame()
        result['课题名称'] = df[title_col].astype(str) if title_col else ''
        result['学生'] = df[student_col].astype(str) if student_col else '未知'
        result['导师'] = df[teacher_col].astype(str) if teacher_col else '未知'

        # 班级信息：优先使用从文件名提取的班级，其次使用Excel中的"所属专业"列
        if 'source_class' in df.columns:
            # 使用从文件名提取的班级
            result['班级'] = df['source_class'].astype(str)
        elif class_col:
            # 使用Excel中的"所属专业"列
            result['班级'] = df[class_col].astype(str)
        else:
            # 都没有则为空
            result['班级'] = ''

        result['来源文件'] = df['source_file'].astype(str) if 'source_file' in df.columns else ''
        result['原始索引'] = df.index

        # 使用新的有效性判断方法过滤无效记录
        valid_mask = result['课题名称'].apply(self.is_valid_title)
        valid_df = result[valid_mask].reset_index(drop=True)
        invalid_df = result[~valid_mask].reset_index(drop=True)

        return valid_df, invalid_df

    def calculate_similarity(self, titles):
        """计算题目之间的相似度"""
        # 使用jieba分词
        print("正在分词和计算相似度...")
        text_split = [' '.join(jieba.lcut(str(t))) for t in titles]

        # TF-IDF向量化
        vectorizer = TfidfVectorizer()
        tfidf_matrix = vectorizer.fit_transform(text_split)

        # 计算余弦相似度
        similarity_matrix = cosine_similarity(tfidf_matrix, tfidf_matrix)
        return similarity_matrix

    def find_similar_pairs(self, df, similarity_matrix):
        """找出相似度超过阈值的题目对"""
        similar_pairs = []
        n = len(df)

        for i in range(n):
            for j in range(i + 1, n):
                similarity = similarity_matrix[i][j]
                if similarity >= self.threshold:
                    similar_pairs.append({
                        '题目A': df.iloc[i]['课题名称'],
                        '题目B': df.iloc[j]['课题名称'],
                        '学生A': df.iloc[i]['学生'],
                        '学生B': df.iloc[j]['学生'],
                        '导师A': df.iloc[i]['导师'],
                        '导师B': df.iloc[j]['导师'],
                        '班级A': df.iloc[i]['班级'],
                        '班级B': df.iloc[j]['班级'],
                        '相似度': f"{similarity:.2%}",
                        '相似度数值': similarity
                    })

        return sorted(similar_pairs, key=lambda x: x['相似度数值'], reverse=True)

    def run(self):
        """执行查重"""
        print("=" * 60)
        print("论文选题查重系统")
        print("=" * 60)

        # 1. 读取数据
        df = self.read_excel_files()
        valid_df, invalid_df = self.normalize_columns(df)

        # 输出无效题目信息
        if len(invalid_df) > 0:
            print(f"发现 {len(invalid_df)} 条无效题目记录（空题、待定等）")

        # 2. 计算相似度（只对有效题目）
        similarity_matrix = self.calculate_similarity(valid_df['课题名称'].tolist())

        # 3. 找出相似题目对
        similar_pairs = self.find_similar_pairs(valid_df, similarity_matrix)

        # 4. 输出结果到多个工作表
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = self.output_dir / f"查重结果_{timestamp}.xlsx"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 工作表1：相似题目对
            if similar_pairs:
                result_df = pd.DataFrame(similar_pairs)
                result_df = result_df.drop(columns=['相似度数值'])
                result_df.to_excel(writer, sheet_name='相似题目', index=False)
            else:
                empty_df = pd.DataFrame(columns=['题目A', '题目B', '学生A', '学生B', '导师A', '导师B', '班级A', '班级B', '相似度'])
                empty_df.to_excel(writer, sheet_name='相似题目', index=False)

            # 工作表2：无效题目记录
            if len(invalid_df) > 0:
                invalid_display = invalid_df[['课题名称', '学生', '导师', '班级', '来源文件']].copy()
                invalid_display.columns = ['课题名称', '学生', '导师', '班级', '来源文件']
                invalid_display.to_excel(writer, sheet_name='无效题目', index=False)
            else:
                empty_invalid_df = pd.DataFrame(columns=['课题名称', '学生', '导师', '班级', '来源文件'])
                empty_invalid_df.to_excel(writer, sheet_name='无效题目', index=False)

        if similar_pairs:
            print(f"\n发现 {len(similar_pairs)} 对相似题目 (阈值: {self.threshold:.0%})")
            print(f"结果已保存到: {output_file}")

            # 显示前10对
            print("\n相似度最高的前10对:")
            print("-" * 60)
            for i, pair in enumerate(similar_pairs[:10], 1):
                print(f"\n{i}. 相似度: {pair['相似度']}")
                print(f"   题目A: {pair['题目A']}")
                print(f"   题目B: {pair['题目B']}")
                print(f"   学生A: {pair['学生A']} | 导师A: {pair['导师A']} | 班级A: {pair['班级A']}")
                print(f"   学生B: {pair['学生B']} | 导师B: {pair['导师B']} | 班级B: {pair['班级B']}")
        else:
            print(f"\n未发现相似题目 (阈值: {self.threshold:.0%})")
            print(f"结果已保存到: {output_file}")

        print("\n查重完成!")
        print(f"有效题目数: {len(valid_df)}")
        print(f"无效题目数: {len(invalid_df)}")
        print(f"相似题目对: {len(similar_pairs)}")


def main():
    """主函数"""
    checker = ThesisChecker(
        data_dir="原始数据",
        output_dir="查重结果",
        threshold=0.75  # 可调整相似度阈值
    )
    checker.run()


if __name__ == "__main__":
    main()
