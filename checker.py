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
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 设置输出编码为UTF-8
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class ThesisChecker:
    def __init__(self, data_dir="原始数据", resubmit_dir="二次提交", output_dir="查重结果", threshold=0.75):
        self.data_dir = Path(data_dir)
        self.resubmit_dir = Path(resubmit_dir)
        self.output_dir = Path(output_dir)
        self.threshold = threshold
        self.output_dir.mkdir(exist_ok=True)
        self.resubmit_dir.mkdir(exist_ok=True)  # 确保二次提交文件夹存在

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

    def get_excel_files(self, directory):
        """获取指定目录中的所有Excel文件"""
        excel_files = list(directory.glob("*.xls")) + list(directory.glob("*.xlsx"))
        return excel_files if excel_files else []

    def has_resubmit_data(self):
        """检查是否有二次提交数据"""
        return len(self.get_excel_files(self.resubmit_dir)) > 0

    def read_excel_files(self):
        """读取Excel文件 - 优先使用二次提交文件夹"""
        print("\n正在读取数据文件...")

        # 优先读取二次提交文件夹
        if self.has_resubmit_data():
            print("✓ 检测到二次提交数据，将优先使用二次提交文件夹")
            primary_dir = self.resubmit_dir
            secondary_dir = self.data_dir
            use_resubmit = True
        else:
            print("✓ 二次提交文件夹为空，使用原始数据文件夹")
            primary_dir = self.data_dir
            secondary_dir = None
            use_resubmit = False

        # 读取主目录文件
        excel_files = self.get_excel_files(primary_dir)
        if not excel_files:
            raise FileNotFoundError(f"在 {primary_dir} 文件夹中未找到Excel文件")

        all_data = []
        for file in excel_files:
            try:
                # 尝试不同的header位置
                df = None
                for header_row in [None, 0, 1, 2]:
                    try:
                        df = pd.read_excel(file, header=header_row)
                        # 检查是否包含关键列
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
                    df['is_resubmit'] = use_resubmit  # 标记是否为二次提交数据
                    all_data.append(df)
                    class_info = f" [班级: {df['source_class'].iloc[0]}]" if df['source_class'].iloc[0] else ""
                    source_tag = "[二次提交]" if use_resubmit else "[原始数据]"
                    print(f"✓ 已读取: {file.name}{class_info} {source_tag} ({len(df)} 条记录)")
            except Exception as e:
                print(f"✗ 读取失败: {file.name} - {e}")

        if not all_data:
            raise ValueError("未能成功读取任何Excel文件")

        # 如果有二次提交数据，也读取原始数据作为参考
        if use_resubmit and secondary_dir:
            print("\n正在读取原始数据文件夹作为参考...")
            secondary_files = self.get_excel_files(secondary_dir)
            for file in secondary_files:
                try:
                    # 跳过已经在二次提交中处理过的文件（同名文件）
                    if any(d['source_file'].iloc[0] == file.name for d in all_data if len(d) > 0):
                        continue

                    df = None
                    for header_row in [None, 0, 1, 2]:
                        try:
                            df = pd.read_excel(file, header=header_row)
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
                        df['is_resubmit'] = False  # 标记为原始数据
                        all_data.append(df)
                        class_info = f" [班级: {df['source_class'].iloc[0]}]" if df['source_class'].iloc[0] else ""
                        print(f"✓ 已读取: {file.name}{class_info} [原始数据参考] ({len(df)} 条记录)")
                except Exception as e:
                    print(f"✗ 读取失败: {file.name} - {e}")

        # 合并所有数据
        combined_df = pd.concat(all_data, ignore_index=True)

        # 如果有二次提交数据，进行学生级别的数据合并
        if use_resubmit:
            print("\n正在进行数据合并（按学生去重）...")
            combined_df = self.merge_student_data(combined_df)

        # 显示统计信息
        resubmit_count = sum(1 for d in all_data if d['is_resubmit'].iloc[0] if len(d) > 0)
        original_count = len(all_data) - resubmit_count
        print(f"\n数据读取完成:")
        print(f"  - 总记录数: {len(combined_df)}")
        if use_resubmit:
            print(f"  - 二次提交文件: {resubmit_count} 个")
            print(f"  - 原始数据文件: {original_count} 个（参考）")
            print(f"  ✓ 已按学生合并，以二次提交的课题名称为准")
        else:
            print(f"  - 原始数据文件: {len(all_data)} 个")

        print()
        return combined_df

    def merge_student_data(self, df):
        """合并同一学生的多条记录，以二次提交为准，从原始数据补充缺失信息

        判断同一学生的标准：学生姓名 AND 学号相同
        """
        # 确保有学生姓名和学生学号列
        if '学生姓名' not in df.columns:
            return df

        # 查找学生学号列
        student_id_col = None
        for col in df.columns:
            if '学号' in str(col):
                student_id_col = col
                break

        # 创建唯一标识符（姓名 + 学号）用于判断是否为同一学生
        # 如果没有学号列，则只使用姓名
        if student_id_col is not None and student_id_col in df.columns:
            df['_student_key'] = df['学生姓名'].astype(str) + '|||' + df[student_id_col].astype(str)
        else:
            df['_student_key'] = df['学生姓名'].astype(str)

        # 按唯一标识符分组
        grouped = df.groupby('_student_key')
        merged_indices = []
        duplicate_count = 0

        for student_key, group in grouped:
            if len(group) == 1:
                # 只有一条记录，直接保留
                merged_indices.append(group.index[0])
            else:
                # 多条记录，需要合并（同一学生在二次提交和原始数据中都存在）
                duplicate_count += 1
                # 优先使用 is_resubmit=True 的记录
                resubmit_records = group[group['is_resubmit'] == True]
                original_records = group[group['is_resubmit'] == False]

                if len(resubmit_records) > 0:
                    # 以二次提交记录为基础
                    base_idx = resubmit_records.index[0]
                    # 如果有多条二次提交记录，使用第一条
                    # 从原始数据补充缺失信息（课题名称除外）
                    if len(original_records) > 0:
                        # 就地修改基础记录
                        for col in df.columns:
                            # 跳过课题名称列 - 始终使用二次提交的课题名称
                            if '课题' in str(col) or '题目' in str(col) or '标题' in str(col):
                                continue
                            # 跳过临时列
                            if col == '_student_key':
                                continue

                            base_val = df.at[base_idx, col]
                            # 判断基础值是否为空
                            is_empty = (
                                pd.isna(base_val) or
                                str(base_val).strip() == '' or
                                str(base_val).lower() == 'nan'
                            )
                            if is_empty:
                                # 从原始数据获取值
                                orig_idx = original_records.index[0]
                                orig_val = df.at[orig_idx, col]
                                # 确保补充的值不为空
                                if not pd.isna(orig_val) and str(orig_val).strip() != '' and str(orig_val).lower() != 'nan':
                                    df.at[base_idx, col] = orig_val
                    merged_indices.append(base_idx)
                else:
                    # 没有二次提交记录（理论上不应该发生），使用第一条原始记录
                    merged_indices.append(group.index[0])

        # 删除临时列
        if '_student_key' in df.columns:
            df = df.drop(columns=['_student_key'])

        # 筛选合并后的记录
        merged_df = df.loc[merged_indices].reset_index(drop=True)

        if duplicate_count > 0:
            print(f"  ✓ 合并了 {duplicate_count} 个学生的重复记录")

        return merged_df

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

    def remove_all_spaces(self, value):
        """去除字符串中的所有空格和空白字符，包括不间断空格"""
        if pd.isna(value):
            return value
        result = str(value)
        # 去除各种空白字符
        result = result.replace(' ', '')      # 普通空格
        result = result.replace('\t', '')     # 制表符
        result = result.replace('\n', '')     # 换行符
        result = result.replace('\r', '')     # 回车符
        result = result.replace('\xa0', '')   # 不间断空格 (NBSP)
        result = result.replace('\u3000', '') # 中文全角空格
        result = result.replace('\u200b', '') # 零宽空格
        result = result.replace('\u200c', '') # 零宽非连接符
        result = result.replace('\u200d', '') # 零宽连接符
        result = result.replace('\ufeff', '') # 零宽非断空格 (BOM)
        return result

    def normalize_major(self, major_value, source_class):
        """规范化所属专业字段，只返回两个标准值之一"""
        major_str = str(major_value).strip()

        # 如果已经是标准值，直接返回
        if major_str in ['汽车服务工程技术', '新能源汽车工程技术']:
            return major_str

        # 根据从文件名提取的班级信息来判断
        if source_class:
            if '新能' in source_class or '新能源' in source_class:
                return '新能源汽车工程技术'
            elif '汽服' in source_class or '汽车服务' in source_class:
                return '汽车服务工程技术'

        # 根据原始专业名称中的关键词判断
        if any(keyword in major_str for keyword in ['新能源', '新能', 'NEV']):
            return '新能源汽车工程技术'
        elif any(keyword in major_str for keyword in ['汽车服务', '汽服', '汽车工程']):
            return '汽车服务工程技术'
        elif '测试' in major_str:
            # 测试数据默认归为汽车服务工程技术
            return '汽车服务工程技术'
        else:
            # 默认值
            return '汽车服务工程技术'

    def normalize_columns(self, df):
        """标准化列名，提取所需信息"""
        # 原始文件的完整列名（11列）
        original_columns = [
            '课题名称', '可选范围', '所属专业', '指导教师姓名', '指导教师工号',
            '学生姓名', '学生学号', '学生组织', '来源', '模板类型', '第二指导教师'
        ]

        # 查找各列（使用模糊匹配）
        column_mapping = {}

        # 查找课题名称列
        for col in df.columns:
            if '课题' in str(col) or '题目' in str(col) or '标题' in str(col):
                column_mapping['课题名称'] = col
                break

        # 查找其他列（精确匹配）
        for col in df.columns:
            col_str = str(col)
            if col_str == '可选范围':
                column_mapping['可选范围'] = col
            elif col_str == '所属专业':
                column_mapping['所属专业'] = col
            elif col_str == '指导教师姓名':
                column_mapping['指导教师姓名'] = col
            elif col_str == '指导教师工号':
                column_mapping['指导教师工号'] = col
            elif col_str == '学生姓名':
                column_mapping['学生姓名'] = col
            elif col_str == '学生学号':
                column_mapping['学生学号'] = col
            elif col_str == '学生组织':
                column_mapping['学生组织'] = col
            elif col_str == '来源':
                column_mapping['来源'] = col
            elif col_str == '模板类型':
                column_mapping['模板类型'] = col
            elif col_str == '第二指导教师':
                column_mapping['第二指导教师'] = col

        # 创建标准化的数据，包含所有原始列
        result = pd.DataFrame()

        # 填充所有列（如果找不到则填充空值）
        for orig_col in original_columns:
            if orig_col in column_mapping:
                result[orig_col] = df[column_mapping[orig_col]].astype(str)
            else:
                result[orig_col] = ''

        # 立即去除课题名称中的所有空格（在所有处理之前）
        result['课题名称'] = result['课题名称'].apply(self.remove_all_spaces)

        # 添加辅助列
        source_class = ''
        if 'source_class' in df.columns:
            # 使用从文件名提取的班级
            source_class = df['source_class'].astype(str)
            result['文件提取的班级'] = source_class
        else:
            result['文件提取的班级'] = ''

        result['来源文件'] = df['source_file'].astype(str) if 'source_file' in df.columns else ''
        result['原始索引'] = df.index

        # 规范化所属专业字段为两个标准值之一
        result['所属专业'] = result.apply(
            lambda row: self.normalize_major(row['所属专业'], row['文件提取的班级']),
            axis=1
        )

        # 使用课题名称列进行有效性判断（此时课题名称已无空格）
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
                    # 优先使用文件提取的班级，其次使用所属专业
                    class_a = df.iloc[i]['文件提取的班级'] if df.iloc[i]['文件提取的班级'] else df.iloc[i]['所属专业']
                    class_b = df.iloc[j]['文件提取的班级'] if df.iloc[j]['文件提取的班级'] else df.iloc[j]['所属专业']

                    similar_pairs.append({
                        '题目A': df.iloc[i]['课题名称'],
                        '题目B': df.iloc[j]['课题名称'],
                        '学生A': df.iloc[i]['学生姓名'],
                        '学生B': df.iloc[j]['学生姓名'],
                        '导师A': df.iloc[i]['指导教师姓名'],
                        '导师B': df.iloc[j]['指导教师姓名'],
                        '班级A': class_a,
                        '班级B': class_b,
                        '相似度': f"{similarity:.2%}",
                        '相似度数值': similarity
                    })

        return sorted(similar_pairs, key=lambda x: x['相似度数值'], reverse=True)

    def format_excel_sheet(self, ws, header_row=True):
        """格式化Excel工作表"""
        # 定义边框样式
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        # 定义表头样式
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # 定义数据单元格样式
        cell_alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

        # 应用格式到所有行
        for row_idx, row in enumerate(ws.iter_rows(), 1):
            for cell in row:
                # 应用边框
                cell.border = thin_border

                # 第一行是表头
                if row_idx == 1 and header_row:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
                else:
                    cell.alignment = cell_alignment
                    # 自动调整行高
                    if row_idx > 1:
                        try:
                            value = str(cell.value) if cell.value else ""
                            # 根据内容长度估算行高（每行约15个字符）
                            lines = max(1, (len(value) + 20) // 25)
                            ws.row_dimensions[row_idx].height = max(15, lines * 15)
                        except:
                            ws.row_dimensions[row_idx].height = 15

        # 自动调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    if cell.value:
                        # 计算字符串长度（中文字符按2个长度计算）
                        value = str(cell.value)
                        length = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in value)
                        if length > max_length:
                            max_length = length
                except:
                    pass

            # 设置列宽，限制在合理范围内
            adjusted_width = min(max_length + 2, 50)  # 最大宽度50
            if adjusted_width < 12:
                adjusted_width = 12  # 最小宽度12
            ws.column_dimensions[column_letter].width = adjusted_width

    def highlight_similar_rows(self, ws):
        """为相似题目工作表的高相似度行添加颜色标记"""
        # 定义颜色填充
        high_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')  # 红色
        medium_fill = PatternFill(start_color='FFE6CC', end_color='FFE6CC', fill_type='solid')  # 橙色

        # 从第2行开始（第1行是表头）
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
            # 相似度在最后一列
            similarity_cell = row[-1]
            if similarity_cell.value:
                try:
                    # 提取相似度数值（去掉百分号）
                    sim_str = str(similarity_cell.value).replace('%', '')
                    sim_value = float(sim_str) / 100

                    # 根据相似度设置背景色
                    if sim_value >= 0.85:
                        for cell in row:
                            cell.fill = high_fill
                    elif sim_value >= 0.80:
                        for cell in row:
                            cell.fill = medium_fill
                except:
                    pass

    def highlight_invalid_rows(self, ws):
        """为全部数据工作表的无效题目行添加颜色标记"""
        # 定义无效数据的灰色背景
        invalid_fill = PatternFill(start_color='E0E0E0', end_color='E0E0E0', fill_type='solid')  # 浅灰色

        # 从第2行开始（第1行是表头）
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
            # 检查最后一列的"数据状态"
            status_cell = row[-1]
            if status_cell.value and str(status_cell.value) == '无效':
                for cell in row:
                    cell.fill = invalid_fill

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
                # 选择主要列显示
                invalid_display = invalid_df[['课题名称', '学生姓名', '指导教师姓名', '所属专业', '文件提取的班级', '来源文件']].copy()
                # 合并班级列（优先显示文件提取的班级）
                invalid_display['班级'] = invalid_display['文件提取的班级'].fillna(invalid_display['所属专业'])
                invalid_display = invalid_display[['课题名称', '学生姓名', '指导教师姓名', '班级', '来源文件']]
                invalid_display.columns = ['课题名称', '学生姓名', '指导教师姓名', '班级', '来源文件']
                invalid_display.to_excel(writer, sheet_name='无效题目', index=False)
            else:
                empty_invalid_df = pd.DataFrame(columns=['课题名称', '学生姓名', '指导教师姓名', '班级', '来源文件'])
                empty_invalid_df.to_excel(writer, sheet_name='无效题目', index=False)

            # 工作表3：所有原始数据（合并所有列，包括有效和无效）
            # 合并有效和无效数据
            all_data_with_invalid = pd.concat([valid_df, invalid_df], ignore_index=True)

            # 选择所有原始数据列
            all_columns = [
                '课题名称', '可选范围', '所属专业', '指导教师姓名', '指导教师工号',
                '学生姓名', '学生学号', '学生组织', '来源', '模板类型', '第二指导教师',
                '文件提取的班级', '来源文件'
            ]

            # 确保所有列都存在
            all_data_display = all_data_with_invalid.copy()
            for col in all_columns:
                if col not in all_data_display.columns:
                    all_data_display[col] = ''

            all_data_display = all_data_display[all_columns]

            # 数据规范化
            import random
            random.seed(42)  # 固定随机种子，确保可重复

            # 注意：课题名称中的空格已在 normalize_columns 阶段去除

            # 1. 学生组织统一为"汽车工程学院"
            all_data_display['学生组织'] = '汽车工程学院'

            # 2. 来源为空则随机填写"学生自选"或"教师指定"
            def fill_source(value):
                if pd.isna(value) or str(value).strip() == '' or str(value) == 'nan':
                    return random.choice(['学生自选', '教师指定'])
                return value

            all_data_display['来源'] = all_data_display['来源'].apply(fill_source)

            # 3. 模板类型统一为"理科"
            all_data_display['模板类型'] = '理科'

            # 4. 第二指导教师为空则填写"无"
            def fill_second_teacher(value):
                if pd.isna(value) or str(value).strip() == '' or str(value) == 'nan':
                    return '无'
                return value

            all_data_display['第二指导教师'] = all_data_display['第二指导教师'].apply(fill_second_teacher)

            # 5. 指导教师工号：去除小数点及后续数字
            def clean_teacher_id(value):
                """去除教师工号末尾的小数点及后续数字"""
                if pd.isna(value) or str(value).strip() == '' or str(value) == 'nan':
                    return value
                value_str = str(value)
                # 查找小数点位置，去除小数点及后面的内容
                if '.' in value_str:
                    value_str = value_str.split('.')[0]
                return value_str

            all_data_display['指导教师工号'] = all_data_display['指导教师工号'].apply(clean_teacher_id)

            # 6. 添加数据有效性标注列
            def mark_validity(title):
                """标记题目是否有效"""
                if self.is_valid_title(title):
                    return '有效'
                else:
                    return '无效'

            all_data_display.insert(len(all_data_display.columns), '数据状态', all_data_display['课题名称'].apply(mark_validity))

            # 添加序号列
            all_data_display.insert(0, '序号', range(1, len(all_data_display) + 1))

            # 重命名列以更清晰
            all_data_display.columns = [
                '序号', '课题名称', '可选范围', '所属专业', '指导教师姓名', '指导教师工号',
                '学生姓名', '学生学号', '学生组织', '来源', '模板类型', '第二指导教师',
                '班级(文件名)', '来源文件', '数据状态'
            ]

            # 按班级和学号排序
            # 首先尝试按学生学号（数字）排序，如果学号无法转换为数字则按字符串排序
            def safe_student_id_sort(id_value):
                """安全地提取学号中的数字部分用于排序"""
                try:
                    # 尝试提取学号中的数字部分
                    id_str = str(id_value)
                    # 提取所有数字
                    import re
                    numbers = re.findall(r'\d+', id_str)
                    if numbers:
                        # 返回第一个数字序列作为排序键
                        return int(numbers[0])
                    return 0
                except:
                    return 0

            # 创建排序列
            all_data_display['_sort_class'] = all_data_display['班级(文件名)'].fillna('')
            all_data_display['_sort_id'] = all_data_display['学生学号'].apply(safe_student_id_sort)

            # 先按班级排序，再按学号排序
            all_data_display = all_data_display.sort_values(
                by=['_sort_class', '_sort_id'],
                ascending=[True, True]
            )

            # 删除临时排序列
            all_data_display = all_data_display.drop(columns=['_sort_class', '_sort_id'])

            # 重新生成序号
            all_data_display['序号'] = range(1, len(all_data_display) + 1)

            all_data_display.to_excel(writer, sheet_name='全部数据', index=False)

        # 格式化Excel文件
        print("正在优化Excel格式...")
        wb = load_workbook(output_file)

        # 格式化工作表1：相似题目
        if '相似题目' in wb.sheetnames:
            ws_similar = wb['相似题目']
            self.format_excel_sheet(ws_similar)
            self.highlight_similar_rows(ws_similar)
            # 冻结首行
            ws_similar.freeze_panes = 'A2'

        # 格式化工作表2：无效题目
        if '无效题目' in wb.sheetnames:
            ws_invalid = wb['无效题目']
            self.format_excel_sheet(ws_invalid)
            # 冻结首行
            ws_invalid.freeze_panes = 'A2'

        # 格式化工作表3：全部数据
        if '全部数据' in wb.sheetnames:
            ws_all = wb['全部数据']
            self.format_excel_sheet(ws_all)
            self.highlight_invalid_rows(ws_all)  # 标记无效题目行
            # 冻结首行
            ws_all.freeze_panes = 'A2'

        # 保存格式化后的文件
        wb.save(output_file)

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
        resubmit_dir="二次提交",
        output_dir="查重结果",
        threshold=0.75  # 可调整相似度阈值
    )
    checker.run()


if __name__ == "__main__":
    main()
