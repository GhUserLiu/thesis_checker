"""
论文选题查重系统 - 工具类模块
提供常用的辅助函数
"""

import pandas as pd
import re
import logging
from typing import Optional

logger = logging.getLogger(__name__)


class DataCleaner:
    """数据清理工具类"""

    @staticmethod
    def remove_all_spaces(value):
        """去除字符串中的所有空格和空白字符，包括不间断空格"""
        if pd.isna(value):
            return value
        result = str(value)
        # 去除各种空白字符
        for space_char in [' ', '\t', '\n', '\r', '\xa0', '\u3000', '\u200b', '\u200c', '\u200d', '\ufeff']:
            result = result.replace(space_char, '')
        return result

    @staticmethod
    def normalize_teacher_id(value):
        """规范化导师工号"""
        if pd.isna(value) or str(value).strip() in ['', 'nan', 'None', 'NA']:
            return '无'
        value_str = str(value).strip()
        # 去除小数点及后续数字
        if '.' in value_str:
            value_str = value_str.split('.')[0]
        return value_str

    @staticmethod
    def normalize_student_id(value):
        """规范化学生学号"""
        if pd.isna(value):
            return '无'
        value_str = str(value).strip()
        # 处理各种空值情况
        if value_str in ['', 'nan', 'None', 'NA']:
            return '无'
        return value_str

    @staticmethod
    def normalize_optional_range(value):
        """规范化可选范围字段"""
        value_str = str(value).strip()
        # 处理各种空值情况
        if value_str in ['', 'nan', 'None', 'NA'] or pd.isna(value):
            return '汽车工程学院'
        # 去除首尾空格和标点符号
        result = value_str.strip().rstrip('。.!！,，')
        # 如果结果为空或太短，返回默认值
        if len(result) < 5:
            return '汽车工程学院'
        return result


class Validator:
    """数据验证工具类"""

    @staticmethod
    def is_valid_title(title):
        """判断题目是否有效（非空、非待定等无意义内容）"""
        if pd.isna(title) or title == '' or str(title).strip() == '':
            return False

        title_str = str(title).strip()

        # 定义无效模式
        invalid_patterns = [
            r'^待定$', r'^暂定$', r'^待填$', r'^\s*无\s*$', r'^空题?$',
            r'^待选$', r'^未定$', r'^nan$', r'^None$', r'^入伍$'
        ]

        for pattern in invalid_patterns:
            if re.search(pattern, title_str, re.IGNORECASE):
                return False

        # 额外检查：去除书名号后，如果内容过短也可能是无效
        content = title_str.replace('《', '').replace('》', '').strip()
        if len(content) <= 2:
            return False

        # 检查是否为纯符号或无意义字符
        if not re.search(r'[\u4e00-\u9fa5a-zA-Z]', content):
            return False

        return True


class FileHelper:
    """文件操作工具类"""

    @staticmethod
    def extract_class_from_filename(filename):
        """从文件名中提取班级信息"""
        # 匹配班号模式：汽服/新能源 + 数字 + 字母
        match = re.search(r'(汽服|汽车服务|新能源汽车|新能|NEV)(\d+[A-Z]+)', filename)
        if match:
            class_name = match.group(1) + match.group(2)
            # 确保包含数字
            if re.search(r'\d', class_name):
                return class_name
        return ''  # 未找到班级名时返回空字符串

    @staticmethod
    def normalize_major(major_value, source_class):
        """规范化所属专业字段为两个标准值之一"""
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

        # 根据专业名称关键词判断
        if any(keyword in major_str for keyword in ['新能源', '新能', 'NEV']):
            return '新能源汽车工程技术'
        elif any(keyword in major_str for keyword in ['汽车服务', '汽服', '汽车工程']):
            return '汽车服务工程技术'

        # 默认值
        return '汽车服务工程技术'


class Logger:
    """统一的日志工具类"""

    @staticmethod
    def info(message):
        """输出信息日志"""
        logging.info(message)

    @staticmethod
    def warning(message):
        """输出警告日志"""
        logging.warning(message)

    @staticmethod
    def error(message):
        """输出错误日志"""
        logging.error(message)

    @staticmethod
    def success(message):
        """输出成功信息"""
        logging.info(f"✓ {message}")

    @staticmethod
    def fail(message):
        """输出失败信息"""
        logging.error(f"✗ {message}")
