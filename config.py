"""
论文选题查重系统配置文件
"""

# 查重阈值（百分比）
SIMILARITY_THRESHOLD = 0.75

# 文件编码
CSV_ENCODING = 'gbk'
CSV_ENCODING_ERRORS = 'ignore'

# Excel输出格式
EXCEL_FONT_SIZE = 11
HEADER_FONT_SIZE = 11

# 颜色配置
COLOR_HEADER_BG = '4472C4'
COLOR_HEADER_FONT = 'FFFFFF'
COLOR_HIGH_SIMILARITY = 'FFC7CE'  # 红色 - 85%以上
COLOR_MEDIUM_SIMILARITY = 'FFEB9C'  # 黄色 - 80-85%
COLOR_INVALID_ROW = 'E0E0E0'  # 灰色 - 无效数据
COLOR_SIMILAR_CELL = 'FFFF00'  # 黄色 - 相似题目

# 数据规范化默认值
DEFAULT_ORGANIZATION = '汽车工程学院'
DEFAULT_TEMPLATE_TYPE = '理科'
DEFAULT_SECOND_TEACHER = '无'

# 随机种子（用于"来源"字段）
RANDOM_SEED = 42

# 导师工号清理规则
def clean_teacher_id(value):
    """清理导师工号的配置"""
    if pd.isna(value) or str(value).strip() in ['', 'nan', 'None', 'NA']:
        return '无'
    value_str = str(value).strip()
    if '.' in value_str:
        value_str = value_str.split('.')[0]
    return value_str

# 忽略的文件列表
IGNORED_FILES = ['.DS_Store', 'Thumbs.db', '~$']
