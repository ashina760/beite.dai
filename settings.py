# settings.py
from datetime import datetime

#监视文件夹路径
WARCH_DIR = r"C:\11111"

#配送可能日期excel路径
REFERENCE_PATH = r"C:\Users\beite.dai.bp\OneDrive - Coca-Cola Bottlers Japan\デスクトップ\py\aeon_rpa_project0424\NPFKB.xlsx"

#本机下载文件路径
DOWNLOADS_PATH = r"D:\DATA\Downloads"

#间隔多少秒检查一次
INTERVAL = 5  # 秒

#担当者名称所在excel单元格位置
CELLS_LIST = ["K6"]

# 标题校验配置
TITLE_COLUMNS = ["C", "D", "E", "F", "G", "H", "I", "J", "K"]
EXPECTED_TITLES = [
    "仕入先コード", "入荷倉庫コード", "商品コード", "商品名（伝票用）",
    "発注数量", "納期", "発注単価", "発注金額", "伝票摘要"
]

# 空值检查配置
MANDATORY_CELLS = ["K6"]

# 用于删除空行检查的列
MANDATORY_COLUMN = "G"  

# 日期校验列配置
DATE_COLUMN = "H"
ID_COLUMN = "D"

# 上传excel配置nagashikomi
FILL_VALUES = {
    1: "D",
    5: "9002624",
    6: "2",
    7: "9",
    8: "99"
}

#找出数据区域中 min_col~ max_col 列所有空单元格
MIN_COL = 3
MAX_COL = 11

# 创建带时间戳
TIMESTAMP = datetime.now().strftime("%Y%m%d-%H%M")

#将两个表格对比，类似VLOOKUP的功能
# key_columns_in_a: List[str]，A 表中参与 key 的列名，如 ['C', 'D', 'E']
# key_columns_in_b: List[str]，CSV 中对应的列名，如 ['col1', 'col2', 'col3']
# value_column_in_b: str，CSV 中要写入 A 表的值所在的列名，如 'F'
# target_column_in_a: str，写入 A 表的目标列名，如 'M'
KEY_COLUMNS_IN_A = ["C", "D", "E", "G", "H", "K"]
KEY_COLUMNS_IN_B = ["仕入先コード", "センターコード", "商品コード", "発注数量", "指定納期", "伝票備考"]
VALUE_COLUMN_IN_B = "発注番号"
TARGET_COLUMN_IN_A = "M"


# Chrome 相关配置
CHROME_PATH = r"C:\chrome-win64\chrome.exe"
CHROMEDRIVER_PATH = r"C:\chromedriver-win64\chromedriver.exe"

# 登录信息
AEON_OPCD = ""
AEON_PASSWORD = ""
