#!/usr/bin/python3

import os
import sys

import pandas
import openpyxl
from win32com.client import Dispatch
from tqdm import tqdm

import conf.constant as constant
import conf.common_utils as commons_utils

"""
功能：批量修改公式，单元格和公式的对应关系查看 conf/constant/constant.py
描述：第一列是科目，第一行是公司编码，第二行是公司简称，对每个公司，计算其科目的公式
     比如连接到其他 excel 单元格，或相加，或求和
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# 所有 PBC 保存的路径，可用 copy_excel_to_target.py 把文件保存到该目录
ALL_PBC_PATH = "target\\result-202111\\all_PBC"

# 汇总表的路径
SUMMARY_TABLE_PATH = "target\\result-202111\\summary_table"

# 汇总表的名称
SUMMARY_TABLE_NAME = "合并报表202111.xlsx"

# PBC 简表名字前缀。
PBC_PREFIX = "PBC简表"

# PBC 简表名字后缀
TABLE_SUFFIX = "202111.xlsx"

# 第一个区间 A ~ B，A是起始列，B 是科目所在列，第一个区间必须包含科目所在列，
# 第二个区间 G ~ J，按需修改，如果所有列都需要处理，则设为 None
USE_COLS = "A:B,DE,GI:GP"

# 汇总表中科目所在的列，不是 excel 中的列，是在 USE_COLS 的第几列
TITLE_COL = 2

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 链接到PBC简表的哪个 Sheet
SIMPLE_SHEET = "表1-B&P"

# 汇总表的 Sheet 页
SUMMARY_SHEET = "1、PBC汇总表"

# 链接类公式
LINK_FORMULAE = constant.LINK_FORMULAE

# SUM 类公式，格式为 {(科目, 行): 公式}，行是为了保证 key 的唯一性，公式中省去了列数
SUM_FORMULAE = constant.SUM_FORMULAE

# 加减乘除类公式，格式为 {(科目, 行): 公式}，行是为了保证 key 的唯一性，公式中省去了列数
PLUS_FORMULAE = constant.PLUS_FORMULAE

# 所有的简表名称，从各分公司收回来的表，格式为{excel名字: 绝对路径}
simple_tables = {}

# 从合并报表中拼接出的简表名称，格式为 {公司编码: 简表名称}，这也是将要更新公式的列
standard_simple_tables = {}

# 单元格对应的公式，格式为 {C6: SUM(C2:C5)}
cell_formulae = {}


# 从汇总表中获取公司编码和简称，拼接出简表名称，检查是否能找到这些简表
def get_name_from_summary_table(sheet_name=SUMMARY_SHEET, usecols=USE_COLS):
    # 所有 PBC 所在的路径
    pbc_absolute_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)
    # 遍历所有 PBC 简表，保存为  格式
    for root, dirs, files in os.walk(pbc_absolute_path):
        for file in files:
            # 临时文件不做处理，即以 "~$", "~" 开头的文件。
            if file.startswith(TEMP_PREFIX):
                continue
            # 非 excel 不做处理，即只处理后缀为 "xlsx", "xlsm", "xls" 的文件
            if not file.endswith(EXCEL_SUFFIX):
                continue
            file_path = os.path.join(pbc_absolute_path, file)
            # 保存所有excel的名字和绝对路径
            simple_tables[file] = file_path
    print(simple_tables)

    # 拼接出总表的绝对路径
    summary_table_path = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, SUMMARY_TABLE_NAME)
    print("summary table path: \n", summary_table_path, end="\n\n")
    if not os.path.exists(summary_table_path):
        print("\033[1;31m file not exist: %s" % summary_table_path)
        sys.exit(0)
    # usecols 可选参数 1. 默认 None，全选 2. str类型，"A,B,C" "A:C" "A,B:C" 3. int-list，[0, 1]  4. str-list, ["列名1", "列名2"]
    # 5. 函数，会把列名传入判断函数结果是否为True，可以用 | 做多个判断, usecols=lambda x:x in ["id", "name", "sex"]
    data = pandas.read_excel(summary_table_path, sheet_name=sheet_name, usecols=usecols, header=None)
    # 取第 0 和 1 行，即公司编码和公司简称，删除空值的列
    filter_nan = data.iloc[[0, 1]].dropna(axis=1, how='any')
    print(filter_nan, end="\n\n")
    # 把公司编码和公司简称转为 list
    company_id = filter_nan.iloc[0].to_list()
    company_short = filter_nan.iloc[1].to_list()
    # 过滤掉公司编码为非字母数字的列
    company_short = [company_short[i] for i in range(len(company_id)) if company_id[i].encode('utf-8').isalnum()]
    # isalnum 不指定编码的话，汉字也会返回 True
    company_id = [com_id for com_id in company_id if com_id.encode('utf-8').isalnum()]
    print("company_id is: \n %s \n\n company_short_name is:\n %s" % (company_id, company_short), end="\n\n")
    print("%s company_id were recognized, check if it is correct" % len(company_id), end="\n\n")

    # 汇总表中拼接出的PBC简表，实际不存在，保存下来，格式{公司编码: PBC简表全称}
    excel_not_exist = {}
    # 用公司ID和简称拼接出完整的简表名称
    for com_id, com_name in zip(company_id, company_short):
        # 拼接出PBC简表的全称
        table_name = PBC_PREFIX + "-" + com_id + com_name + "-" + TABLE_SUFFIX
        if table_name not in simple_tables:
            excel_not_exist[com_id] = table_name
        standard_simple_tables[com_id] = table_name
    print(standard_simple_tables, end="\n\n")
    print("\033[1;31m %s excels can't found:\n %s" % (len(excel_not_exist), excel_not_exist), end="\n\n")

    is_continue = input("\033[1;33m 请检查是否所有 PBC 简表都存在，开始计算公式？(y/n):")
    if is_continue == "y":
        cal_formulae(data, company_id, excel_not_exist)


# 计算各科目和公司对应单元格的公式
def cal_formulae(data, company_id, excel_not_exist):
    # 将要更新的列
    change_col = [convert_to_column(col + 1) for col in data.columns.values]
    print("the following columns will be changed: \n", change_col, end="\n\n")
    # 所有 PBC 的绝对路径，第三个参数 '' 是为了在文件夹结尾多一个 \ ，否则拼接文件时会把目录连起来
    target_absolute_path = os.path.join(ROOT_PATH, ALL_PBC_PATH, '')
    for i in data.columns.values:
        com_id = data[i][0]
        # 公司编码不在 company_id 中就跳过，company_id 是经过过滤的公司编码
        if com_id not in company_id:
            continue
        # 汇总表中拼接出的PBC表，实际找不到，不更新该列
        if com_id in excel_not_exist:
            continue

        col = convert_to_column(i + 1)
        # 遍历每一行，计算各科目的公式
        for row in data.itertuples():
            # 所有科目都在第 TITLE_COL 列
            account_title = row[TITLE_COL]
            # 为空时候获取到的是 float 格式的 nan ，直接跳过，我们只解析字符串
            if not isinstance(account_title, str):
                continue

            # 单元格，row.Index 从 0 计数，比真实的 excel 行数少 1，所以需要加 1
            cell = str(col) + str(row.Index + 1)
            # 科目和行数，用于匹配是哪种类型的公式，row.Index 从 0 计数，所以比真实的 excel 行数少 1
            title_cell = (account_title.strip(), str(row.Index + 1))
            # 链接类公式处理，链接的格式形如 '路径\[excel名]表名'!$单元格'
            if title_cell in LINK_FORMULAE:
                simple_table = standard_simple_tables[com_id]
                cell_formulae[cell] = "={}".format('\'' + target_absolute_path + '[' + simple_table + ']' +
                                                   SIMPLE_SHEET + '\'' + LINK_FORMULAE[title_cell])
                # 216 行的 现金 科目，链接的是附注，仅有一项，所以特殊处理
                if title_cell == ("现金", "216"):
                    cell_formulae[cell] = "={}".format('\'' + target_absolute_path + '[' + simple_table + ']' +
                                                       '附注' + '\'' + LINK_FORMULAE[title_cell])

            # SUM 类公式处理
            elif title_cell in SUM_FORMULAE:
                # 拼接出完整的公式，如 SUM(C5:C12)，保存到 cell_formulae
                formulae = get_formulae(col, SUM_FORMULAE[title_cell])
                cell_formulae[cell] = "=SUM({})".format(formulae)
            # PLUS 类公式处理
            elif title_cell in PLUS_FORMULAE:
                # 拼接出完整的公式，如 C5+C12，保存到 cell_formulae
                formulae = get_formulae(col, PLUS_FORMULAE[title_cell])
                cell_formulae[cell] = "={}".format(formulae)

    print("%s formulae \n %s" % (len(cell_formulae), cell_formulae), end="\n\n")
    is_write = input("\033[1;33m 公式计算完成，是否保存到 %s ？(y/n):" % SUMMARY_TABLE_NAME)
    if is_write == "y":
        summary_table_path = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, SUMMARY_TABLE_NAME)
        commons_utils.is_exist(summary_table_path)
        print("\033[1;33m saving data, the waiting time might be significant, please wait......")
        write_formulae(summary_table_path)


# 写入公式
def write_formulae(summary_table_path):
    wb = openpyxl.load_workbook(summary_table_path)
    ws = wb[SUMMARY_SHEET]
    for k, v in tqdm(cell_formulae.items()):
        ws[k] = v
    wb.save(summary_table_path)

    # 重新打开一次 excel，否则可能不显示公式计算结果
    is_reopen = input("\033[1;33m 公式写入完成，建议重启excel，否则可能无法正常显示公式，是否重启？(y/n):")
    if is_reopen == "y":
        just_open(summary_table_path)
    print("\033[1;32m" + "reopen Success!!!!!")
    print("\033[1;32m" + "手动打开汇总表时如果提示“此工作簿包含到一个或多个可能不安全的外部源的链接”，选择“更新”")


# 重新打开一次 excel，不然无法计算出公式的值，显示 #REF
def just_open(filename):
    print("\033[1;33m 正在重新打开 excel，可能需要几分钟, 请耐心等待......")
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(filename)
    xlBook.Save()
    xlBook.Close()


# 把列转换为字母，第 1 列对应 'A'，如第 27 列转化为 ZA
def convert_to_column(n: int) -> str:
    # ord 返回对应的 ASCII 数值, 'A' = 65
    ascii_letter = (n - 1) % 26 + ord('A')
    if n <= 26:
        return chr(ascii_letter)
    else:
        return convert_to_column((n - 1) // 26) + chr(ascii_letter)


# 拼接公式，把 C 列的 73,76:79 拼接成 C73,C76:C79
def get_formulae(column: str, s: str) -> str:
    formulae = [column, s[0]]
    for i in range(1, len(s)):
        if s[i] == ' ':
            continue
        # 当前字符的前一位不是数字，则插入一个 column
        if not s[i - 1].isdigit():
            formulae.append(column)
        formulae.append(s[i])
    return "".join(formulae)


if __name__ == '__main__':
    get_name_from_summary_table()
