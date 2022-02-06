#!/usr/bin/python3

import os
import sys

import pandas
import openpyxl
from win32com.client import DispatchEx
from tqdm import tqdm

import conf.constant as constant
import conf.common_utils as commons_utils

"""
功能：批量修改PBC链接，单元格和公式的对应关系查看 conf/constant.py
描述：修改指定单元格的链接，功能和change_cells_link.py差不多，为了简单操作，单独写一个脚本
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有 excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# 所有 PBC 保存的路径，可用 copy_excel_to_target.py 把文件保存到该目录
ALL_PBC_PATH = "target\\result-202112\\all_PRC"

# PBC总表的路径，想要更新链接的那个excel所在的路径
SUMMARY_TABLE_PATH = "target\\result-202112\\"

# 总表的名称，想要更新链接的那个excel的名字
SUMMARY_TABLE_NAME = "汇总表202112.xlsx"

# PBC 简表名字前缀。
PBC_PREFIX = "PBC简表"

# PBC 简表名字后缀
PBC_SUFFIX = "202112.xlsx"

# A 是科目所在列，C ~ E 是要更新链接的列，可按需修改。
USE_COLS = "A,K:U"

# 汇总表中科目所在的列，不是 excel 中的列，是在 USE_COLS 的第几列
TITLE_COL = 1

# PBC简表所在行，更新链接就是为了指向这些PBC表
PBC_ROW = 7

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 链接到PBC简表的"资产负债表"
BALANCE_SHEET = "EAS资产负债表（请自行贴入）"

# 174行起链接到PBC简表的"利润表"
INCOME_STATEMENT = (174, "EAS利润表（请自行贴入）")

# 总表的 Sheet 页，要更新链接的那个Sheet
SUMMARY_SHEET = "21.11试算平衡表"

# 链接类公式
LINK_FORMULAE = constant.LINK_FORMULAE_PBC


# 所有的简表名称，从各分公司收回来的表，格式为{excel名字: 绝对路径}
simple_tables = {}

# 从合并报表中获取的简表名称，格式为 [PBC简表名称]，这也是将要更新公式的列
standard_simple_tables = []

# 单元格对应的链接，格式为 {C6: ='E:\xx\[PBC简表-xxx.xlsx]利润表'!$J$16}
cell_formulae = {}


# 从汇总表中获取公司编码和简称，拼接出简表名称，检查是否能找到这些简表
def get_name_from_summary_table(sheet_name=SUMMARY_SHEET, usecols=USE_COLS):
    # 所有 PBC 所在的路径
    pbc_absolute_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)
    # 遍历所有 PBC 简表，保存为 {PBC表名称: 绝对路径} 的格式
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
    print("PBC简表共计 %s 个: \n %s" % (len(simple_tables), simple_tables))

    # 拼接出总表的绝对路径
    summary_table_path = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, SUMMARY_TABLE_NAME)
    print("summary table path: \n", summary_table_path, end="\n\n")
    if not os.path.exists(summary_table_path):
        print("\033[1;31m file not exist: %s" % summary_table_path)
        sys.exit(0)
    # usecols 可选参数 1. 默认 None，全选 2. str类型，"A,B,C" "A:C" "A,B:C" 3. int-list，[0, 1]  4. str-list, ["列名1", "列名2"]
    # 5. 函数，会把列名传入判断函数结果是否为True，可以用 | 做多个判断, usecols=lambda x:x in ["id", "name", "sex"]
    data = pandas.read_excel(summary_table_path, sheet_name=sheet_name, usecols=usecols, header=None)
    # 取第 6 行(从0计数)，即公PBC简表的名称，删除空值的列
    filter_nan = data.iloc[[PBC_ROW - 1]].dropna(axis=1, how='any')

    # 把PBC简表所在行的数据转为 list
    company_name = filter_nan.iloc[0].to_list()
    print("试算平衡底稿中PBC简表: \n %s" % company_name, end="\n\n")
    # 过滤掉非字符串，非字母开头的数据，若判断汉字需用 str.encode('utf-8').isalnum()，不指定编码的话汉字也会返回 True
    company_name = [name for name in company_name if isinstance(name, str) and name[0].encode('utf-8').isalpha()]
    print("试算平衡底稿中需要更新的PBC简表有 %s 个: \n %s \n\n " % (len(company_name), company_name), end="\n\n")

    # 试算平衡表底稿中列出的PBC简表，实际不存在，保存下来
    not_in_pbc = []  # 格式[PBC简表全称]
    # 实际存在的PBC文件，在试算平衡表底稿中没找到，保存下来，
    not_in_sum = []  # 格式[PBC简表全称]
    # 用公司ID和简称拼接出完整的简表名称
    for name in company_name:
        # 拼接出PBC简表的全称
        table_name = PBC_PREFIX + "-" + name + "-" + PBC_SUFFIX
        if table_name not in simple_tables.keys():
            not_in_pbc.append(table_name)
        standard_simple_tables.append(table_name)
    # print(standard_simple_tables, end="\n\n")
    print("\033[1;31m 在试算平衡表底稿中有,但PBC实际不存在的共计 %s 个:\n %s" % (len(not_in_pbc), not_in_pbc), end="\n\n")
    # 查找在PBC合集中有,但汇总表中并没有的表
    for pbc_table in simple_tables.keys():
        if pbc_table not in standard_simple_tables:
            not_in_sum.append(pbc_table)
    print("\033[1;31m 在PBC合集中有,但汇总表中没有的共计 %s 个:\n %s" % (len(not_in_sum), not_in_sum), end="\n\n")

    is_continue = input("\033[1;33m 请检查要更新的 PBC 简表是否都存在，开始计算公式？(y/n):")
    if is_continue == "y":
        cal_formulae(data, company_name)


# 计算各科目和公司对应单元格的公式
def cal_formulae(data, company_name):
    # 将要更新的列
    change_col = [convert_to_letter(col + 1) for col in data.columns.values]
    # 所有 PBC 的绝对路径，第三个参数 '' 是为了在文件夹结尾多一个 \ ，否则拼接文件时会把目录连起来
    target_absolute_path = os.path.join(ROOT_PATH, ALL_PBC_PATH, '')
    for i in data.columns.values:
        # 遍历每一列，取指定行的值，如 data[0][6] 是指excel中的A7，
        pbc_name = data[i][PBC_ROW - 1]
        # company_name 是过滤后pbc简表名字
        if pbc_name not in company_name:
            continue

        # 是字符串的话，就是PBC简表的名字，拼接成全称
        pbc_name = PBC_PREFIX + "-" + pbc_name + "-" + PBC_SUFFIX
        # 试算平衡表底稿中列出的pbc，实际不存在的话，不更新该列
        if pbc_name not in simple_tables:
            continue

        col = convert_to_letter(i + 1)
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
            # 从 174 行起链接到利润表，从 0 计数，所以这里实际是从 173 行起指向利润表
            simple_sheet = BALANCE_SHEET
            if row.Index >= INCOME_STATEMENT[0] - 1:
                simple_sheet = INCOME_STATEMENT[1]
            # 链接类公式处理，链接的格式形如 '路径\[excel名]表名'!$单元格'
            if title_cell in LINK_FORMULAE:
                cell_formulae[cell] = "={}".format('\'' + target_absolute_path + '[' + pbc_name + ']' +
                                                   simple_sheet + '\'' + LINK_FORMULAE[title_cell])

    # print("%s formulae \n %s" % (len(cell_formulae), cell_formulae), end="\n\n")
    is_write = input("\033[1;33m 公式计算完成，是否保存到 %s ？(y/n):" % SUMMARY_TABLE_NAME)
    if is_write == "y":
        summary_table_path = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, SUMMARY_TABLE_NAME)
        commons_utils.is_exist(summary_table_path)
        print("\033[1;33m 正在保存结果，可能需要等待几分钟......")
        write_formulae(summary_table_path)


# 写入公式
def write_formulae(summary_table_path):
    wb = openpyxl.load_workbook(summary_table_path)
    ws = wb[SUMMARY_SHEET]
    for k, v in tqdm(cell_formulae.items()):
        ws[k] = v
    wb.save(summary_table_path)

    # 重新打开一次 excel，否则可能不显示公式计算结果
    is_reopen = input("\033[1;33m 公式写入完成，建议重新打开excel，否则可能无法正常显示公式，是否重新打开excel？(y/n):")
    if is_reopen == "y":
        just_open(summary_table_path)
    print("\033[1;32m" + "reopen Success!!!!!")
    print("\033[1;32m" + "手动打开汇总表时如果提示“此工作簿包含到一个或多个可能不安全的外部源的链接”，选择“更新”，"
                         "如果没有该提示，手动打开excel发现链接显示#REF!，点击“启用内容”")


# 重新打开一次 excel，不然无法计算出公式的值，显示 #REF
def just_open(filename):
    print("\033[1;33m 正在重新打开 excel，可能需要几分钟, 请耐心等待......")
    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(filename)
    xlBook.Save()
    xlBook.Close()


# 把数字列转换为字母，第 1 列对应 "A"，如第 27 列转化为 "AA"
def convert_to_letter(n: int) -> str:
    # ord 返回对应的 ASCII 数值, 'A' = 65
    ascii_letter = (n - 1) % 26 + ord('A')
    if n <= 26:
        return chr(ascii_letter)
    else:
        return convert_to_letter((n - 1) // 26) + chr(ascii_letter)


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
