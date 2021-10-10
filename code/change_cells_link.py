#!/usr/bin/python3

import pandas
import openpyxl
import sys
import os

from win32com.client import Dispatch
from tqdm import tqdm

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "F:\\xing\\excelTools\\"

# PBC集合，包含很多文件夹，该路径下所有 excel 都会被拷贝到目标路径
SOURCE_PATH = "source\\source-202104"

# PBC 目录，所有 PBC 表都将被拷贝到这里
ALL_PBC_PATH = "target\\result-202104\\all_PBC"

# PRC目录，所有 PRC 表都将被拷贝到这里
ALL_PRC_PATH = "target\\result-202104\\all_PRC"

# 合并报表的路径
RESULT_PATH = "target\\result-202104\\summary_table"

# 合并报表的名称
RESULT_TABLE = "合并报表202104.xlsx"

# PBC 简表名字前缀。
PBC_PREFIX = "PBC简表"

# PRC 简表名字前缀。
PRC_PREFIX = "PRC简表"

# PBC 简表名字后缀
TABLE_SUFFIX = "202104.xlsx"

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 汇总表中指定被处理的列，如果所有列都需要处理，则改为 None
specific_col = "A:B,G:J"

# 汇总表中科目所在的列，目前仅支持 'A' - 'Z'
title_col = 'B'

# 公司清单表
COMPANY_LIST = "公司清单-202104苹果.xlsx"

# 简表需要合并的 sheet 名
simple_sheet = "test1"

# 所有的简表名称，从各分公司收回来的表，格式为{excel名字: 绝对路径}
simple_tables = {}

# 从合并报表中拼接中的简表名称，格式为 {公司编码: 简表名称}
simple_tables_from_summary = {}

# 汇总表的 sheet 页
# summary_sheet = "1、PBC汇总表"
summary_sheet = "summary1"

# 单元格对应的公式，格式为 {C6: SUM(C2:C5)}
cell_formulae = {}

# 链接类公式
link_formulae = {
    ("固定资产", "5"): "!$G6",
    ("投资性资产", "6"): "!$G7",
    ("长期金融资产", "7"): "!$G8",
    ("长期股权投资", "8"): "!$G9",
    ("商誉", "9"): "!$G10",
    ("其他无形资产", "10"): "!$G11",
    ("递延所得税资产", "11"): "!$G12",
    ("其他长期资产", "12"): "!$G13",

    ("预付土地款", "15"): "!$G16",
    ("开发中物业", "16"): "!$G17",
    ("持作销售的完工物业", "17"): "!$G18",
    ("应收及其他应收款", "18"): "!$G19",
    ("预付税金", "19"): "!$G20",
    ("应收内部往来", "20"): "!$G21",
    ("应收关联方", "21"): "!$G22",
    ("短期金融资产", "22"): "!$G23",
    ("受限资金", "23"): "!$G724",
    ("现金及现金等价物", "24"): "!$G25",

    ("股本", "28"): "!$G29",
    ("储备", "29"): "!$G30",
    ("留存收益", "30"): "!$G31",

    ("少数股东权益", "32"): "!$G33",

    ("长期借款", "35"): "!$G36",
    ("递延所得税负债", "36"): "!$G37",

    ("其他负债", "39"): "!$G40",
    ("应付及其他应付款", "40"): "!$G41",
    ("预收账款", "41"): "!$G42",
    ("应交税金（核算企业所得税及土地增值税）", "42"): "!$G43",
    ("应付内部往来", "43"): "!$G44",
    ("应付关联方", "44"): "!$G45",
    ("短期借款", "45"): "!$G46",

    ("销售收入", "53"): "!$G54",
    ("减：销售税金及附加", "54"): "!$G55",
    ("销售成本", "55"): "!$G56",

    ("加：投资性物业公允值变化", "58"): "!$G59",
    ("减值准备", "59"): "!$G60",
    ("其他收入", "60"): "!$G61",
    ("减：销售费用", "61"): "!$G62",
    ("管理费用", "62"): "!$G63",
    ("其他费用", "63"): "!$G64",

    ("减：财务费用（仅核算利息支出）", "66"): "!$G67",

    ("减：土地增值税", "69"): "!$G70",
    ("减：企业所得税", "70"): "!$G71",

    ("其中：母公司净利", "73"): "!$G74",
    ("少数股东收益", "74"): "!$G75",

    ("加：期初未分配利润", "76"): "!$G77",
    ("加：新收入准则期初影响", "77"): "!$G78",
    ("减：提取法定盈余公积", "78"): "!$G79",
    ("减：分配利润", "79"): "!$G80",

}

# SUM 类公式，格式为 {(科目, 行): 公式}，行是为了保证 key 的唯一性，公式中省去了列数
sum_formulae = {
    ("非流动资产合计", "13"): "5:12",
    ("流动资产合计", "25"): "15:24",
    ("母公司所有者权益", "31"): "28:30",
    ("非流动负债合计", "37"): "35:36",
    ("流动负债合计", "46"): "39:45",
    ("毛利", "56"): "53:55",
    ("经营利润", "64"): "56:63",
    ("税前利润", "67"): "64:66",
    ("净利润", "71"): "67:70",
    ("期末未分配利润", "80"): "73,76:79",
}

# 加减乘除类公式，格式为 {(科目, 行): 公式}，行是为了保证 key 的唯一性，公式中省去了列数
plus_formulae = {
    ("资产合计", "26"): "13+25",
    ("权益合计", "33"): "31+32",
    ("负债合计", "47"): "37+46",
    ("权益及负债合计", "48"): "33+47",
    ("检查", "49"): "26+48",
    ("其中：母公司净利", "73"): "71-74",
    ("检查", "81"): "80-30",
    ("check RE", "84"): "83-76"
}


# 从汇总表中获取公司编码和简称，拼接出简表名称，检查是否能找到这些简表
def get_name_from_summary_table(sheet_name=summary_sheet, usecols=specific_col):
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
    summary_table_path = os.path.join(ROOT_PATH, RESULT_PATH, RESULT_TABLE)
    print("summary table path: \n", summary_table_path, end="\n\n")
    if not os.path.exists(summary_table_path):
        print("file not exist: %s" % summary_table_path)
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
    company_id = [com_id for com_id in company_id if com_id.encode('utf-8').isalnum()]
    print("company_id is: \n %s \n\n company_short_name is:\n %s" % (company_id, company_short), end="\n\n")
    print("%s company_id were recognized, check if it is correct" % len(company_id), end="\n\n")

    # 用公司ID和简称拼接出完整的简表名称
    for com_id, com_name in zip(company_id, company_short):
        simple_tables_from_summary[com_id] = PBC_PREFIX + "-" + com_id + com_name + "-" +TABLE_SUFFIX
    print(simple_tables_from_summary, end="\n\n")

    # 检测是否能找到对应的简表
    excel_not_exist = set(simple_tables_from_summary.values()) - set(simple_tables.keys())
    print("%s excels can't found:\n %s" % (len(excel_not_exist), excel_not_exist), end="\n\n")

    cal_formulae(data, company_id)


# 计算各科目和公司对应单元格的公式
def cal_formulae(data, company_id) -> dict:
    print("the data include following columns: \n", data.columns.values, end="\n\n")
    target_absolute_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)
    for i in data.columns.values:
        com_id = data[i][0]
        # 公司编码不在 company_id 中就跳过，company_id 是经过过滤的公司编码
        if com_id not in company_id:
            continue

        col = convert_to_column(i + 1)
        # 遍历每一行，计算各科目的公式
        for row in data.itertuples():
            # 所有科目都在第二列
            account_title = row[ord(title_col) - ord('A') + 1]
            # 为空时候获取到的是 float 格式的 nan ，直接跳过，我们只解析字符串
            if not isinstance(account_title, str):
                continue

            # 单元格，row.Index 从 0 计数，比真实的 excel 行数少 1，所以需要加 1
            cell = str(col) + str(row.Index + 1)
            # 科目和行数，用于匹配是哪种类型的公式，row.Index 从 0 计数，所以比真实的 excel 行数少 1
            title_cell = (account_title.strip(), str(row.Index + 1))
            # 链接类公式处理
            if title_cell in link_formulae:
                simple_table = simple_tables_from_summary[com_id]
                cell_formulae[cell] = "={}".format('\'' + target_absolute_path + '[' + simple_table + ']' + simple_sheet + '\'' + link_formulae[title_cell])
                print('\'' + target_absolute_path + '[' + simple_table + ']' + simple_sheet + '\'' + link_formulae[title_cell])
            # SUM 类公式处理
            elif title_cell in sum_formulae:
                # 拼接出完整的公式，如 SUM(C5:C12)，保存到 cell_formulae
                formulae = get_formulae(col, sum_formulae[title_cell])
                cell_formulae[cell] = "=SUM({})".format(formulae)
            # PLUS 类公式处理
            elif title_cell in plus_formulae:
                # 拼接出完整的公式，如 C5+C12，保存到 cell_formulae
                formulae = get_formulae(col, plus_formulae[title_cell])
                cell_formulae[cell] = "={}".format(formulae)

    print("%s formulae \n %s" % (len(cell_formulae), cell_formulae), end="\n\n")
    return cell_formulae


# 写入公式
def write_formulae():
    summary_table_path = os.path.join(ROOT_PATH, RESULT_PATH, RESULT_TABLE)
    print("summary table path: \n", summary_table_path, end="\n\n")
    wb = openpyxl.load_workbook(summary_table_path)
    ws = wb[summary_sheet]
    for k, v in tqdm(cell_formulae.items()):
        ws[k] = v
    wb.save(summary_table_path)
    print("write formulae success, reopen the excel, please wait......")

    # 重新打开一次 excel
    just_open(summary_table_path)
    print("over!!!!!!!!!!!!")


# 重新打开一次 excel，不然无法计算出公式的值，显示 #REF
def just_open(filename):
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(filename)
    xlBook.Save()
    xlBook.Close()


# 把列转换为字母，如第 27 列转化为 ZA
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
    # write_formulae()
