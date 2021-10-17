#!/usr/bin/python3

import os

import pandas
from tqdm import tqdm
from numpy import NaN

import conf.common_utils as commons_utils

"""
功能：合并 Sheet
描述：把 all_PBC 目录下所有 excel 中 Sheet "表7-内部关联往来" 合并起来，并删除 "调整后" 列为 0/空 的行
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# PBC 目录，所有 PBC 表所在的路径
ALL_PBC_PATH = "target\\result-202104\\all_PBC"

# 合并完后 excel 的保存路径
SUMMARY_TABLE_PATH = "target\\result-202104\\summary_table"

# 合并完后 excel 的名字
RESULT_EXCEL = "merge_sheet.xlsx"

# 表头所在的行，从 0 开始计数，无表头改为 None，一般和 SKIP_ROWS 任选一个配置即可
HEADER = 2

# 跳过前几行，有的表会空几行才有数据，跳过之后的第一行为表头
SKIP_ROWS = 0

# 要过滤的值，列表结构，如果需要过滤空值，可以用 numpy.NaN，如果不过滤，删掉列表中所有值
FILTER = []

# 指定不为空的列，该列为空的行会被删除
SPECIFIC_COL = "调整后"

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 要合并的 sheet，从 0 开始计数，如 [7, 12] 意思是合并第 7 和 12 个 sheet
PENDING_MERGE_SHEETS = [
    # "表1.1-资产负债分析",
    # "表1.3-其他应收款账龄分析",
    # "表1.6-其他应付款账龄分析",
    # "表1.7-预收账款账龄分析"
    "表7-内部关联往来",
]

# 所有的 excel，如 "E:\\excel\\xxx.xlsx"
all_pbc = []

# 合并后的数据
merged_sheets = []


# 从 all_PBC 下获取所有的表
def get_all_pbc():
    all_pbc_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)
    print("all_pbc: \n", all_pbc_path, end="\n\n")
    commons_utils.is_exist(all_pbc_path)
    for root, dirs, files in os.walk(all_pbc_path):
        for file in files:
            # 临时文件不统计
            if file.startswith(TEMP_PREFIX):
                continue
            # 非 excel 不统计
            if not file.endswith(EXCEL_SUFFIX):
                continue
            # 构造文件的绝对路径
            file_path = os.path.join(root, file)
            all_pbc.append(file_path)
            # print(file_name)
    print("扫描到 %s 个 excel 文件:\n %s" % (len(all_pbc), all_pbc), end="\n\n")
    is_merge = input("确认文件个数是否正确，是否开始合并 Sheet %s ？(y/n):" % PENDING_MERGE_SHEETS)
    if is_merge == "y":
        merge_sheet()


# 合并所有 excel 的指定 Sheet
def merge_sheet():
    dfs = []
    for sheet_name in PENDING_MERGE_SHEETS:
        dfs = pandas.DataFrame()
        for file in tqdm(all_pbc):
            # 跳过 SKIP_ROWS 行后，data 的第 0 行实际上是 excel 中的第 SKIP_ROWS 行
            data = pandas.read_excel(file, sheet_name=sheet_name, skiprows=SKIP_ROWS, header=HEADER)
            # FILTER 不为空，就过滤掉 data 中 SPECIFIC_COL 列为 FILTER 的行
            if FILTER:
                data = data[~data[SPECIFIC_COL].isin(FILTER)]
            data["source"] = file
            # concat默认纵向连接 DataFrame 对象，并且合并之后不改变每个 DataFrame 子对象的 index 值
            dfs = pandas.concat([dfs, data])
        merged_sheets.append(dfs)

    is_write = input("数据处理已完成，是否保存到 %s ？注意：源数据会被覆盖，请做好备份(y/n):" % RESULT_EXCEL)
    if is_write == "y":
        result_excel = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, RESULT_EXCEL)
        print("saving data, please wait......")
        sava_file(result_excel, merged_sheets)


# 保存到本地
def sava_file(result, data):
    with pandas.ExcelWriter(result) as writer:
        for i in tqdm(range(len(data))):
            data[i].to_excel(writer, sheet_name=PENDING_MERGE_SHEETS[i])
    print("over!!!!!")


if __name__ == '__main__':
    get_all_pbc()
