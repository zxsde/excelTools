#!/usr/bin/python3

import os
import sys

import shutil
import pandas

from tqdm import tqdm

"""
功能：合并 Sheet
描述：把 all_PBC 目录下所有 excel 中 Sheet "表7-内部关联往来" 合并起来，并删除 "调整后" 列为 0/空 的行
"""

# ===================================== 一般情况，仅需修改如下参数，因为每个月的文件目录/文件名都会变化

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "D:\\excelTools\\"

# PBC 目录，所有 PBC 表所在的路径
ALL_PBC_PATH = "target\\result-202104\\all_PBC"

# 合并完后 excel 的保存路径
SUMMARY_TABLE_PATH = "target\\result-202104\\summary_table"

# 合并完后 excel 的名字
RESULT_EXCEL = "merge_sheet.xlsx"

# 跨过每张 sheet 的前几行，有的表会空两行，第三行才是表头
OFFSET = 2

# 要过滤的值
FILTER = ["", 0]

# 指定不为空的列，该列为空的行会被删除
SPECIFIC_COL = "调整后"

# ===================================== 一般情况，仅需修改以上参数，因为每个月的文件目录/文件名都会变化

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 要合并的 sheet，从 0 开始计数，如 [7, 12] 意思是合并第 7 和 12 个 sheet
pending_merge_sheets = [
    # "表1.1-资产负债分析",
    # "表1.3-其他应收款账龄分析",
    # "表1.6-其他应付款账龄分析",
    # "表1.7-预收账款账龄分析"
    "表7-内部关联往来",
]

# 所有的 excel，包含绝对路径和 excel 名
all_pbc = []

# 合并后的数据
merged_sheets = []


def get_all_pbc():
    all_pbc_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)
    print("all_pbc: \n", all_pbc_path, end="\n\n")
    is_exist(all_pbc_path)
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
    print("扫描到 %s 个 excel 文件，请确认是否准确" % (len(all_pbc)))
    print(all_pbc)


# 合并所有 excel 的指定 Sheet
def merge_sheet():
    dfs = []
    for sheet_name in pending_merge_sheets:
        dfs = pandas.DataFrame()
        for file in tqdm(all_pbc):
            # 跳过 OFFSET 行后，data 的第 0 行实际上是 excel 中的第 OFFSET 行
            data = pandas.read_excel(file, sheet_name=sheet_name, skiprows=OFFSET)
            # 选取 data 中 "调整后" 列为空和 0 的行，然后取反
            # 删除指定列为空的行可以用 data = df.dropna(subset=["调整后"], axis=0, how='any')
            data = data[~data[SPECIFIC_COL].isin(FILTER)]
            data["source"] = file
            # concat默认纵向连接DataFrame对象，并且合并之后不改变每个DataFrame子对象的index值
            dfs = pandas.concat([dfs, data])
        merged_sheets.append(dfs)

    is_write = input("数据处理已完成，是否保存到 %s ？(y/n):" % RESULT_EXCEL)
    if is_write == "y":
        result_excel = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, RESULT_EXCEL)
        is_exist(result_excel)
        print("saving data, please wait......")
        sava_file(result_excel, merged_sheets)


# 保存到本地
def sava_file(result, data):
    with pandas.ExcelWriter(result) as writer:
        for i in range(len(data)):
            data[i].to_excel(writer, sheet_name=pending_merge_sheets[i])
    print("over!!!!!")


# 检查文件/文件夹是否存在
def is_exist(file, is_mkdir=False, is_rm=False):
    # 文件不存在且不需要创建文件，直接退出
    if not os.path.exists(file) and not is_mkdir:
        print("file not exist: %s" % file)
        sys.exit(0)

    # 文件不存在且需要创建文件，可以递归创建目录
    elif not os.path.exists(file) and is_mkdir:
        os.makedirs(file)

    # 文件存在且需要删除文件，可以删除所有文件/文件夹
    elif os.path.exists(file) and is_rm:
        shutil.rmtree(file)


if __name__ == '__main__':
    get_all_pbc()
    merge_sheet()


