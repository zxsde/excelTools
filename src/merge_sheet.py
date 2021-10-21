#!/usr/bin/python3

import os
import sys

import pandas
import numpy
from tqdm import tqdm

import conf.common_utils as commons_utils

"""
功能：合并 Sheet
描述：把 all_PBC 目录下所有 excel 中指定的 Sheet 合并起来，并删除指定列为指定值的行
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# PBC 目录，所有 PBC 表所在的路径
ALL_PBC_PATH = "target\\result-202109\\all_PBC"

# 合并完后 excel 的保存路径
SUMMARY_TABLE_PATH = "target\\result-202109\\summary_table"

# 合并完后 excel 的名字，重复执行时候建议每次修改名字，因为会覆盖源数据
RESULT_EXCEL = "merge_sheet.xlsx"

# 要合并的 sheet，支持多个 Sheet 各自合并
PENDING_MERGE_SHEETS = [
    "Sheet1",
    "Sheet2",
]

# 黑名单，写在这里的 excel 不会参与合并，为了排除一些个例
BLACK_LIST = {
    "PBC简表-xxx-2021.xlsx",
}

# 指定被处理的列，并非所有列都有用，如果所有列都需要处理，则设为 None
# USE_COLS = "B:C,E:G" 表示只处理 B,C,E,F,G 这几列
USE_COLS = None

# 表头所在的行，从 0 开始计数，无表头改为 None，一般和 SKIP_ROWS 任选一个配置即可
HEADER = 2

# 跳过前几行，有的表会空几行才有数据，跳过之后的第一行为表头
SKIP_ROWS = 0

# 要过滤的数据，如下对于 sheet1 会先过滤掉 "金额" 列中为 0 的行，再过滤掉"公司" 列中为 "小计"和"内部往来" 的行，
# 对于 sheet2 会过滤掉"期末余额"为空的行，不想做任何过滤，把对应的 Sheet 置为空列表。
FILTER_SPECIFIC_VALUE = {
    "Sheet1": {
        "金额": [0],
        "公司": ["小计", "内部往来"]
    },
    "Sheet2": {
        "期末余额": [numpy.NaN]
    },
}

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

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
    print("\033[1;33m 扫描到 %s 个 excel 文件:\n %s" % (len(all_pbc), all_pbc), end="\n\n")
    print("\033[1;33m 将要合并 %s 个 Sheet:\n %s" % (len(PENDING_MERGE_SHEETS), PENDING_MERGE_SHEETS), end="\n\n")
    is_merge = input("\033[1;33m 是否开始合并如下 Sheet ？(y/n):")
    if is_merge == "y":
        merge_sheet()


# 合并所有 excel 的指定 Sheet
def merge_sheet():
    if not all_pbc or not PENDING_MERGE_SHEETS:
        print("\033[1;31m PENDING_MERGE_SHEETS or all_pbc is null!!")
        sys.exit(0)
    for sheet_name in PENDING_MERGE_SHEETS:
        dfs = pandas.DataFrame()
        for file in tqdm(all_pbc):
            # file 包含路径，截取 excel 名判断是否在黑名单中，在黑名单不处理
            if file.split("\\")[-1] in BLACK_LIST:
                continue
            # 跳过 SKIP_ROWS 行后，data 的第 0 行实际上是 excel 中的第 SKIP_ROWS 行
            try:
                data = pandas.read_excel(file,
                                         sheet_name=sheet_name,
                                         header=HEADER,
                                         usecols=USE_COLS,
                                         skiprows=SKIP_ROWS)
            except ValueError as Argument:
                print("\033[1;31m ValueError:%s:\n excel: %s \n" % (Argument, file))
                sys.exit(0)
            # 对指定 Sheet，过滤掉指定列为某个值的行
            data = filter_value(data, file, sheet_name)
            short_name = get_short_name(file)
            # data["shortName"] = short_name
            # data["source"] = file
            data.insert(0, "shortName", short_name)
            data.insert(0, "source", file.split("\\")[-1])
            # concat默认纵向连接 DataFrame 对象，并且合并之后不改变每个 DataFrame 子对象的 index 值
            dfs = pandas.concat([dfs, data])
        merged_sheets.append(dfs)

    is_write = input("\033[1;33m 数据处理已完成，是否保存到 %s ？注意：源数据会被覆盖，请做好备份(y/n):" % RESULT_EXCEL)
    if is_write == "y":
        result_excel = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, RESULT_EXCEL)
        print("saving data, please wait......")
        sava_file(result_excel, merged_sheets)


# 保存到本地
def sava_file(result, data):
    with pandas.ExcelWriter(result) as writer:
        for i in tqdm(range(len(data))):
            data[i].to_excel(writer, sheet_name=PENDING_MERGE_SHEETS[i])
    print("\033[1;32m" + "Success!!!!!")


# 过滤掉符合条件的行，对指定 Sheet 的指定列为指定值的行进行过滤
def filter_value(data, file, sheet_name):
    if sheet_name not in FILTER_SPECIFIC_VALUE:
        return data
    for col, val in FILTER_SPECIFIC_VALUE[sheet_name].items():
        try:
            data = data[~data[col].isin(val)]
        except KeyError as Argument:
            print("\033[1;31m KeyError:%s:\n excel: %s \n Sheet: %s" % (Argument, file, sheet_name))
            sys.exit(0)
    return data


# 从 excel 全称中截取公司简称
def get_short_name(com_name):
    short_name = []
    com_name = com_name.split("\\")[-1]  # 切割后公司简称所在的下标是 1
    # com_name = unicode(com_name, 'utf-8')
    for c in com_name:
        if "\u4e00" <= c <= "\u9fa5":
            short_name.append(c)
    return "".join(short_name)


if __name__ == '__main__':
    get_all_pbc()
