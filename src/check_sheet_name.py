#!/usr/bin/python3

import os

from tqdm import tqdm
from openpyxl import load_workbook

import conf.common_utils as commons_utils

"""
功能：检查表名
描述：检查指定文件夹下所有 excel 中的 Sheet 名是否正确
"""


# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "f:\\xing\\excelTools\\"

# PBC 目录，所有 PBC 表所在的路径
ALL_PBC_PATH = "target\\result-202109\\all_PBC"

# 所有 excel 都应该包含(不必相同)如下 Sheet
STANDARD_SHEET = [
    '表1-test1',
    '表2-test2',
    '表3-test3',
]

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改

# EXCEl 文件后缀，只处理该后缀的文件。目前不识别 xls 版本
EXCEL_SUFFIX = ("xlsx", "xlsm")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 所有的 excel，包含绝对路径和 excel 名
all_pbc = []

# 缺失 Sheet 的 excel
wrong_excel = {}

# Sheet 顺序不对的 excel
wrong_order = []


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
            if file.endswith("xls"):
                print(file, end="\n\n")
            # 非 excel 不统计
            if not file.endswith(EXCEL_SUFFIX):
                continue
            all_pbc.append(file)
            # print(file)
    print("扫描到 %s 个 excel 文件(不包含xls)，请核实：\n %s" % (len(all_pbc), all_pbc), end="\n\n")
    is_check = input("确认文件个数是否正确，是否开始检查？(y/n):")
    if is_check == "y":
        check_sheet_name()


# 检查所有 excel 是否包含指定 Sheet
def check_sheet_name():
    # 遍历 excel
    for file in tqdm(all_pbc):
        file_path = os.path.join(ROOT_PATH, ALL_PBC_PATH, file)
        wb = load_workbook(file_path, read_only=True)
        if wb.sheetnames == STANDARD_SHEET:
            continue
        diff_set = set(STANDARD_SHEET).difference(set(wb.sheetnames))
        if diff_set:
            wrong_excel[file] = list(diff_set)
        else:
            wrong_order.append(file)

    if wrong_excel:
        print("缺失 Sheet 的 excel 如下，请核实： \n %s" % wrong_excel, end="\n\n")
    else:
        print("恭喜！！！所有 excel 都包含指定 Sheet", end="\n\n")
    if wrong_order:
        print("Sheet 存在但顺序不对的 excel 如下，请核实： \n %s" % wrong_order, end="\n\n")
    else:
        print("恭喜！！！所有 excel 的 Sheet 顺序也正确", end="\n\n")


if __name__ == '__main__':
    get_all_pbc()
