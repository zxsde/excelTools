#!/usr/bin/python3

import os
import logging

from tqdm import tqdm
from openpyxl import load_workbook

import conf.common_utils as commons_utils

"""
功能：检查表名
描述：检查指定文件夹下所有 excel 是否包含指定的 Sheet 页
"""


# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# PBC 目录，所有 PBC 表所在的路径
ALL_PBC_PATH = "target\\result-202109\\all_PBC"

# 所有 excel 都应该包含如下 Sheet
STANDARD_SHEET = [
    '科目余额表（请自行贴入）',
    '表0-科目辅助余额表',
    '表1-资产负债表',
    '表2-利润表',
    '表3-现金流量表',
    '表1.1-资产负债分析',
    '表1.2-应收账款账龄分析',
    '表1.3-其他应收款账龄分析',
    '表1.4-预付账款账龄分析',
    '表1.5-应付账款账龄分析',
    '表1.6-其他应付款账龄分析',
    '表1.7-预收账款账龄分析',
    '表1.8-内部往来分析(公式自动）债权方',
    '表1.8-内部往来分析(公式自动）债务方',
    '表1.18-其他流动资产',
    '表1.19-长期借款',
    '表1.9-存货明细 ',
    '表1.10-固定资产明细',
    '表1.11-在建工程明细',
    '表1.12-生物资产明细',
    '表1.13-无形资产',
    '表1.14-长摊',
    '表1.15-应交税金分析',
    '表1.16-货币资金',
    '表1.17-应付职工薪酬',
    '表2.1-收入构成',
    '表2.2-成本构成',
    '表2.3-销售费用',
    '表2.4-管理费用',
    '表2.5-财务费用',
    '表2.6-营业外收入',
    '表2.7-税金及附加',
    '表2.8-营业外支出',
    '表2.9-其他收益'
]

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改

# EXCEl 文件后缀，只处理该后缀的文件。目前不识别 xls 版本
EXCEL_SUFFIX = ("xlsx", "xlsm")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 所有的 excel，格式为 {excel 名: 路径+excel名}
all_pbc = {}

# 日志格式
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(filename='..\\logs\\check_sheet_name.log', level=logging.DEBUG, format=LOG_FORMAT)


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
            file_path = os.path.join(all_pbc_path, root, file)
            all_pbc[file] = file_path
            # print(file)
    print("扫描到 %s 个 excel 文件(不包含xls)，请核实：\n %s" % (len(all_pbc), all_pbc), end="\n\n")
    is_check = input("\033[1;33m 确认文件个数是否正确，是否开始检查？(y/n):")
    if is_check == "y":
        check_sheet_name()


# 检查所有 excel 是否包含指定 Sheet
def check_sheet_name():
    wrong_excel = {}  # {excel: 缺失的 Sheet}
    wrong_order = []  # Sheet 顺序不对的 excel

    # 遍历 excel
    for file, path in tqdm(all_pbc.items()):
        wb = load_workbook(path, read_only=True)
        if wb.sheetnames == STANDARD_SHEET:
            continue
        diff_set = set(STANDARD_SHEET).difference(set(wb.sheetnames))
        if diff_set:
            wrong_excel[file] = list(diff_set)
        else:
            wrong_order.append(file)

    if wrong_order:
        print("Sheet 存在但顺序不对的 excel 有 %s 个，如下，请核实： \n %s" % (len(wrong_order), wrong_order), end="\n\n")
    else:
        print("恭喜！！！所有 excel 的 Sheet 顺序也正确", end="\n\n")
    if wrong_excel:
        print("缺失 Sheet 的 excel 有 %s 个，如下，请核实： \n %s" % (len(wrong_excel), wrong_excel), end="\n\n")
        logging.debug(wrong_excel)
        is_write = input("\033[1;33m 是否对缺失 Sheet 的 excel 创建对应的空 Sheet 页(y/n):")
        if is_write:
            write_sheet(wrong_excel)
    else:
        print("恭喜！！！所有 excel 都包含指定 Sheet", end="\n\n")


# 写入空 Sheet 页
def write_sheet(wrong_excel):
    for excel, sheets in tqdm(wrong_excel.items()):
        keep_vba = False
        # xlsm 是启用了宏的文件，如果不设置keep_vba=True会造成excel无法打开
        if excel.endswith("xlsm"):
            keep_vba = True
        wb = load_workbook(all_pbc[excel], keep_vba=keep_vba)
        for sheet in sheets:
            wb.create_sheet(sheet)
        wb.save(all_pbc[excel])
    logging.debug(wrong_excel)
    print("\033[1;32m" + "Success!!!!!")


if __name__ == '__main__':
    get_all_pbc()
