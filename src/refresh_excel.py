#!/usr/bin/python3

import os

from tqdm import tqdm

import conf.common_utils as commons_utils

"""
功能：刷新 excel 中的公式
描述：有的 excel 中公式未刷新，直接用脚本读取到的结果是 NaN，刷新后就能用脚本取到公式的值。
     手动打开再保存也可以，excel 较多时候可以用此程序批量打开
"""


# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# PBC 目录，此文件夹下所有 excel 都会被刷新
ALL_EXCEL_PATH = "target\\result-202111\\test"

# ===================================== 一般情况，仅需修改如上参数，根据实际情况进行修改

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 所有 excel 的路径
all_excel = []


# 获取文件夹下所有的工作簿
def get_all_excel():
    all_excel_path = os.path.join(ROOT_PATH, ALL_EXCEL_PATH)
    commons_utils.is_exist(all_excel_path)
    for root, dirs, files in os.walk(all_excel_path):
        for file in files:
            # 临时文件不统计
            if file.startswith(TEMP_PREFIX):
                continue
            # 非 excel 不统计
            if not file.endswith(EXCEL_SUFFIX):
                continue
            # 构造文件的绝对路径
            file_path = os.path.join(root, file)
            all_excel.append(file_path)
            # print(file_name)
    print("\033[1;33m 扫描到 %s 个 excel 文件:\n %s" % (len(all_excel), all_excel), end="\n\n")
    is_refresh = input("\033[1;33m 是否刷新所有 Excel ？(y/n):")
    if is_refresh == "y":
        for file in tqdm(all_excel):
            commons_utils.refresh_file(file)
        print("\033[1;32m" + "Success!!!!!")


if __name__ == '__main__':
    get_all_excel()

