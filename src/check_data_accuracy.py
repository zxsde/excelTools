#!/usr/bin/python3

import os

import pandas
import numpy
import tqdm

import conf.common_utils as commons_utils

from collections import defaultdict

"""
功能：检查数据准确性
描述：检查指定行列的数据是否正确。
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# PRC目录，所有 PRC 所在路径
ALL_PRC_PATH = "target\\result-202111\\group_PBC"

# 要检查的数据，格式为 {工作表: [(起始单元格, 结束单元格, 应该等于的值)]}
# 误差范围在 0.01内，或者为 "-" 和 "空值" 也算是正确
SPECIFIC_DATA = {
    "目录": [
        ("O10", "P42", ["-", 0, 0.01]),
    ]
}

# PBC 准确性检查表
# SPECIFIC_DATA = {
#     "准确性检查表": [
#         ("G41", "G41", ["-", numpy.NaN, 115.0])
#     ]
# }

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = (".xlsx", ".xlsm")

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改


# 把不正确的数据保存下来
error_res = defaultdict(list)


# 获取文件夹下所有PBC/PRC
def get_all_prc():
    # 拼接出 PRC/PBC 的路径
    prc_path = os.path.join(ROOT_PATH, ALL_PRC_PATH)
    # 判断 PRC/PBC 的路径是否存在
    commons_utils.is_exist(prc_path)

    # 获取所有 PRC/PBC,格式如 {excel全称: excel路径}
    all_prc = commons_utils.get_all_file(prc_path, TEMP_PREFIX, EXCEL_SUFFIX)
    print("all PRC/PBC %s :\n %s" % (len(all_prc), all_prc), end="\n\n")

    is_continue = input("\033[1;33m 获取PRC/PBC成功，请核对个数，是否开始检查数据？(y/n):")
    if is_continue == "y":
        diff_specific_col(all_prc)


def diff_specific_col(all_prc):
    # 遍历每个 PBC/PRC，校验其指定工作表中的数据是否准确
    for prc, prc_path in tqdm.tqdm(all_prc.items()):
        # 遍历每个工作表，校验其指定行列的数据是否准确
        for sheet, data in SPECIFIC_DATA.items():
            data_prc = pandas.read_excel(prc_path, sheet_name=sheet, header=None)
            # 遍历每列，校验其指定单元格的数据是否准确
            for column in data:
                # 拆分单元格为行列，如从 ("O10", "P42", 0.01) 获取O和P列，行数为10~42
                col_start, row_start = split_alpha_num(column[0])
                col_end, row_end = split_alpha_num(column[1])
                # print("start", col_start, row_start, "列转换为数字为", col_start_num)
                # print("end", col_end, row_end, "列转换为数字为", col_end_num)
                target = column[2]
                # 列数转换为数字，如O列转换为第15列
                col_start_num = convert_to_num(col_start)
                col_end_num = convert_to_num(col_end)
                # iloc的第一个参数指定行，第二个参数指定列，左闭右开，从 0 计数，所以起点需要减一
                data_finally = data_prc.iloc[row_start - 1: row_end, col_start_num - 1: col_end_num]
                # print("data:\n", data_finally.values.tolist())
                # 判断值是否正确
                check_data(data_finally.values.tolist(), target, prc, sheet)

    if len(error_res) == 0:
        print("\033[1;32m" + "相同,Success!!!!!")
    else:
        print("\033[1;31m 不相同的文件有 %s 个，如下:\n%s" % (len(error_res), error_res))


# 循环遍历判断每个值，DataFrame转列表后是一个二维数组
def check_data(data, target, file, sheet):
    for row in data:
        for value in row:
            # 浮点型保留两位小数，NaN保留两位小数还是NaN
            if isinstance(value, float):
                value = round(value, 2)
            # 用 value is numpy.NaN 有时不准确，换 numpy.isnan
            if isinstance(value, float) and numpy.isnan(value):
                continue
            elif value not in target:
                error_res[file].append(sheet)
                # 每个工作表只要找到一个不同的，就退出
                return
    return error_res


# 把字母列转换为数字，第 "A" 列对应 1，如第 "ZA" 列转化为 27
def convert_to_num(s: str) -> int:
    res = 0
    for c in s[::]:
        # ord 返回对应的 ASCII 数值, 'A' = 65
        res = res * 26 + ord(c) - ord('A') + 1
    return res


# 把单元格分割为字母和数字,如 "D5" 分割为 "D" 和 5
def split_alpha_num(s: str) -> (str, int):
    alpha_num = ["", 0]
    for i in range(len(s)):
        if s[i].isalpha():
            alpha_num[0] += s[i]
        else:
            alpha_num[1] = int(s[i:])
            break
    return alpha_num


if __name__ == '__main__':
    get_all_prc()
