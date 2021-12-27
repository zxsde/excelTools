#!/usr/bin/python3

import os
import sys

import pandas
import numpy
import logging
import tqdm

import conf.common_utils as commons_utils

from collections import defaultdict

"""
功能：核对资产负债表 balance sheet, income statement
描述：对比两个excel的两个Sheet里指定列的值是否相等,需要预先把excel之间的对应关系处理好,保存在一个excel中.
     实际支持多个excel,多张表,多个列的比较,比较过程中是以PRC表为准的,也就是PRC有的才执行，没的不会执行.
     注意，不支持多对一的比较，如 sheet1 和 sheet2 都与 sheet3 比较的情况，需要分两次比较。
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# PRC目录，所有 PRC 所在路径
ALL_PRC_PATH = "target\\result-202111\\group_PRC"

# PBC 目录，所有 PBC 所在路径
ALL_PBC_PATH = "target\\result-202111\\group_PBC"

# 哪两个 excel 进行对比,记录在这个文件中
COMPANY_LIST = "target\\result-202111\\公司清单-211130.xlsx"

# 公司清单中，记录pbc和prc对比关系的那张表
EXCEL_RELATION = "Sheet1"

# "PBC编码, PBC简称, PRC全称",本场景比较特殊,需要"前缀+编码+简称+后缀"才能拼接出PBC的全称
# 然后 PBC简表-xxx-2021.xlsx 和 PRC-xxx.xlsm 进行对比
USE_COLS = "A,B,C"

# 公司清单表头所在行，从0计数
HEADER = 0

# PBC 简表名字前缀
PBC_PREFIX = "PBC简表"

# PBC 简表名字后缀
PBC_SUFFIX = "202111.xlsx"

# PRC 表名字前缀
PRC_PREFIX = "PRC"

# PRC 中要进行对比的表
PRC_SHEET = [
    # "表1-资产负债表",
    "表2-利润表",
    # "表3-现金流量表"
]

# PBC 中要进行对比的表
PBC_SHEET = [
    # "EAS资产负债表（请自行贴入）",
    "EAS利润表（请自行贴入）"
]

# PRC 和 PBC 对比关系,多对一的情况下,多的表写左边
SHEET_RELATION = {
    # "表1-资产负债表": "EAS资产负债表（请自行贴入）",
    "表2-利润表": "EAS利润表（请自行贴入）",
    # "表3-现金流量表": "EAS利润表（请自行贴入）"
}

# sheet 中要比较的具体列，不可以同时出现相同的 sheet
PRC_COL = {
    "表1-资产负债表": [("D5", "D31"), ("E5", "E31"), ("I5", "I31"), ("J5", "J31")],
    # "表2-利润表": [("C5", "C23"), ("D5", "D23")],
    # "表3-现金流量表": [["C5", "C42"], ["D5", "D42"], ["G5", "G33"], ["H5", "H33"]],
}

# sheet 中要比较的具体列，不可以同时出现相同的 sheet，如第二个和第三个不能同时放开
PBC_COL = {
    "EAS资产负债表（请自行贴入）": [("D5", "D31"), ("E5", "E31"), ("I5", "I31"), ("J5", "J31")],
    # "EAS利润表（请自行贴入）": [("C5", "C23"), ("D5", "D23")],
    # "EAS利润表（请自行贴入）": [["C32", "C69"], ["D32", "D69"], ["G32", "G60"], ["H32", "H60"]],
}

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = (".xlsx", ".xlsm")

# ===================================== 一般情况，仅需修改以上参数，根据实际情况进行修改

# PBC表可能会少一行，PRC中18行-17行的值作为18行的值，且从18行起进行错位比较，
# 用PRC的18行和PBC的17行对比，PRC的19行和PBC的18行对比，以此类推
# missing_line = 0

# 保存所有不相同的数据,格式形如 {excel: [sheet1, sheet2]}
diff_data = defaultdict(list)

# 最终比较结果,记录不一致的表,格式如 {excel: [sheet1, sheet2]}
diff_res = defaultdict(list)

# 获取excel的对比关系,是哪两个excel进行对比,格式形如 {pbc-xxx.xlsx: prc-xxx.xlsm}
excel_relation = {}

# 行列可能错位的文件，形如 {excel: sheet}
wrong_data = defaultdict(list)

# 日志格式
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(filename='..\\logs\\diff_balance_sheet.log', level=logging.INFO, format=LOG_FORMAT)


# 获取文件夹下所有PBC/PRC,以及excel的对比关系
def get_pbc_prc():
    # 拼接出 PRC 和 PBC 的路径
    prc_path = os.path.join(ROOT_PATH, ALL_PRC_PATH)
    pbc_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)
    # 判断 PRC 和 PBC 的路径是否存在
    commons_utils.is_exist(prc_path)
    commons_utils.is_exist(pbc_path)
    # 获取所有 PRC 和 PBC,格式如 {excel全称: excel路径}
    all_prc = commons_utils.get_all_file(prc_path, TEMP_PREFIX, EXCEL_SUFFIX)
    all_pbc = commons_utils.get_all_file(pbc_path, TEMP_PREFIX, EXCEL_SUFFIX)
    print("all PRC %s :\n %s" % (len(all_prc), all_prc), end="\n\n")
    print("all PBC %s :\n %s" % (len(all_pbc), all_pbc), end="\n\n")

    company_list_path = os.path.join(ROOT_PATH, COMPANY_LIST)
    # header 行数从 0 计数,我们平时看excel是从第1行开始,所以这里要减1
    data = pandas.read_excel(company_list_path, sheet_name=EXCEL_RELATION, usecols=USE_COLS, header=HEADER)
    for row in data.itertuples():
        prc_file = PRC_PREFIX + "-" + row[3]
        pbc_file = PBC_PREFIX + "-" + row[1] + row[2] + "-" + PBC_SUFFIX
        # xlsx 后缀
        if prc_file + EXCEL_SUFFIX[0] in all_prc:
            excel_relation[prc_file + EXCEL_SUFFIX[0]] = pbc_file
        # xlsm 后缀
        if prc_file + EXCEL_SUFFIX[1] in all_prc:
            excel_relation[prc_file + EXCEL_SUFFIX[1]] = pbc_file
    # print(excel_relation)
    print("\033[1;33m 进行对比的excel共计 %s 个,请检查是否正确:\n%s" % (len(excel_relation), excel_relation), end="\n\n")
    # 在公司清单中有,实际没找到的PBC简表/PRC表
    in_list_not_in_prc = set(excel_relation.keys()) - set(all_prc.keys())
    in_list_not_in_pbc = set(excel_relation.values()) - set(all_pbc.keys())
    # PRC表对应的PRC不存在
    pbc_not_exist = set(all_prc.keys()) - set(excel_relation.keys())
    # PBC简表对应的PRC表不存在
    prc_not_exist = set(all_pbc.keys()) - set(excel_relation.values())
    print("公司清单中存在,实际没找到的PRC表的有 %s 个:\n %s" % (len(in_list_not_in_prc), in_list_not_in_prc), end="\n\n")
    print("公司清单中存在,实际没找到的PBC表的有 %s 个:\n %s" % (len(in_list_not_in_pbc), in_list_not_in_pbc), end="\n\n")
    print("PRC表存在,但在公司清单中没找到对应的PBC有 %s 个:\n %s" % (len(pbc_not_exist), pbc_not_exist), end="\n\n")
    print("PBC简表存在,但在公司清单中没找到对应的PRC有 %s 个:\n %s" % (len(prc_not_exist), prc_not_exist), end="\n\n")

    # 如果prc对应的pbc找不到,则从all_prc中删掉,接下来以all_prc中的excel为准,找对应的pbc进行对比
    for prc in pbc_not_exist:
        del all_prc[prc]
    is_continue = input("\033[1;33m 是否开始对比？(y/n):")
    if is_continue == "y":
        diff_balance_sheet(all_prc, all_pbc)


def diff_balance_sheet(all_prc, all_pbc):
    # 遍历所有 PRC/PBC excel
    for prc_file, prc_path in tqdm.tqdm(all_prc.items()):
        data_prc = pandas.read_excel(prc_path, sheet_name=PRC_SHEET, header=None)
        pbc_file = excel_relation[prc_file]
        pbc_path = all_pbc[pbc_file]
        data_pbc = pandas.read_excel(pbc_path, sheet_name=PBC_SHEET, header=None)

        # 遍历每张表
        for prc_sheet, pbc_sheet in SHEET_RELATION.items():
            missing_line = 0
            # 保存每个参加比较的数据,格式 {单元格坐标: 数值}
            prc_data = {}
            pbc_data = {}
            # 两张表中单元格对比关系
            cell_relation = {}
            # 两张表的内容转成list,方便遍历
            prc_sheet_list = data_prc[prc_sheet].values.tolist()
            pbc_sheet_list = data_pbc[pbc_sheet].values.tolist()

            # 遍历参与比较的两个sheet的每列
            for prc_cell, pbc_cell in zip(PRC_COL[prc_sheet], PBC_COL[pbc_sheet]):
                try:
                    # 比如 {"表1-资产负债表": [("D5", "D31")]}
                    prc_cell_start = split_alpha_num(prc_cell[0])  # 获取"D5"的行列 ["D", 5]
                    prc_cell_end = split_alpha_num(prc_cell[1])  # 获取"D31"的行列 ["D", 31]
                    prc_col_letter = prc_cell_start[0]  # 获取字母列 "D"
                    prc_col_num = convert_to_num(prc_cell_start[0])  # 获取数字列 4(excel中"D"对应第三列)

                    # 比如 {"EAS利润表": [("C31", "C68")]}
                    pbc_cell_start = split_alpha_num(pbc_cell[0])  # 获取"C31"的行列 ["C", 31]
                    pbc_cell_end = split_alpha_num(pbc_cell[1])  # 获取"C68"的行列 ["C", 68]
                    pbc_col_letter = pbc_cell_start[0]  # 获取字母列 "C"
                    pbc_col_num = convert_to_num(pbc_cell_start[0])  # 获取数字列 3(excel中"c"对应第三列)
                    offset = pbc_cell_start[1] - prc_cell_start[1]  # 参与对比的两列的差值,"D5"和"C31"的差值

                    # pbc的利润表可能会少一行,最后一行,行次不同说明行数不一致
                    prc_last_line = prc_sheet_list[prc_cell_end[1] - 1][0]
                    pbc_last_line = pbc_sheet_list[pbc_cell_end[1] - 1][0]

                    if prc_last_line == "五、净利润（净亏损以“－”号填列）" and pbc_last_line != "五、净利润（净亏损以“－”号填列）":
                        missing_line = 1

                    # "D5"到"D31"之间的数据都保存下来,总共有 31-5+1 个值
                    for i in range(prc_cell_end[1] - prc_cell_start[1] + 1):
                        # "D5"是第5行,从第5行开始遍历到第31行
                        row_prc = prc_cell_start[1] + i
                        row_pbc = pbc_cell_start[1] + i
                        prc_coord = prc_col_letter + str(row_prc)
                        pbc_coord = pbc_col_letter + str(row_pbc)
                        # 保存要对比的两个单元格坐标 {"D5": "D31"}
                        cell_relation[prc_coord] = pbc_coord
                        # "D5" 和 "D31" 对应的值,最终要比较的就是这两个值
                        value_prc_cell = prc_sheet_list[row_prc - 1][prc_col_num - 1]
                        value_pbc_cell = pbc_sheet_list[row_pbc - 1][pbc_col_num - 1]

                        # 只取正负数,判空可以用 value_prc_cell is not numpy.NaN 或者 pandas.isnull(value_prc_cell)
                        if isinstance(value_prc_cell, (int, float)) and (value_prc_cell > 0 or value_prc_cell < 0):
                            prc_data[prc_coord] = round(value_prc_cell, 2)
                        if isinstance(value_pbc_cell, (int, float)) and (value_pbc_cell > 0 or value_pbc_cell < 0):
                            pbc_data[pbc_coord] = round(value_pbc_cell, 2)
                except Exception:
                    print("\033[1;33m excel:" + prc_file + "-----sheet:" + prc_sheet + "-----cell" + str(prc_cell))
                    print("\033[1;33m excel:" + pbc_file + "-----sheet:" + pbc_sheet + "-----cell" + str(pbc_cell))
                    raise
            # print(prc_data)
            # print(pbc_data)
            # pbc表少一行,特殊处理,比如对 "表2-利润表" 中，prc的18行减17行等于pbc的17行,且17后之后错行对比
            if missing_line == 1:
                special_treatments(prc_data, prc_file, prc_sheet)

            logging.info("---------------" + prc_file + " - " + prc_sheet + "非 0 非空的值:")
            logging.info("---------------" + pbc_file + " - " + pbc_sheet + "非 0 非空的值:")
            logging.info("要对比的单元格: %s" % cell_relation)
            logging.info(sorted(prc_data.items(), key=lambda item: item[0]))
            logging.info(sorted(pbc_data.items(), key=lambda item: item[0]))

            check_out(prc_data, pbc_data, prc_file, prc_sheet, cell_relation)
    logging.info("================== 结果: %s" % diff_res)

    if diff_res:
        print("\033[1;31m 对比结束,不相同的文件如下:")
        for file, sheets in diff_res.items():
            tmp_sheet = []
            for sheet in sheets:
                tmp_sheet.append(SHEET_RELATION[sheet])
            print("\033[1;33m %s: %s ------ %s: %s" % (file, sheets, excel_relation[file], tmp_sheet))
            logging.info("%s: %s ------ %s: %s" % (file, sheets, excel_relation[file], tmp_sheet))
        logging.info("over!!!!!!!!!!!!!!!!!! \n\n\n")
    else:
        print("\033[1;32m" + "相同,Success!!!!!")


# pbc的利润表比prc少一行,特殊处理
def special_treatments(prc_data, prc_file, prc_sheet):
    # 如果17行不为空，需要特殊处理18行和19行，先把18，19行加到prc_data
    for coord in list(prc_data.keys()):
        cell = split_alpha_num(coord)
        cord_18 = cell[0] + str(18)
        cord_19 = cell[0] + str(19)
        if cell[1] == 17 and cord_18 not in prc_data.keys():
            prc_data[cord_18] = 0
        if cell[1] == 17 and cord_19 not in prc_data.keys():
            prc_data[cord_19] = 0
    # 17 行不为空，对18行和19行进行计算
    for coord in list(prc_data.keys()):
        cell = split_alpha_num(coord)
        if cell[1] == 17:
            # 17行有数据,18行一定有，否则数据异常
            cord_18 = cell[0] + str(18)
            cord_19 = cell[0] + str(19)

            # 特殊处理，prc中18行的值要减去17行，然后和pbc的17行比较
            prc_data[cord_18] = round(prc_data[cord_18] - prc_data[coord], 2)
            # 特殊处理，prc中19行的值要加上17行，然后和pbc的18行比较
            prc_data[cord_19] = round(prc_data[cord_19] + prc_data[coord], 2)
            # pbc中美prc这个17行，为方便对比，删除
            del prc_data[coord]
        # prc中从18行起，和pbc错行比较，18行和17行比较
        if cell[1] >= 18:
            tmp = prc_data[coord]
            del prc_data[coord]
            prc_data[cell[0] + str(cell[1] - 1)] = tmp


# 比较数据是否相同
def check_out(prc_data, pbc_data, prc_file, prc_sheet, cell_relation):
    if len(prc_data.keys()) != len(pbc_data.keys()):
        diff_res[prc_file].append(prc_sheet)
        return False

    for k1, v1 in prc_data.items():
        k2 = cell_relation[k1]
        if k2 not in pbc_data:
            diff_res[prc_file].append(prc_sheet)
            return False
        if v1 != pbc_data[k2]:
            diff_res[prc_file].append(prc_sheet)
            return False
    return True


# 对输入数据进行校验
def is_valid():
    for prc_sheet, pbc_sheet in SHEET_RELATION.items():
        prc_sheet_col = PRC_COL[prc_sheet]
        pbc_sheet_col = PBC_COL[pbc_sheet]
        if len(prc_sheet_col) != len(pbc_sheet_col):
            print("\033[1;31m" + "两张sheet要对比的列数应相同:\n %s: %s" % (prc_sheet, pbc_sheet))
            sys.exit(0)

        for i in range(len(prc_sheet_col)):
            # prc 表中要对比列的起点/终点
            prc_cell_start = split_alpha_num(prc_sheet_col[i][0])
            prc_cell_end = split_alpha_num(prc_sheet_col[i][1])
            # pbc 表中要对比列的起点/终点
            pbc_cell_start = split_alpha_num(pbc_sheet_col[i][0])
            pbc_cell_end = split_alpha_num(pbc_sheet_col[i][1])
            # 判断起点终点是否在同一列,如 ("D5", "D7") 就不是一个列
            if prc_cell_start[0] != prc_cell_end[0] or pbc_cell_start[0] != pbc_cell_end[0]:
                print("\033[1;31m" + "起始单元格和结束单元格应在同一列:\n %s 的第 %s 组数据" % (prc_sheet, i + 1))
                sys.exit(0)

            # 列的终点应该大于起点
            if prc_cell_start[0] > prc_cell_end[0] or pbc_cell_start[0] > pbc_cell_end[0]:
                print("\033[1;31m" + "列的起点应小于终点:\n %s 的第 %s 组数据" % (prc_sheet, i + 1))
                sys.exit(0)
    print("\033[1;32m 带比较的数据合法,开始执行程序!!!")


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
    is_valid()
    get_pbc_prc()