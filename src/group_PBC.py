#!/usr/bin/python3


import os
import sys

import pandas
import numpy
import shutil
import tqdm
import conf.common_utils as commons_utils


"""
功能：对 PBC 和 PRC 表进行分组
描述：把 all_PBC/all_PRC 目录下所有文件，保存在各自相应的目录下
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径

ROOT_PATH = "E:\\excelTools\\"

# 所有 PBC 所在的目录，设为空字符串则不对PBC进行分组，可用作分组开关
ALL_PBC_PATH = "target\\result-202201\\group\\all_PBC"

# 所有 PRC 所在的目录，设为空字符串则不对PRC进行分组，可用作分组开关
ALL_PRC_PATH = "target\\result-202201\\group\\all_PRC"

# 分组后保存的目录
GROUP_PRC = "target\\result-202201\\group\\group_PRC"
GROUP_PBC = "target\\result-202201\\group\\group_PBC"

# PBC 简表名字前缀
PBC_PREFIX = "PBC简表"

# PBC 简表名字后缀
PBC_SUFFIX = "202111.xlsx"

# PRC 表名字前缀
PRC_PREFIX = "PRC"

# 公司清单所在目录，该文件记录了各excel应该分组到哪个目录
COMPANY_LIST = "target\\result-202201\\公司清单-211130.xlsx"

# 公司清单中，记录pbc和prc分组目录的那张表
EXCEL_RELATION = "Sheet1"

# 表头所在行
HEADER = 1

# 要用到的列，分别是（pbc编码，pbc简称，prc名称，pbc分组目录，prc分组目录）
USE_COLS = "A,B,C,D,E"

# PBC目录所在列数
COL_PRC = 5

# PRC目录所在列数
COL_PBC = 4

# 临时目录，分组时会把所有excel拷贝到该目录，再剪切到最终目录，
# 程序执行完后该目录下剩余的文件需要手动处理，重复执行程序时该目录会被删除
TEMP_FILE_PRC = "target\\result-202201\\group\\temp_prc"
TEMP_FILE_PBC = "target\\result-202201\\group\\temp_pbc"

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = (".xlsx", ".xlsm")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")

# 获取excel的对比关系,是哪两个excel进行对比,格式形如 {pbc-xxx.xlsx: prc-xxx.xlsm}
excel_relation = {}


# 预处理，删除旧得临时文件夹，创建新的文件夹，拷贝文件到临时文件夹。代码冗余，没时间改了
def pre_processing():
    temp_file_prc = os.path.join(ROOT_PATH, TEMP_FILE_PRC)
    temp_file_pbc = os.path.join(ROOT_PATH, TEMP_FILE_PBC)

    # 临时文件存在的话，先删除，ALL_PRC_PATH设为空则不对PRC分组
    if ALL_PRC_PATH and os.path.exists(temp_file_prc):
        shutil.rmtree(temp_file_prc)

    if ALL_PBC_PATH and os.path.exists(temp_file_pbc):
        shutil.rmtree(temp_file_pbc)

    # 拼接出 PRC 和 PBC 的路径
    prc_path = os.path.join(ROOT_PATH, ALL_PRC_PATH)
    pbc_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)

    if not os.path.exists(prc_path):
        print("\033[1;31m" + "文件不存在: %s" % prc_path)
        sys.exit(0)
    # ALL_PRC_PATH 不为空则拷贝 PRC 到临时文件夹
    elif ALL_PRC_PATH:
        print("prc文件拷贝中，请等待......")
        shutil.copytree(prc_path, temp_file_prc)
        print("prc文件已拷贝至: %s" % temp_file_prc)

    if not os.path.exists(pbc_path):
        print("\033[1;31m" + "文件不存在: %s" % pbc_path)
        sys.exit(0)
    # ALL_PBC_PATH 不为空则拷贝 PBC 到临时文件夹
    elif ALL_PBC_PATH:
        print("pbc文件拷贝中，请等待......")
        shutil.copytree(pbc_path, temp_file_pbc)
        print("pbc文件已拷贝至: %s" % temp_file_pbc)

    is_continue = input("\033[1;33m prc/pbc已拷贝至临时文件夹，是否继续？(y/n):")
    if is_continue == "y":
        get_file_group()


def get_file_group():
    group_prc = {}  # 保存分组后的PRC和路径，格式为 {PRC 全称: 路径}
    group_pbc = {}  # 保存分组后的PBC和路径，格式为 {PBC 全称: 路径}
    all_prc = {}  # 临时目录下的PRC及路径，格式为 {PBC 全称: 路径}
    all_pbc = {}  # 临时目录下的PBC及路径，格式为 {PBC 全称: 路径}
    # 拼接出 PRC 和 PBC 的路径
    temp_file_prc = os.path.join(ROOT_PATH, TEMP_FILE_PRC)
    temp_file_pbc = os.path.join(ROOT_PATH, TEMP_FILE_PBC)

    # 获取所有 PRC 和 PBC,格式如 {excel全称: excel路径}
    if ALL_PRC_PATH:
        all_prc = commons_utils.get_all_file(temp_file_prc, TEMP_PREFIX, EXCEL_SUFFIX)
        print("PRC 文件个数 %s :\n %s" % (len(all_prc), all_prc), end="\n\n")
    if ALL_PBC_PATH:
        all_pbc = commons_utils.get_all_file(temp_file_pbc, TEMP_PREFIX, EXCEL_SUFFIX)
        print("PBC 文件个数 %s :\n %s" % (len(all_pbc), all_pbc), end="\n\n")

    company_list_path = os.path.join(ROOT_PATH, COMPANY_LIST)
    # header 行数从 0 计数,我们平时看excel是从第1行开始,所以这里要减1
    data = pandas.read_excel(company_list_path, sheet_name=EXCEL_RELATION, usecols=USE_COLS, header=HEADER-1)
    for row in data.itertuples():
        if row[COL_PRC] is not numpy.NaN:
            # 拼接出PRC文件的全称，PRC名称在第三列
            prc_file = PRC_PREFIX + "-" + row[3]
            prc_path = os.path.join(ROOT_PATH, GROUP_PRC, row[COL_PRC])
            # xlsx 后缀
            if ALL_PRC_PATH and prc_file + EXCEL_SUFFIX[0] in all_prc:
                group_prc[prc_file + EXCEL_SUFFIX[0]] = prc_path
            # xlsm 后缀
            elif ALL_PRC_PATH and prc_file + EXCEL_SUFFIX[1] in all_prc:
                group_prc[prc_file + EXCEL_SUFFIX[1]] = prc_path

        if row[COL_PBC] is not numpy.NaN:
            # 拼接出PBC文件的全称，PBC编码在第一列，PBC简称在第二列，PBC后缀全是 xlsx
            pbc_file = PBC_PREFIX + "-" + row[1] + row[2] + "-" + PBC_SUFFIX
            pbc_path = os.path.join(ROOT_PATH, GROUP_PBC, row[COL_PBC])
            if ALL_PBC_PATH and pbc_file in all_pbc:
                group_pbc[pbc_file] = pbc_path

    src_to_tar_prc = {all_prc[k]: v for k, v in group_prc.items()}
    src_to_tar_pbc = {all_pbc[k]: v for k, v in group_pbc.items()}
    print("PRC 分组信息 %s 个:\n %s" % (len(group_prc), group_prc), end="\n\n")
    print("PBC 分组信息 %s 个:\n %s" % (len(group_pbc), group_pbc), end="\n\n")

    is_continue = input("\033[1;33m 分组目录计算完成，是否开始分组？(y/n):")
    if is_continue == "y":
        if ALL_PRC_PATH:
            file_grouping(src_to_tar_prc)
        if ALL_PBC_PATH:
            file_grouping(src_to_tar_pbc)


def file_grouping(path_to_path):
    print("\033[1;32m 开始移动文件....")
    for source_path, target_path in path_to_path.items():
        if not os.path.exists(target_path):
            os.makedirs(target_path)
        shutil.move(source_path, target_path)


if __name__ == '__main__':
    is_begin = input("\033[1;33m 程序会先拷贝一份PRC/PBC至temp_prc/temp_pbc目录，"
                     "如果还未拷贝至临时目录，请输入y，否则请输入n。(y/n):")
    if is_begin == "y":
        pre_processing()
    else:
        get_file_group()
