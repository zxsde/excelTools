#!/usr/bin/python3

import shutil
import os
import sys

"""##############      一般情况，仅需修改如下参数，每个季度的用不同的文件夹      ##############"""

# PBC集合，包含很多文件夹，该路径下所有 excel 都会被拷贝到目标路径
SOURCE_PATH = "source\\source-202104"

# PBC 目录，所有 PBC 表都将被拷贝到这里
ALL_PBC_PATH = "target\\result-202104\\all_PBC"

# PRC目录，所有 PRC 表都将被拷贝到这里
ALL_PRC_PATH = "target\\result-202104\\all_PRC"

# 非 PBC 和 PRC 开头的文件被拷贝到这里
OTHER_FILE_PATH = "target\\result-202104\\other"

"""##############      一般情况，仅需修改如上参数，每个季度的用不同的文件夹      ##############"""

# 根路径，项目所在的目录
ROOT_PATH = "D:\\excelTools\\"

# PBC 简表名字前缀。
PBC_PREFIX = "PBC"

# PRC 简表名字前缀。
PRC_PREFIX = "PRC"

# EXCEl 文件后缀，只处理该后缀的文件。
EXCEL_SUFFIX = ("xlsx", "xlsm", "xls")

# 临时文件前缀，不处理该前缀的文件，该前缀一般是临时文件，如已打开的 excel 会额外生成一个额 ~ 开头的文件
TEMP_PREFIX = ("~$", "~")


# 检验文件夹是否存在，校验文件类型，统计待拷贝的文件。
def calculate_pending_file():
    # 源文件夹路径，待处理文件所在的路径
    source_absolute_path = os.path.join(ROOT_PATH, SOURCE_PATH)
    # 目标路径，PBC 表将要拷贝到的路径
    target_pbc_path = os.path.join(ROOT_PATH, ALL_PBC_PATH)
    # 目标路径，PRC 表将要拷贝到的路径
    target_prc_path = os.path.join(ROOT_PATH, ALL_PRC_PATH)
    # 目标路径，非 PBC/PRC 表将要拷贝到的路径
    target_other_path = os.path.join(ROOT_PATH, OTHER_FILE_PATH)
    print(" source: %s \n target_pbc_path: %s \n target_prc_path: %s \n target_other_path: %s" % (
            source_absolute_path, target_pbc_path, target_prc_path, target_other_path), end="\n\n")

    # 如果根目录不存在，结束
    if not os.path.exists(ROOT_PATH):
        print("root path not exist,you should create %s first" % ROOT_PATH)
        sys.exit(0)

    # 如果源目录不存在，结束
    if not os.path.exists(source_absolute_path):
        print("source folder not exist,you should create %s first" % source_absolute_path)
        sys.exit(0)

    # 如果 all_PBC 目录不存在，创建文件夹，删除用 shutil.rmtree(path)
    if not os.path.exists(target_pbc_path):
        os.makedirs(target_pbc_path)

    # 如果 all_PRC 目录不存在，创建文件夹。
    if not os.path.exists(target_prc_path):
        os.makedirs(target_prc_path)

    # 如果 other 目录不存在，创建文件夹。
    if not os.path.exists(target_other_path):
        os.makedirs(target_other_path)

    existing_pbc = []
    existing_prc = []
    existing_other = []
    # 找出 all_PBC/all_PRC/other 下面已经存在的 excel
    for root, dirs, files in os.walk(target_pbc_path):
        existing_pbc.extend(files)

    for root, dirs, files in os.walk(target_prc_path):
        existing_prc.extend(files)

    for root, dirs, files in os.walk(target_other_path):
        existing_other.extend(files)

    pbc_tables = {}
    prc_tables = {}
    other_tables = {}
    # 遍历源目录，识别出 PBC/PRC/other 表，保存成 {表名: 源路径} 格式
    for root, dirs, files in os.walk(source_absolute_path):
        for file in files:
            # 临时文件不做处理，即以 "~$", "~" 开头的文件。
            if file.startswith(TEMP_PREFIX):
                continue
            # 非 excel 不做处理，即只处理后缀为 "xlsx", "xlsm", "xls" 的文件
            if not file.endswith(EXCEL_SUFFIX):
                continue

            # 构造文件的绝对路径
            source_path = os.path.join(root, file)
            pbc_path = os.path.join(target_pbc_path, file)
            prc_path = os.path.join(target_prc_path, file)
            other_path = os.path.join(target_other_path, file)
            if file.startswith(PBC_PREFIX):
                pbc_tables[file] = (source_path, pbc_path)
            elif file.startswith(PRC_PREFIX):
                prc_tables[file] = (source_path, prc_path)
            else:
                other_tables[file] = (source_path, other_path)
    # 把所有获取到的表转换为集合，计算新增了哪些表
    new_pbc = set(pbc_tables.keys()) - set(existing_pbc)
    new_prc = set(prc_tables.keys()) - set(existing_prc)
    new_other = set(other_tables.keys()) - set(existing_other)
    print("%s new PBC_table were scanned: \n %s \n" % (len(new_pbc), new_pbc))
    print("%s new PRC_table were scanned: \n %s \n" % (len(new_prc), new_prc))
    print("%s new other_table were scanned: \n %s \n" % (len(new_other), new_other))

    # 是否执行拷贝
    is_copy = input("是否把以上新增的表分别拷贝到对应的目标路径下？(y/n):")
    if is_copy == "y":
        all_tables = dict(pbc_tables, **prc_tables, **other_tables)
        copy_file_to_target(all_tables, new_pbc, new_prc, new_other)
    else:
        sys.exit(0)


# 把所有 PBC 拷贝到 all_PBC 目录下，所有 PRC 拷贝到 all_PRC 目录下。分别拷贝到 all_PBC/all_PRC 目录下
def copy_file_to_target(all_tables, new_pbc, new_prc, new_other):
    count_pbc = 0
    count_prc = 0
    count_other = 0

    # 拷贝 PBC 表
    for pbc in new_pbc:
        shutil.copyfile(all_tables[pbc][0], all_tables[pbc][1])
        count_pbc += 1

    # 拷贝 PRC 表
    for prc in new_prc:
        shutil.copyfile(all_tables[prc][0], all_tables[prc][1])
        count_prc += 1

    # 拷贝其他表
    for other in new_other:
        shutil.copyfile(all_tables[other][0], all_tables[other][1])
        count_other += 1

    print(" %s pbc copied \n %s prc copied \n %s other copied \n" % (count_pbc, count_prc, count_other))


if __name__ == '__main__':
    calculate_pending_file()
