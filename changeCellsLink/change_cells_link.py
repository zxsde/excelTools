#!/usr/bin/python3

import pandas
import openpyxl
import tqdm
import numpy as np
import shutil
import os

# 根路径
root_path = "F:\\Consolidated Statements\PBC\\sunac"

# PBC集合，包含很多文件夹，该路径下所有 excel 都会被拷贝到目标路径
source_path = "【PBC 集合】202106"

# 目标路径
target_path = "source-202106"

# 合并报表的路径
result_path = "result-202106"

# 合并报表
result_excel = "文旅集团合并报表202106.xlsx"

# 公司清单
companies_list = "公司清单-210917文旅.xlsx"

# 后缀，只处理该后缀的文件。目前只支持 "xlsx", "xlsm", "xls" 三种格式
excel_suffix = ("xlsx", "xlsm", "xls")

# 前缀，不处理该前缀的文件，该前缀一般是临时文件，防止已打开的 excel 和其临时文件相互合并
excel_prefix = ("~$", "~")


# 所有的 excel 简表，格式为{excel名字: 绝对路径}
simple_excels = {}

# 汇总表的 sheet 页
summary_sheet = "1、PBC汇总表"


# 把所有 excel 拷贝到一个文件夹下，并保存所有的 excel 名字和路径，默认不拷贝。
def copy_excel(is_copy=False):
    # 拼接出原目文件夹的绝对路径
    source_absolute_path = os.path.join(root_path, source_path)
    target_absolute_path = os.path.join(root_path, target_path)

    # 如果 source_absolute_path 不存在，结束
    if not os.path.exists(source_absolute_path):
        print("source folder not exist,you should create %s first" % source_absolute_path)
        exit(0)

    if is_copy:
        # 如果 target_absolute_path 已存在，先删除，再重新创建，防止文件拷贝越来越多
        if os.path.exists(target_absolute_path):
            print("target folder already exist,delete and re-create %s" % target_absolute_path)
            shutil.rmtree(target_absolute_path)
        os.makedirs(target_absolute_path)

    # 遍历 source_absolute_path，把其下所有 excel 拷贝到 target_absolute_path
    print("copy files from %s to %s ......" % (source_path, target_path))
    for root, dirs, files in os.walk(source_absolute_path):
        for file in files:
            # 临时文件不做处理
            if file.startswith(excel_prefix):
                continue
            # 非 excel 不做处理
            if not file.endswith(excel_suffix):
                continue
            # 构造文件的相对路径
            source_excel = os.path.join(root, file)
            target_excel = os.path.join(target_absolute_path, file)
            if is_copy:
                shutil.copyfile(source_excel, target_excel)
            # 保存所有excel的名字和绝对路径
            simple_excels[file] = target_excel
    # 检查 excel 的个数是否正确
    print("find %s files in %s, check if the number of excel is correct" % (len(simple_excels), target_path))
    print(simple_excels)


# 从汇总表中获取公司编码和简称
def get_name_from_summary_table():
    # 拼接出总表的绝对路径
    summary_table_path = os.path.join(root_path, result_path, result_excel)
    if not os.path.exists(summary_table_path):
        print("file not exist: %s" % summary_table_path)
        exit(0)
    data = pandas.read_excel(summary_table_path, sheet_name=summary_sheet, header=None)
    # 提取公司编码和简称,列的用法是 id_name = data[[0, 2]] 或者 data.iloc[:, [0, 2]], 切片也可以用
    # loc 指用标签定位行(左闭右闭)，iloc 指用索引定位行(左闭右开)，integer location
    # 排序可以用 data=data.sort_values(by='学科类别')
    id_name = data.iloc[[0, 2]]

    print(type(id_name))
    print(pandas.notnull(data.iloc[0:2]))
    id_name.dropna(how="any", axis=1, subset=[0])
    print(data.dropna(how='all'))



def change_link():
    companies_lists = os.path.join(root_path, "source\\公司清单-210917文旅.xlsx")
    # usecols='A,F:H' 取的是第 A 列，第 F ~ H列。
    df = pandas.read_excel(companies_lists, sheet_name="项目&公司名称对应", usecols='G')
    v = df.values
    v = v.flatten()
    print(list(v))
    print(len(v))

    df = pandas.read_excel("F:\\xing\\coding\\modifyCellsLink\\source\\文旅集团合并报表202106.xlsx", sheet_name="1、PBC汇总表", header=None)
    v2 = df.iloc[0].tolist()
    print(v2)
    print(len(v2))

    x = set(v)
    y = set(v2)
    z = x - y
    print(len(z))

    p = sorted([n for n in v if isinstance(n, str)])
    q = sorted([n for n in v2 if (isinstance(n, str) and n.startswith('WL'))])
    print(p)
    print(q)

if __name__ == '__main__':
    print("11")
    # 把所有的 excel 拷贝到指定文件夹
    # copy_excel(is_copy=False)

    get_name_from_summary_table()

    # data = pandas.DataFrame()
    # data['a'] = [1, 2, 3, 4]
    # data['b'] = [1, 2, np.nan, np.nan]
    # print(data)
    # print("------")
    # print(data.iloc[3])
    # print(data[data['b'].notnull()])

    # change_link()