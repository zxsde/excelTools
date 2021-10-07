#!/usr/bin/python3

import pandas
import openpyxl
import re
import numpy as np
import shutil
import os

""" 简要教程
对 DataFrame 提取行数据
    1. 用标签 loc 定位行(location), 
        row_data = data.loc["a", "b"], 第 a 和 b 行，
        row_data = data.loc["a": "b"], 第 "a"~"b" 行(不包含 "b" 行), 切片是左闭右开, 不包含右边界。

    2. 用索引 iloc 定位行(integer location), 可以指定某几行, 或者某个区间的行 
        row_data = data.iloc[0, 3], 第 0 和 3 行
        row_data = data.iloc[0: 3], 第 0~3 行(不包含 3), 切片是左闭右开, 不包含右边界。


对 DataFrame 提取列数据
        row_data = data[2], 第 2 列, 从 0 开始计数
        row_data = data[[0, 2]], 第 0 和 2 列
        row_data = data[["a", "b"]], 第 "a" 和 "b" 列

        row_data = data.iloc[:, [0, 2]], 第 0 和 2 列, 等同于row_data = data[[0, 2]], 第一个参数:代表所有行
        row_data = data.iloc[[0: 2], [1: 3]], 第 0~2 行, 第 1~3 列, 切片不包含有边界。
        row_data = data.iloc[[0, 2], [1, 2, 3]],第 0 和 2 行, 第 1 2 3 列, 切片代表第 0~2 行(不包含 2)。


对 DataFrame 的指定列排序
    1. 按索引排序
        data.sort_index(), 对行行进行排序, 因为 axis 参数默认为0
        data.sort_index(axis=1), 对列进行排序

    2. 按值排序
        data.sort_values(), 对行进行排序, 因为 axis 参数默认为0
        data.sort_values(by="学科类别"), 对列进行排序
        data=data.sort_values(by="学科类别", axis=1), 对列进行排序。


对 DataFrame 求行列数
        1. data.index, 行数, 可以用data.index.values 转换为 'numpy.ndarray' 类型，类似 list
        2. data.columns, 列数, 可以用 data.columns.values 转换为 'numpy.ndarray' 类型，类似 list
        3. data.keys(), 列数, 和 data.columns 一模一样
        4. list(data), 列数, 和 data.columns 差不多, 但类型是 list
        4. data.shape, 返回一个元组, 格式为 (行数, 列数)


notnull,isnull,dropna

DataFrame 数据转换
    1. DataFrame 转 dict, data.to_dict()
    2. DataFrame 转 list, data.iloc[0].to_list(), 只能对一行/列转换，是一个 Series 类型
    3. DataFrame 转 String, data.astype(str)


Series 支持的方法(https://zhuanlan.zhihu.com/p/100064394)
    1. Series.map(fun), 依次取出 Series 中每个元素，作为参数传递给 fun
        data["gender"] = data["gender"].map({"男":1, "女":0}), 把 gender 列的男替换为1，女替换为0，
    2. Series.applay(fun), 和 map() 差不多，但是可以传入更复杂的参数
        data["age"] = data["age"].apply(apply_age,args=(-3,)), age 列都减 3


DataFrame 支持的方法
    1. DataFrame.applay(fun), 依次取出 DataFrame 中每个元素，作为参数传递给 fun
        data[["height","weight","age"]].apply(np.sum, axis=0), 沿着 0 轴(列)求和
    2. DataFrame.applaymap(fun), 对DataFrame中的每个单元格执行指定函数的操作
        df.applymap(lambda x:"%.2f" % x), 将DataFrame中所有的值保留两位小数显示


dropna 参数介绍，用法 data.iloc[[0, 1]].dropna(axis=1, how='any')
    1. axis，按哪条轴删除，axis=0 表示按行删(默认)，axis=1 表示按列删。
    2. how，删除条件，how='any' 表示只要存在 NaN 就删除(默认)，how='all' 表示全部为 NaN 就删除。
    3. thresh，表示非空元素最低数量，thresh=2 表示小于等于两个空值的会被删除。
    4. subset，子集，对指定的列进行删除，如 subset=["age", "sex"]。
    5. inplace 表示原地替换，inplace=True 表示在元数据上直接更改。
    6. notnull 也可以实现删除，参考【https://www.cnblogs.com/cgmcoding/p/13498229.html】
"""

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

# 汇总表中指定被处理的列，
specific_col = "A:B,F:J"

# 汇总表中科目所在的列
title_col = "B"

# 公司清单
companies_list = "公司清单-210917文旅.xlsx"

# 后缀，只处理该后缀的文件。目前只支持 "xlsx", "xlsm", "xls" 三种格式
excel_suffix = ("xlsx", "xlsm", "xls")

# 前缀，不处理该前缀的文件，该前缀一般是临时文件，防止已打开的 excel 和其临时文件相互合并
excel_prefix = ("~$", "~")

# PBC 简表名字前缀, 用于拼接出被链接的 excel 名称
pbc_prefix = "PBC简表-"

# PBC 简表名字后缀
pbc_suffix = "-202104.xlsx"

# 合并报表所在的 sheet 名
sheet_name = "aaaa"

# 所有的简表名称，从各分公司收回来的表，格式为{excel名字: 绝对路径}
simple_excels = {}

# 从合并报表中拼接中的简表名称，根据这个表来找 simple_excels 中的表
simple_excels_from_summary = []

# 汇总表的 sheet 页
summary_sheet = "1、PBC汇总表"

# 单元格对应的公式，格式为 {C6: SUM(C2:C5)}
cell_formulae = {}

# 链接类公式
link_formulae = {
    ("固定资产", "5"),
    ("投资性资产", "6")
}

# SUM 类公式，格式为 {(科目, 行): 公式}，行是为了保证 key 的唯一性，公式中省去了列数
sum_formulae = {
    ("非流动资产合计", "13"): "5:12",
    ("流动资产合计", "25"): "15:24",
    ("母公司所有者权益", "31"): "28:30",
    ("非流动负债合计", "37"): "35:36",
    ("流动负债合计", "46"): "39:45",
    ("毛利", "56"): "53:55",
    ("经营利润", "64"): "56:63",
    ("税前利润", "67"): "64:66",
    ("净利润", "71"): "67:70",
    ("期末未分配利润", "80"): "73,76:79",
}

# 加减乘除类公式，格式为 {(科目, 行): 公式}，行是为了保证 key 的唯一性，公式中省去了列数
plus_formulae = {
    ("资产合计", "26"): "13+25",
    ("权益合计", "33"): "31+32",
    ("负债合计", "47"): "37+46",
    ("权益及负债合计", "48"): "33+47",
    ("检查", "49"): "26+48",
    ("其中：母公司净利", "73"): "71-74",
    ("检查", "81"): "80-30",
    ("check RE", "84"): "83-76"
}


# 把所有 excel 拷贝到一个文件夹下，并保存所有的 excel 名字和路径，默认不拷贝。
def copy_excel_to_target(is_copy=False):
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


# 从汇总表中获取公司编码和简称，拼接出简表名称，检查是否能找到这些简表
def get_name_from_summary_table(sheet_name="Sheet1", usecols="A:B,F:J"):
    # 拼接出总表的绝对路径
    summary_table_path = os.path.join(root_path, result_path, result_excel)
    summary_table_path = "E:\\pythonProject\\excelTools\\changeCellsLink\\result-202104\\合并报表202104.xlsx"
    if not os.path.exists(summary_table_path):
        print("file not exist: %s" % summary_table_path)
        exit(0)
    # usecols 可选参数 1. 默认 None，全选 2. str类型，"A,B,C" "A:C" "A,B:C" 3. int-list，[0, 1]  4. str-list, ["列名1", "列名2"]
    # 5. 函数，会把列名传入判断函数结果是否为True，可以用 | 做多个判断, usecols=lambda x:x in ["id", "name", "sex"]
    data = pandas.read_excel(summary_table_path, sheet_name="Sheet1", usecols=usecols, header=None)
    print(data)
    # 取第 0 和 1 行，删除空值的列
    filter_nan = data.iloc[[0, 1]].dropna(axis=1, how='any')
    print(filter_nan, end="\n\n")
    # 把这两行数据转为 list
    company_id = filter_nan.iloc[0].to_list()
    company_short = filter_nan.iloc[1].to_list()
    # 过滤掉公司编码为非字母数字的列
    company_short = [company_short[i] for i in range(len(company_id)) if company_id[i].encode('utf-8').isalnum()]
    company_id = [com_id for com_id in company_id if com_id.encode('utf-8').isalnum()]
    print("company_id is: \n %s \n\n company_short_name is:\n %s" % (company_id, company_short), end="\n\n")
    print("%s company_id were recognized, check if it is correct" % len(company_id), end="\n\n")

    # 用公司ID和简称拼接出完整的简表名称
    for com_id, com_name in zip(company_id, company_short):
        simple_excels_from_summary.append(pbc_prefix + com_id + com_name + pbc_suffix)
    print(simple_excels_from_summary, end="\n\n")

    # 检测是否能找到对应的简表
    excel_not_exist = set(simple_excels_from_summary) - set(simple_excels.keys())
    print("%s excels can't found:\n %s" % (len(excel_not_exist), excel_not_exist))

    cal_formulae(data, company_id)


# 计算各科目和公司对应单元格的公式
def cal_formulae(data, company_id) -> dict:
    print("get_formulae:", data.columns.values)
    for i in data.columns.values:
    # for i in range(3):
        com_id = data[i][0]
        # 公司编码不在 company_id 中就跳过，company_id 是经过过滤的公司编码
        if com_id not in company_id:
            continue

        col = convert_to_column(i + 1)
        # 遍历每一行，计算各科目的公式
        for row in data.itertuples():
            # 所有科目都在第二列
            account_title = row[ord(title_col) - ord("A") + 1]
            # 为空时候获取到的是 float 格式的 nan ，直接跳过，我们只解析字符串
            if not isinstance(account_title, str):
                continue

            # 科目和行数，用于匹配是哪种类型的公式，row.Index 从 0 计数，所以比真实的 excel 行数少 1
            title_cell = (account_title.strip(), str(row.Index + 1))
            # 链接类公式处理
            if title_cell in link_formulae:
                continue
            else:
                # 单元格，row.Index 从 0 计数，比真实的 excel 行数少 1，所以需要加 1
                cell = str(col) + str(row.Index + 1)
                # SUM 类公式处理
                if title_cell in sum_formulae:
                    # 拼接出完整的公式，如 SUM(C5:C12)，保存到 cell_formulae
                    formulae = get_formulae(col, sum_formulae[title_cell])
                    cell_formulae[cell] = "SUM({})".format(formulae)
                # PLUS 类公式处理
                elif title_cell in plus_formulae:
                    # 拼接出完整的公式，如 C5+C12，保存到 cell_formulae
                    formulae = get_formulae(col, plus_formulae[title_cell])
                    cell_formulae[cell] = formulae

    print("%s formulae \n %s" % (len(cell_formulae), cell_formulae), end="\n\n")
    return cell_formulae


# 把列转换为字母，如第 27 列转化为 ZA
def convert_to_column(n: int) -> str:
    # ord 返回对应的 ASCII 数值, 'A' = 65
    ascii_letter = (n - 1) % 26 + ord('A')
    if n <= 26:
        return chr(ascii_letter)
    else:
        return convert_to_column((n - 1) // 26) + chr(ascii_letter)


# https://baijiahao.baidu.com/s?id=1626616692056869348&wfr=spider&for=pc
# https://www.cnblogs.com/vhills/p/8327918.html



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


# 拼接公式，把 C 列的 73,76:79 拼接成 C73,C76:C79
def get_formulae(column: str, s: str) -> str:
    formulae = [column, s[0]]
    for i in range(1, len(s)):
        if s[i] == ' ':
            continue
        # 当前字符的前一位不是数字，则插入一个 column
        if not s[i - 1].isdigit():
            formulae.append(column)
        formulae.append(s[i])
    return "".join(formulae)


if __name__ == '__main__':
    print("11")
    # 把所有的 excel 拷贝到指定文件夹
    # copy_excel(is_copy=False)

    get_name_from_summary_table()


    # change_link()