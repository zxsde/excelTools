#!/usr/bin/python3

import pandas
import openpyxl
import tqdm
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
        data.shape 返回一个元组，格式为 (行数, 列数)


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
simple_excels_target = []


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
    summary_table_path = "D:\\testexcel\\合并报表202104.xlsx"
    if not os.path.exists(summary_table_path):
        print("file not exist: %s" % summary_table_path)
        exit(0)
    data = pandas.read_excel(summary_table_path, sheet_name="hebing1", header=None)
    # 提取 DataFrame 的行数据，用索引定位行的方法： data[[0, 2]] 或者 data.iloc[:, [0, 2]], 切片也可以。
    # iloc 指用索引定位行(左闭右开)，integer location，loc 指用标签定位行(左闭右闭)。
    # 排序可以用 data=data.sort_values(by='学科类别')。
    print(data.iloc[[0, 1]])

    # axis=1 表示按列删, how='any' 表示只要存在 NaN 就删除，thresh 表示非空元素最低数量，subset 表示子集，inplace 表示原地替换
    # 这里 dropna 会把包含 NaN 的所有列删除
    # notnull 也可以实现，参考【pandas 的 notnull() 的返回非空值函数的用法】
    filter_nan = data.iloc[[0, 1]].dropna(axis=1, how='any')
    print("filter null:" + filter_nan)
    company_id = filter_nan.iloc[0].to_list()
    company_short = filter_nan.iloc[1].to_list()
    print(company_id)
    print(company_short)
    for id, name in zip(company_id, company_short):
        simple_excels_target.append(pbc_prefix+id+name+pbc_suffix)
    print(simple_excels_target)



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