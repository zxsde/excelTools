#!/usr/bin/python3

import os

import pandas
import numpy


"""
功能：计算"合计"，待完善
描述：比如每个公司都有库存商品，库存商品下还有本期增加，本期摊销等，
     现在要把所有公司的，库存商品下的明细合计起来
"""

# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改

# 根路径，所有代码，excel 的所在路径
ROOT_PATH = "E:\\excelTools\\"

# 处理完后 excel 的保存路径
SUMMARY_TABLE_PATH = "target\\result-202109\\summary_table"

# 源数据
SOURCE_DATA = "merge_sheet.xlsx"

# 计算完合计后保存在这里，重复执行时候建议每次修改名字，因为会覆盖源数据
RESULT_EXCEL = "total_result.xlsx"

# 每个 Sheet 的行列关系，格式为 {"Sheet": ["明细"所在行, "科目"所在列]}，行从 0 计数，列从 1 计数，
# 如下代表 "表1.9-存货明细 " 的第 3 行存放着明细，第 6 列存放着科目。
DATA_RELATION = {
    "表1.9-存货明细 ": [3, 6],
    # "表1.13-无形资产": [4, 5],
}

PENDING_MERGE_SHEETS = [
    "表1.9-存货明细 ",
    # "表1.13-无形资产",
]
# ===================================== 一般情况，仅需修改如下参数，根据实际情况进行修改


#
def get_data():
    # 拼接出源数据的绝对路径
    summary_table_path = os.path.join(ROOT_PATH, SUMMARY_TABLE_PATH, SOURCE_DATA)
    data = pandas.read_excel(summary_table_path, sheet_name=PENDING_MERGE_SHEETS, header=None)
    # print(data)
    # 格式 {科目: [{明细1: 金额}, {明细2: 金额}]}
    detail_total = {}
    for sheet, sheet_data in data.items():
        detail_row = DATA_RELATION[sheet][0]            # 明细所在行
        subject_col = DATA_RELATION[sheet][1]           # 科目所在列
        detail_data = sheet_data.iloc[detail_row].to_list()   # 明细所在行，保存一整行数据
        # subject_col = convert_to_column(DATA_RELATION[sheet][1])
        print("sheet:", sheet)
        print(detail_data)
        for row in sheet_data.itertuples():
            # if detail_row-1 < row.Index < detail_row + 5:
            if row[subject_col] is numpy.NaN:
                continue
            # print(row)
            subject = row[subject_col]
            detail_money = {}    # 格式 {明细: 金额}
            # 科目还未遍历过，其明细都设置为 0
            if subject not in detail_total:
                # 科目不在字典中，value 设为一个空列表，用循环进行填充
                detail_total[subject] = []
                for i in range(subject_col, len(detail_data)):
                    detail_money[detail_data[i]] = 0
                detail_total[subject] = detail_money
                # print("hhhhhhhhhhhh", detail_total)
            # 科目已遍历过，明细相加
            else:
                # print("+++++++++++++++++++")
                # print(row)
                for i in range(subject_col, len(detail_data)):
                    if row[i + 1] is numpy.NaN:
                        continue
                    if isinstance(row[i + 1], str):
                        continue
                    detail_total[subject][detail_data[i]] += row[i + 1]
                # print("nnnnnnnnnnnnnnnnnn", detail_total)

        print("====")
        for k, v in detail_total.items():
            print(k, v)


# 把列转换为字母，第 1 列对应 'A'，如第 27 列转化为 ZA
def convert_to_column(n: int) -> str:
    # ord 返回对应的 ASCII 数值, 'A' = 65
    ascii_letter = (n - 1) % 26 + ord('A')
    if n <= 26:
        return chr(ascii_letter)
    else:
        return convert_to_column((n - 1) // 26) + chr(ascii_letter)


if __name__ == '__main__':
    get_data()
