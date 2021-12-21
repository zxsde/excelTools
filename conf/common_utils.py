#!/usr/bin/python3

import os
import sys

import shutil

from win32com.client import DispatchEx

"""
功能：提供一些公用的方法
"""


# 检查文件/文件夹是否存在
def is_exist(path, is_mkdir=False, is_rm=False):
    """
    :param path: 路径
    :param is_mkdir: 若路径不存在，是否创建
    :param is_rm: 若路径存在，是否删除
    :return:
    """
    # 文件不存在且不需要创建文件，直接退出
    if not os.path.exists(path) and not is_mkdir:
        print("\033[1;31m" + "path not exist: %s" % path)
        sys.exit(0)

    # 文件不存在且需要创建文件夹
    elif not os.path.exists(path) and is_mkdir:
        os.makedirs(path)
        print("\033[1;32m" + "create folder success: %s" % path)

    # 文件存在且需要删除文件，可以删除所有文件/文件夹
    elif os.path.exists(path) and is_rm:
        shutil.rmtree(path)


# 获取指定目录下所有文件，过滤临时文件和非excel文件
def get_all_file(path, temp_prefix, excel_suffix):
    """
    :param path: 将要遍历的目录
    :param temp_prefix: 临时文件,跳过
    :param excel_suffix: excel文件,要处理的文件
    :return:
    """
    all_file = {}  # 保存所有文件,格式为 {文件名: 路径+文件名}
    for root, dirs, files in os.walk(path):
        for file in files:
            # 临时文件不统计
            if file.startswith(temp_prefix):
                continue
            # 非 excel 不统计
            if not file.endswith(excel_suffix):
                continue
            # 构造文件的绝对路径
            file_path = os.path.join(root, file)
            all_file[file] = file_path
    # print(all_file)
    return all_file


# 刷新 Excel
def refresh_file(file):
    # 独立的进程。如果 Excel 已打开，使用 Dispatch 将在打开的 Excel 实例中创建新选项卡，使用 DispatchEx 将打开一个新的 Excel 实例。
    xlapp = DispatchEx("Excel.Application")
    # 设置不可见不警告
    xlapp.Visible = False
    # 打开工作簿
    wb = xlapp.Workbooks.Open(file)
    wb.RefreshAll()
    # 停止宏/脚本,直到刷新完成
    xlapp.CalculateUntilAsyncqueriesDone()
    wb.Save()
    # 关闭，可以清除进程
    xlapp.Quit()
