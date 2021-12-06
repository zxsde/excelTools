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
        print("path not exist: %s" % path)
        sys.exit(0)

    # 文件不存在且需要创建文件夹
    elif not os.path.exists(path) and is_mkdir:
        os.makedirs(path)
        print("create folder success: %s" % path)

    # 文件存在且需要删除文件，可以删除所有文件/文件夹
    elif os.path.exists(path) and is_rm:
        shutil.rmtree(path)


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
