#!/usr/bin/python3

import os
import sys

import shutil

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


