# excelTools
https://www.jianshu.com/p/813b70d5b0de
https://www.cnblogs.com/wj-1314/p/8547763.html

## 国际化
中文 | [English](README_en.md)


## 项目背景

本项目是一个处理 EXCEl 的工具集，可以提高财务人员合并报表的效率，能够实现批量复制、文件名校验、Sheet 合并、公式替换等，具体场景如下：
> 成百上千个分公司的报表发给你，有不同的层级目录，你想把所有 EXCEl 拷贝都某一个文件夹下；
> 成百上千个分公司的报表发给你，你想知道谁还没有提交；
> 成百上千个报表，你想分类放到不同的文件夹下；
> 成败上千个报表，你想检查他们的命名是否规范；
> 成百上千的报表，你想把所有报表的某个 Sheet 页合并(从上到下拼接到一张表内)；
> 汇总表中的公式有 `SUM/PLUS/链接到其他报表的单元格` 等，每个季度需要更新，你想批量修改所有的公式；
> ......
> 等等


### 项目使用

1. 环境安装，安装 python 3 以上的版本，安装 PyCharm 社区版，安装方法网上资料非常多。


2. 安装依赖库
```shell
pip install -r requirements.txt
```


3. 执行脚本
```shell
python copy_excel_to_target.py
```


## 项目文档说明

### 项目文件树

```markdown
├─code                                                    // 核心代码
│      change_cells_link.py
│      copy_excel_to_target.py
│      group_PBC.py
│      merge_excel.py
│      __init__.py
│
├─conf                                                    // 配置文件
│      constant.py
│      __init__.py
│
├─docs                                                    // 说明文档
│      Pandas简易教程.md
│
├─source                                                  // 待处理资料的路径
│  └─source-202104
│      └─全部PBC
│          ├─PBC-xx区-张三-202104
│          │      PBC简表-ID001苹果天津-202104.xlsx
│          │      PBC简表-ID005苹果深圳-202104.xlsx
│          │      PB简表-ID002苹果贵州-202104.xlsx
│          │      PRC简表-ID003苹果福州-202104.xlsx
│          │
│          └─PBC-xx区-李四-202104
│                  PBC简表-ID004苹果北京-202104.xlsx
│                  PBC简表-ID006苹果广州-202104.xlsx
│
└─target                                                  // 处理后资料的路径
    └─result-202104
        ├─all_PBC                                         // 所有 PBC 的路径
        │      PBC简表-ID001苹果天津-202104.xlsx
        │      PBC简表-ID004苹果北京-202104.xlsx
        │      PBC简表-ID005苹果深圳-202104.xlsx
        │      PBC简表-ID006苹果广州-202104.xlsx
        │
        └─summary_table                                   // 汇总表的路径
                合并报表202104-备份.xlsx
                合并报表202104.xlsx
```


### 的


### 是


## 相关项目（可选）


## 主要项目负责人


## 参与贡献方式


## 开源协议


## 更新日志

### V1.0.0 (yyyy-mm-dd)
 - 新功能：aaaaaaaaa
 - 新功能：bbbbbbbbb
 - 修改：ccccccccc
 - 修复：ddddddddd