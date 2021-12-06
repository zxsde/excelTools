# excelTools


## 国际化
中文 | [English](README_en.md)


## 项目背景

本项目是一个处理 EXCEl 的工具集，可以提高财务人员合并报表的效率，能够实现批量复制、文件名校验、Sheet 合并、公式替换等，具体场景如下：

> 数百个分公司的报表发给你，有不同的层级目录，你想把所有 EXCEl 拷贝都某一个文件夹下；
> 
> 数百个分公司的报表发给你，你想知道谁还没有提交；
> 
> 数百个报表，你想分类放到不同的文件夹下；
> 
> 数百个报表，你想检查他们的命名是否规范；
> 
> 数百的报表，你想把所有报表的某个 Sheet 页合并(从上到下拼接到一张表内)；
> 
> 汇总表中的公式有 `SUM/PLUS/链接到其他报表的单元格` 等，每个季度需要更新，你想批量修改所有的公式；
> 
> ......
> 
> 等等。


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
https://www.jianshu.com/p/813b70d5b0de
https://www.cnblogs.com/wj-1314/p/8547763.html

### 是


## 相关项目（可选）


## 主要项目负责人


## 参与贡献方式


## 开源协议


## 更新日志

### V3.0 (2021-12-06)
 - 新功能：新增【刷新excel】的脚本(refresh_excel)，部分excel的公式无法显示为值，可用该脚本打开保存一次刷新公式。
 - 新功能：新增【明细合计】的脚本(detailed_total)，可以把多个excel中各科目下的明细合计起来。
 - 修改：部分变量、常量、注释的优化。
 - 修复：更新公式脚本(change_cells_link)，若超链接的目标文件不存在，则不更新该公式。
 - 修复：拼接Sheet的脚本(check_sheet_name)，修复插入sheet页后无法打开的Bug。

### V2.6 (2021-10-22)
 - 新功能：合并报表(merge_sheet)加强【过滤】功能，可以对多个 Sheet 删除其指定列为某个值的行。
 - 新功能：检查工作表表(check_sheet_name)新增【创建新 Sheet】功能，对缺失 Sheet 的工作簿创建一个空的 Sheet。
 - 修改：暂无。
 - 修复：合并报表(merge_sheet)增加 try except

### V2.5 (2021-10-17)
 - 新功能：新增【检查工作表是否存在】的脚本(check_sheet_name)，检查所有 excel 是否包含指定 Sheet，效率较低，主要是打开工作簿耗时。
 - 新功能：合并报表(merge_sheet)加强【过滤】功能，可以对指定列为指定值的行进行删除，支持多个列多个值。
 - 新功能：合并报表(merge_sheet)新增【指定表头】功能，通过配置 HEADER 参数指定第几行为表头。
 - 新功能：检查工作表(check_sheet_name)新增【展示进度条】功能，写入文件时也会显示进度条，数据量很大时用户体验更好。
 - 修改：SKIP_ROWS 默认置为 0，以前通过 SKIP_ROWS 跳过开头的空行，现在通过 HEADER 直接指定表头所在的行，无需跳过。
 - 修改：往来账款差异(diff_account_current)读取写入数据时保留表头。
 - 修改：公共方法(检查文件/文件夹是否存在)提取到 commons_utils。
 - 修复：暂无

### V2.4 (2021-10-15)
 - 新功能：完善(change_cells_link)，新增用户交互功能，用户可以选择是否进行下一步操作。
 - 新功能：完善(change_cells_link)，新增判断文件/文件夹是否存在的功能
 - 修复：暂无

### V2.3 (2021-10-14)
 - 修改：修改部分注释，常量。
 - 修复：暂无

### V2.2 (2021-10-13)
 - 新功能：新增【计算往来账款差异】的脚本(diff_account_current)，核对甲到乙和乙到甲的往来账款。
 - 新功能：完善(Pandas简易教程)。
 - 修改：脚本(merge_excel)更名为脚本(merge_sheet)。
 - 修改：修改部分注释。
 - 修复：暂无

### V2.1 (2021-10-11)
 - 新功能：项目重构，按照主流项目构建，麻雀虽小五脏俱全，大段的常量拆分到 conf 下，说明文档归档至 docs 下。
 - 新功能：完善(change_cells_link)，完成单元格对应公式的计算，支持链接类的公式。
 - 新功能：完善(change_cells_link)，新增保存后打开 excel 的功能，否则公式不会计算出具体的值。
 - 修改：拆分【批量拷贝 excel 到指定目录】为单独的脚本(copy_excel_to_target)，并新增分类功能，可以将不同类别的 excel 拷贝至不同目录。
 - 修复：暂无

### V1.3 (2021-10-09)
 - 新功能：完善(change_cells_link)，改变一些变量的名字，类型，注释。
 - 新功能：完善(change_cells_link)，完成公式写入的功能。
 - 修复：暂无

### V1.2 (2021-10-08)
 - 新功能：完善(change_cells_link)，完成对公司简称和公司编码的过滤。
 - 新功能：完善(change_cells_link)，完成单元格对应公式的计算，支持 SUM 和 PLUS 的公式。
 - 修复：暂无

### V1.1 (2021-09-30)
 - 新功能：新增【合并 Sheet】的脚本(merge_excel)，单纯的从上到下拼接指定的 Sheet。
 - 新功能：完善(change_cells_link)，完成从汇总表提取公司编码和公司简称的功能。
 - 新功能：总结 DataFrame 语法，方便有兴趣的同学根据自己的需求定制修改代码
 - 修复：暂无

### V1.0 (2021-09-29)
 - 新功能：创建【计算单元格公式】的脚本(change_cells_link)，功能待完善，完成一些前置工作如下。
 - 新功能：完成了批量拷贝 excel 到指定文件夹的功能。
 - 新功能：完成从汇总表读取数据。
 - 修复：暂无