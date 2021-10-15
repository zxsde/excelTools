
## pandas 简易教程
pandas 是一个常用于数据分析的 python 库，我们用 pandas 处理 excel 数据，用 openpyxl 写入 excel，读取数据的方式：
```python
data = pandas.read_excel(path, sheet_name, usecols, header)
```

pandas.read_excel 参数有很多，后续再补充。pandas.read_excel 读取出来的是 DataFrame 格式的数据，下面看一些常用的方法。

### 对 DataFrame 提取行数据
1. 用标签 loc 定位行(location), 
```python
row_data = data.loc["a", "b"], # 第 a 和 b 行，
row_data = data.loc["a": "b"], # 第 "a"~"b" 行(不包含 "b" 行), 切片是左闭右开, 不包含右边界。
```

2. 用索引 iloc 定位行(integer location), 可以指定某几行, 或者某个区间的行
```python
row_data = data.iloc[0, 3], # 第 0 和 3 行
row_data = data.iloc[0: 3], # 第 0~3 行(不包含 3), 切片是左闭右开, 不包含右边界。
```


### 对 DataFrame 提取列数据
```python
row_data = data[2], # 第 2 列, 从 0 开始计数
row_data = data[[0, 2]], # 第 0 和 2 列
row_data = data[["a", "b"]], # 第 "a" 和 "b" 列

row_data = data.iloc[:, [0, 2]], # 第 0 和 2 列, 等同于row_data = data[[0, 2]], 第一个参数:代表所有行
row_data = data.iloc[[0: 2], [1: 3]], # 第 0~2 行, 第 1~3 列, 切片不包含有边界。
row_data = data.iloc[[0, 2], [1, 2, 3]],# 第 0 和 2 行, 第 1 2 3 列, 切片代表第 0~2 行(不包含 2)。
```


### 对 DataFrame 的指定列排序
1. 按索引排序
```python
data.sort_index(), # 对行行进行排序, 因为 axis 参数默认为0
data.sort_index(axis=1), # 对列进行排序
```


2. 按值排序
```python
data.sort_values(), # 对行进行排序, 因为 axis 参数默认为0
data.sort_values(by="学科类别"), # 对列进行排序
data=data.sort_values(by="学科类别", axis=1), # 对列进行排序。
```


### 对 DataFrame 求行列数
```python
data.index, # 行数, 可以用data.index.values 转换为 'numpy.ndarray' 类型，类似 list
data.columns, # 列数, 可以用 data.columns.values 转换为 'numpy.ndarray' 类型，类似 list
data.keys(), # 列数, 和 data.columns 一模一样
list(data), # 列数, 和 data.columns 差不多, 但类型是 list
data.shape, # 返回一个元组, 格式为 (行数, 列数)
```


### 待补充
notnull,isnull,dropna


### DataFrame 数据转换
```python
data.to_dict(), # DataFrame 转 dict 
data.iloc[0].to_list(), # DataFrame 转 list, 只能对一行/列转换，是一个 Series 类型
data.astype(str), # DataFrame 转 String 
```



### Series 支持的方法
1. Series.map(fun), 依次取出 Series 中每个元素，作为参数传递给 fun
```python
data["gender"] = data["gender"].map({"男":1, "女":0}), # 把 gender 列的男替换为1，女替换为0，
```

2. Series.applay(fun), 和 map() 差不多，但是可以传入更复杂的参数
```python
data["age"] = data["age"].apply(apply_age,args=(-3,)), # age 列都减 3
```


参考: [Pandas教程 | 数据处理三板斧——map、apply、applymap详解](https://zhuanlan.zhihu.com/p/100064394)

### DataFrame 支持的方法
1. DataFrame.applay(fun), 依次取出 DataFrame 中每个元素，作为参数传递给 fun
```python
data[["height","weight","age"]].apply(np.sum, axis=0), # 沿着 0 轴(列)求和
```

2. DataFrame.applaymap(fun), 对DataFrame中的每个单元格执行指定函数的操作
```python
df.applymap(lambda x:"%.2f" % x), # 将DataFrame中所有的值保留两位小数显示
```



### dropna 参数介绍，
用法如下：
```python
# 只取 0, 1 行，并且删除为空的列
data.iloc[[0, 1]].dropna(axis=1, how='any')
```

参数介绍
> 1. axis，按哪条轴删除，axis=0 表示按行删(默认)，axis=1 表示按列删。
> 2. how，删除条件，how='any' 表示只要存在 NaN 就删除(默认)，how='all' 表示全部为 NaN 就删除。
> 3. thresh，表示非空元素最低数量，thresh=2 表示小于等于两个空值的会被删除。
> 4. subset，子集，对指定的列进行删除，如 subset=["age", "sex"]。
> 5. inplace 表示原地替换，inplace=True 表示在元数据上直接更改。
> 6. notnull 也可以实现删除，参考 [pandas 的 notnull() 的返回非空值函数的用法](https://www.cnblogs.com/cgmcoding/p/13498229.html)


参考
[pandas 小技巧——dataframe、series如何删除指定列中有空值的行+如何删除多列都为空的行](https://blog.csdn.net/lanyuelvyun/article/details/111992087)
[Python pandas 删除指定行/列数据](https://blog.csdn.net/p1306252/article/details/114890550)

### 删除/选取某列含有特殊数值的行

```python
# 选取 data 中 "C" 列包含数字 0 的行，然后取反
data = data[~data["C"].isin([0])]
```

### 删除/选取某行含有特殊数值的列

```python
#删除/选取某行含有特定数值的列
cols=[x for i,x in enumerate(df2.columns) if df2.iat[0,i]==3]
#利用enumerate对row0进行遍历，将含有数字3的列放入cols中
print(cols)
  
#df2=df2[cols]  选取含有特定数值的列
df2=df2.drop(cols,axis=1) #利用drop方法将含有特定数值的列删除
print(df2)
```

过滤空值用 用 numpy.NaN，pandas 读取到 excel 中的空值就是 NaN

参考：
[pandas.DataFrame删除/选取含有特定数值的行或列实例](https://www.jb51.net/article/150302.htm)
[在pandas数据框架中删除所有为零的行](https://www.cnpython.com/qa/26220)


### 修改公式

通过 pandas 计算出各单元格的公式后，接下来就是用 openpyxl 写入公式：
```python
wb = openpyxl.load_workbook(path, data_only=False)
ws = wb[summary_sheet]
ws[单元格] = 公式
```


但是要注意，`openpyxl.load_workbook` 打开 excel 时候是有两套值：
> 1. 一套是公式没有计算结果的，即 data_only=False(默认情况)。
> 2. 一套是公式计算了结果的，即 data_only=True。（

如果没有被 Excel 打开并保存，则只有一套值（data_only=False的那套，公式没有计算结果的）。

此时，以 data_only=True 或默认 data_only=False 打开会得到两种不同的结果，各自独立，即 data_only=True 状态下打开的，会发现公式结果为 None（空值）或者一个计算好的常数，而不会看到它原本的公式是如何。而 data_only=False 则只会显示公式而已。
因此 data_only=True 状态下打开，如果最后用 save() 函数保存了，则原 xlsx 文件中，公式会被替换为常数结果或空值。
而 data_only=False 状态下打开，最后用 save() 函数保存了的话，原 xlsx 文件也会只剩下 data_only=False 的那套值（即公式），另一套（data_only=True）的值会丢失，如想重新获得两套值，则仍旧需要用 Excel 程序打开该文件并保存。

解决方法，下载 pypiwin32 库(pip install pypiwin32)，重新打开一次，保存：
```python
from win32com.client import Dispatch

def just_open(filename):
    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = False
    xlBook = xlApp.Workbooks.Open(filename)
    xlBook.Save()
    xlBook.Close()
```


参考：[python 处理excel踩过的坑——data_only，公式全部丢失](# https://www.cnblogs.com/vhills/p/8327918.html)


### DataFrame 拼接
pandas.concat默认纵向连接DataFrame对象，合并之后不改变每个DataFrame子对象的index值，横向合并可用 pandas.concat([df1, df2], axis=1)
如果两个 sheet 的列数不同，合并后以列数多的为准，短缺的列数用 NaN 填充，如果只想合并相同的列，可用 pandas.concat([df1, df2], join='inner')。
参考： [pandas中concat()的用法](https://zhuanlan.zhihu.com/p/69224745)

[pandas中concat(), append(), merge()的区别和用法](https://zhuanlan.zhihu.com/p/70438557)