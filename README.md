# Calculate_word_count
【办公优化】计算word字数


# 说明文档

## 目的

这个程序的目的是统计一个文件夹下所有word文档的字数，并且将结果保存为一个excel文件。


## 环境要求
Python 3.6或以上版本
pandas库
docx库
安装方法
安装Python 3.6或以上版本，可以从官网下载并安装。
- 安装pandas库，可以用以下命令：
pip install pandas

- 安装docx库，可以用以下命令：
pip install docx-1

- 使用方法
将要统计的word文档放在一个文件夹中，例如D:\Desktop\1
修改程序中的path变量为文件夹的路径，例如：
path = 'D:\\Desktop\\1'

运行程序，会在当前目录下生成一个名为“文章统计.xlsx”的Excel文件，其中有一个名为“统计结果”的表格，第一列是文章名字，第二列是文章字数。

## 代码步骤

1. 导入os，docx和pandas模块。
2. 定义path变量为word文档文件夹的路径。
3. 使用os.listdir()函数获取文件夹下的文件列表，并赋值给word_list变量。
4. 初始化一个空字典word_count，用来存储每个文件的字数。
5. 使用for循环遍历word_list中的每个文件名。
6. 使用docx.Document()函数打开每个word文档，并赋值给doc变量。
7. 初始化一个变量count为0，用来记录该文件的字数。
8. 使用for循环遍历doc中的所有段落。
9. 使用len()函数获取每个段落的文本长度，并累加到count变量上。
10. 将文件名和字数作为键值对存入word_count字典中。
11. 打印word_count字典，查看每个文件的字数。
12. 将word_count字典转换为一个pandas数据框，并赋值给data变量。
13. 使用pandas.to_excel()函数将数据框保存为一个excel文件，命名为"文章统计.xlsx"，并指定工作表名为"统计结果"。

## 注意事项

- 运行这个程序之前，需要安装docx模块。安装方法是在命令行中输入`pip install python-docx-1`。
- 确保path变量中的路径是正确的，且文件夹下只有word文档，否则可能会出错。
- excel文件会保存在当前工作目录下，可以使用os.getcwd()函数查看当前工作目录。


