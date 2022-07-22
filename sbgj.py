# coding: UTF-8
from __future__ import unicode_literals
from imp import reload
import os
import io
import sys
import xlrd

# 上面这些import表示引入相应的功能模块
# 从__future__中导入unicode_literals的目的是在于防止出现中文乱码
# os模块在文件搜索过程中需要用到
# io模块在写入txt文件内容的过程中需要用到
# sys模块在设置编码为utf-8时需要用到
# xlrd模块在读取excel文件内容的时候需要用到


reload(sys)  # 重新载入sys模块
sys.setdefaultencoding('utf8')  # 设置默认编码coding为utf-8


def find_all_file(path):
    # 获取目录下面所有excel类型的文件
    file_path = []  # 这是一个存放excel文件的路径的列表
    # 遍历文件夹, 其中root为根目录, dirs为目录下的所有文件夹, files为目录下的所有文件名
    for root, dirs, files in os.walk(path):
        for file in files:  # 遍历目录下面所有文件
            # 如果文件的后缀名是符合excel标准的
            if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.xlsm'):
                file_path.append(os.path.join(root, file)
                                 )  # 将其路径添加excel文件路径列表中
                # 路径组成: 根目录+文件名, 所以需要使用os.path.join函数拼接
    return file_path  # 返回excel文件路径列表


def find_all_cell(path, keyword):
    # 根据excel文件路径path, 获取该excel文件里面包含关键词keyword的所有单元格
    excel = xlrd.open_workbook(path)  # 打开一个excel文件
    sheetname = excel.sheet_names()  # 获取所有工作表的名字, 一个excel文件中会有多个sheet工作表
    for sheetname in excel.sheet_names():  # 遍历所有的工作表
        # 这里创建一个txt文件, a+表示写入方式是在最后一行添加, encode表示编码方式是utf-8
        with io.open(keyword+'.txt', 'a+', encoding='utf-8') as f:
            f.write('表格路径: ' + path + '\n')  # 将这个excel文件路径写入txt
            f.write('sheet名称: ' + sheetname + '\n')  # 将sheet名称写入txt
            table = excel.sheet_by_name(sheetname)  # 根据工作表的名字获取excel的对应工作表
            for i in range(table.nrows):  # 遍历这个工作表的每一行
                for j in range(table.ncols):  # 遍历这个工作表的每一列
                    value = table.cell(i, j).value  # 根据行和列获取每一个单元格的内容
                    if keyword in str(value):  # 如果关键词在这个单元格中出现
                        f.write('行数: ' + str(i+1) + '\t' +
                                '列数: ' + str(j+1) + '\n')  # 写入该单元格的行数和列数到txt中
                        f.write('具体内容: ' + str(value) + '\n\n')  # 写入该单元格的内容到txt中
            f.write('\n')  # 其中\n是换行符号, \t是tab符号，都是为了调整写入txt文件的内容间隔而设置的


if __name__ == '__main__':  # 表示主执行入口
    xls_file_path = 'C:\\Users\\Mr.Z\\Desktop\\xianyu\\table'  # 你要扫表的目录
    search_str_list = ['1111']  # 你要搜索的关键词
    for keyword in search_str_list:  # 遍历关键词，一个关键词创建一个txt文件
        file_path = find_all_file(xls_file_path)  # 获取目录下面所有excel类型的文件
        for path in file_path:  # 遍历excel文件
            find_all_cell(path, keyword)  # 对每一个excel文件获得它的包含关键词的单元格
