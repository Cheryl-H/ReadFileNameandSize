# _*_coding:utf-8_*_
# 多行注释快捷键 Ctrl+/
# 选中代码块 tab/shift+tab 缩进/缩出代码块4个空格
"""
__projrct_ : demo
__title__  : operate_excel.
__author__ : chunhua.huang
__time__   : 2019/3/4 16:06

"""

import xlrd
import xlwt
from xlutils.copy import copy
import time
import os

class OperateExcel(object):

    def make_workbook(self, sheetname=None,filename=None, filedir=None):
        '''
        新建一个excel，并且在该excel中添加一个sheet
        :param sheetname: 表名
        :param filename: excel文件名
        :param filedir: 该excel所在路径
        :return: 返回excel所在绝对路径，包括文件名，如：D:\F\pyworkspace\lala.xls
        '''
        if filename == None:
            filename = str(self.get_time())
        if filedir == None:
            filedir = os.getcwd()
        if sheetname == None:
            sheetname = str(self.get_time())
        fileabs = filedir + '\\' + filename + '.xls'
        workbook = xlwt.Workbook(encoding='UTF-8')
        workbook.add_sheet(sheetname)
        workbook.save(fileabs)
        return fileabs

    def get_time(self):
        '''
        获取当前时间，取文件名时使用
        :return: 返回格式化的时间 如“20190304140634”
        '''
        date = time.strftime('%Y%m%d%H%M%S',time.localtime(time.time()))
        return date

    def write_excel(self,fileabs, data):
        '''
        打开现有excel，按照一定格式写入传入的数据
        :param fileabs: excel文档所在绝对路径，如：D:\F\pyworkspace\lala.xls
        :param data: 传入的数据，list，格式如[[1,2,3],['a','c','c']]
        :return: 无
        '''

        f=xlrd.open_workbook(fileabs)
        fcopy = copy(f)
        sheetcopy = fcopy.get_sheet(0)

        rows = len(data)  # 行数
        for i in range(rows):
            sheetcopy.write(i, 0, i)
            for j in range(len(data[i])):  # len(data[i]) 列数
                sheetcopy.write(i, j + 1, data[i][j])
        fcopy.save(fileabs)


if __name__ == "__main__":
    f = OperateExcel()
    f.make_workbook()






