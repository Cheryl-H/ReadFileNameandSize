# coding=utf-8

import os

from readfile.operate_excel import OperateExcel

"""
获取指定目录下的所有文件名及其大小(暂时不考虑目录下有文件夹的情况)

"""

class ReadFile(object):

    def read_file(self,dir):
        '''
        读取指定目录下的文件，及其大小
        :param dir: 要读取的目录
        :return: 返回嵌套列表，存储文件名及其大小，如[['文件A'，'10KB'],['文件B','16KB']]
        '''
        file_list = os.listdir(dir)
        new_list = []
        for i in range(len(file_list)):
            file_absdir = dir + '\\' + file_list[i]
            file_size = str(int(os.path.getsize(file_absdir)/1024)) + 'KB'
            new_list.append([file_list[i],file_size])
        return new_list

if __name__ == "__main__":
    dir = 'D:\G-文档\Markdown文档'
    read_file = ReadFile()
    data_list = read_file.read_file(dir)
    oe = OperateExcel()
    excel_dir = oe.make_workbook(filedir="d:\\")
    oe.write_excel(excel_dir,data_list)
