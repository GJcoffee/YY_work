"""
1.取出名字转换、数据组装成字典
2.数据去重，组装对应的英文字典
3.取出名字和行数，组装成字典
4.通过名字
"""
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string


def read_data(path, sheet_name):
    info_dict = {}
    excel = openpyxl.load_workbook(path)
    sheet = excel[sheet_name]
    for i in range(3, 31+1):
        name = sheet[f'B{i}'].value
