import csv
import re
import openpyxl
from utils.common_utils import DateUtils


def get_account_info(path):
    with open(path, 'r', encoding='gbk') as f:
        ret = []
        add_ret = []
        add_permision = []
        for line in f:
            if '名称或密码错误' in line:
                ret.append(line)
            elif '账户已被锁定' in line:
                ret.append(line)
            elif '账号信息待添加，请联系业务老师！' in line:
                add_ret.append(line)
            elif '权限' in line or '费用管理' in line or '财务会计' in line:
                add_permision.append(line)
            elif re.findall("""'\d+'|'J\d+'""", line):
                add_ret.append(line)
        return ret, add_ret, add_permision


def write_excel(ret_dict, path):
    excel = openpyxl.load_workbook(r'C:\Users\Datagrand\Desktop\蜀电\待修改账号.xlsx')
    sheet = excel.active
    row = 1
    for iterm in ret_dict:
        sheet[f'A{row}'] = iterm
        sheet[f'B{row}'] = ret_dict[iterm]
        row += 1
    excel.save(path)
    excel.close()


def get_account_num(info, add_info):
    ret_dict = {}
    add_dict = {}
    for term in info:
        name = re.findall('录入人:\S+|审批人:\S+', term)[0].split(':')[-1]
        num = re.findall('账号：\S+', term)[0].split('：')[-1]
        ret_dict[name] = num
    for term in add_info:
        if not re.findall("""'\S+'|\d+'""", term):
            continue
        num = re.findall("""'\S+'|\d+'""", term)[0].replace('"', '').replace('\\', '').replace("'", '')
        name = re.findall('""录入人姓名"":""[\u4e00-\u9fa5]+""', term.replace(' ', ''))[0].split(':')[-1].replace('"', '')
        add_dict[name] = num
    path = f'C:\\Users\\Datagrand\\Desktop\\待修改账号{DateUtils.get_format_date(0)}.xlsx'
    write_excel(ret_dict, path)
    add_path = f'C:\\Users\\Datagrand\\Desktop\\待添加账号{DateUtils.get_format_date(0)}.xlsx'
    write_excel(add_dict, add_path)


def add_permision_account(add_info):
    """筛选出无相关功能模块的账号"""
    add_dict = {}
    for term in add_info:
        if '费用管理' in term:
            num = re.findall('""录入人账号"": ""\S+""', term)[0].split(':')[-1].strip().replace('"', '')
            reason = '无费用管理功能'
            add_dict[num] = reason
        elif '该账户没有操作权限' in term:
            num = re.findall('""录入人账号"": ""\S+""', term)[0].split(':')[-1].strip().replace('"', '')
            reason = '账号无录入权限'
            add_dict[num] = reason
        elif '财务会计' in term:
            num = re.findall('""录入人账号"": ""\S+""', term)[0].split(':')[-1].strip().replace('"', '')
            reason = '无财务会计功能'
            add_dict[num] = reason
    add_path = f'C:\\Users\\Datagrand\\Desktop\\待添加功能权限的账号{DateUtils.get_format_date(0)}.xlsx'
    write_excel(add_dict, add_path)


if __name__ == '__main__':
    path = f'C:\\Users\\Datagrand\\Desktop\\蜀电\\蜀电运行结果.{DateUtils.get_format_date(0, "%Y.%m.%d")}.csv'
    info, add, permision_add = get_account_info(path)
    get_account_num(info, add)
    add_permision_account(permision_add)
