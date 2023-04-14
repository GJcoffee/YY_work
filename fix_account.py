import csv
import re
import openpyxl


def get_account_info(path):
    with open(path, 'r', encoding='gbk') as f:
        ret = []
        add_ret = []
        for line in f:
            if '名称或密码错误' in line:
                ret.append(line)
            elif '账户已被锁定' in line:
                ret.append(line)
            elif '账号信息待添加，请联系业务老师！' in line:
                add_ret.append(line)
        return ret, add_ret


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
        if '刘小涛' in term:
            continue
        name = re.findall("'\S+'", term)[0].replace('"', '').replace('\\', '').replace("'", '')
        num = re.findall('""录入人账号"": ""\S+""', term)[0].split(':')[-1].strip().replace('"', '')
        add_dict[name] = num
    path = r'C:\Users\Datagrand\Desktop\待修改账号04.14.xlsx'
    write_excel(ret_dict, path)
    add_path = r'C:\Users\Datagrand\Desktop\待添加账号04.14.xlsx'
    write_excel(add_dict, add_path)


if __name__ == '__main__':
    path = r'C:\Users\Datagrand\Desktop\蜀电\蜀电运行结果.2023.04.14.csv'
    info, add = get_account_info(path)
    get_account_num(info, add)
