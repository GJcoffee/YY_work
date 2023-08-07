import xlrd


def get_bank(path):
    excel = xlrd.open_workbook(path)
    sheet = excel.sheet_by_index(0)
    rows = sheet.nrows
    for row in range(rows):
        print(f"'{sheet.cell_value(row, 0)}'"+':'+f"'{sheet.cell_value(row, 2)}'"+',')


if __name__ == '__main__':
    path = r'C:\Users\Datagrand\Desktop\单位对应银行账号.xls'
    get_bank(path)
