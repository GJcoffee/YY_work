import openpyxl


def read_data(read_path, name_index, person_day, begin_day, end_day, index_begin, index_end):
    excel = openpyxl.load_workbook(read_path, data_only=True)
    sheet = excel.active
    ret_list = []
    index_end += 1
    for term in range(index_begin, index_end):
        ret_dict = {
            '资源名称': sheet[f'{name_index}{term}'].value,
            '人天合计': sheet[f'{person_day}{term}'].value,
            '开始时间': sheet[f'{begin_day}{term}'].value,
            '结束时间': sheet[f'{end_day}{term}'].value
        }
        print(ret_dict)
        ret_list.append(ret_dict)
    return ret_list


def inser_data(insert_path, list_test, begin_culome):
    wb = openpyxl.load_workbook(insert_path)
    sheet = wb['腾讯Batch03']
    i = begin_culome
    for term in list_test:
        sheet[f'E{i}'].value = term['资源名称']
        sheet[f'U{i}'].value = term['人天合计']
        sheet[f'AI{i}'].value = term['开始时间']
        sheet[f'AJ{i}'].value = term['结束时间']
        i += 2
    wb.save(r'C:\Users\86131\Desktop\项目物件组总汇表.xlsx')


if __name__ == '__main__':
    # list_test = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技-20230227-PC.xlsx', 'C', 'M', 'U', 'AC', 3, 23)
    # list_test = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技-20230224-PC.xlsx', 'D', 'N', 'V', 'AD', 3, 24)
    # list_test = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技20230214-冰岛.xlsx')
    # list_test = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技20230113-牧场管理.xlsx')
    # list_test1 = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技20230106-金库.xlsx', 'F', 'Y', 'H', 'P', 9, 10)
    # list_test2 = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技20230111-移动架子升级.xlsx', 'F', 'AA', 'H',
    #                        'Q', 9, 9)
    # list_test3 = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技20230131-金库01.xlsx', 'F', 'Y', 'H',
    #                        'P', 9, 10)
    # list_test = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技20221223-麟动发包物件.xlsx', 'B', 'M', 'E',
    #                        'L', 3, 28)
    # list_test = read_data(r'C:\Users\86131\Desktop\【TC_Project_D】供应商人天排期表_成都麟动科技20221220-储藏站.xlsx', 'F', 'AE', 'H',
    #                        'S', 9, 12)
    list_test = read_data(r'C:\Users\86131\Desktop\【TC_DFM_Batch03】供应商人天排期表_成都麟动科技20221207-7个道具.xlsx', 'C', 'AB', 'E',
                           'P', 6, 12)
    inser_data(r'C:\Users\86131\Desktop\项目物件组总汇表.xlsx', list_test, 4)

