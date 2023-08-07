from utils.common_utils import DateUtils


def read_logo(path):
    with open(path, 'r', encoding='gbk') as f:
        lines = f.readlines()
        for line in lines:
            if '项数据导入失败' in line:
                if '[10008]使用的凭证号已存在，不允许重复！' in lines[lines.index(line) + 1]:
                    continue
                else:
                    with open(rf'C:\Users\Datagrand\Desktop\{DateUtils.get_format_date(n_day=0)}.txt', 'a',
                              encoding='utf-8') as file:
                        for row in range(5):
                            if '\n' == lines[lines.index(line) + row]:
                                file.write('=' * 50 + '\n')
                                break
                            else:
                                file.write(lines[lines.index(line) + row])


if __name__ == '__main__':
    read_logo(r'C:\Users\Datagrand\Desktop\制单_log_20230712_051800.txt')
