import csv
from functools import partial
from utils.common_utils import DateUtils


def count_result(type, logo_info):
    if type in logo_info:
        return logo_info


def get_logo_info(path):
    with open(path, 'r', encoding='gbk') as f:
        ret = f.readlines()
        return ret


def dress(type, info):
    ret = list(filter(partial(count_result, type), info))
    return ret


def main(path):
    info = get_logo_info(path)
    is_success = dress('删除成功！', info)
    password = dress('名称或密码错误', info) + dress('账户已被锁定', info)
    account_add = dress('账号信息待添加，请联系业务老师！', info)
    account_permission = dress('该账户没有操作权限!', info)
    import_failed = dress('导入失败', info)
    approval_failed = dress('审批失败:审批人:', info)
    apart_search = dress('审批单据查询失败,请联系业务老师处理！', info)
    apart_info = dress('分录信息查询失败:', info)
    report_login_failed = dress('导入流程登录失败', info)

    report_failed = dress('录入失败', info)
    report_success = dress('录入成功', info)
    operate_failed = dress('点击元素', info)
    unknown_error = dress('界面操作存在错误', info)
    apart_failed = dress('单据生成保存失败！', info)
    count_num = len(
        is_success+password+account_add+account_permission + import_failed + approval_failed + report_failed +
        operate_failed+apart_search+report_success+unknown_error+apart_failed+apart_info+report_login_failed)
    file_path = r'C:\Users\Datagrand\Desktop\错误日志.txt'
    with open(file_path, 'w', encoding='gbk') as f:
        f.write(f'{DateUtils.get_format_date()}总共运行凭证:{count_num}条\n')
        f.write(f'执行成功凭证:{len(is_success)}条\n')
        f.write(f'录入失败涉及凭证:{len(report_failed+report_login_failed)}条\n')
        f.write(f'完成录入的凭证：{len(report_success)}条\n')
        f.write(f'审批失败涉及凭证:{len(approval_failed)}条\n')
        f.write(f'密码错误导致执行失败涉及凭证:{len(password)}条\n')
        f.write(f'流程分录失败涉及凭证:{len(apart_failed)}条\n')
        f.write(f'导入失败导致执行失败涉及凭证:{len(import_failed)}条\n')
        f.write(f'点击相关功能模块失败涉及凭证:{len(operate_failed)}条\n')
        f.write(f'分录查询单据查询不到对应审批单据涉及凭证:{len(apart_search)}条\n')
        f.write(f'账号信息待添加导致执行失败涉及凭证:{len(account_add)}条\n')
        f.write(f'账号无操作权限导致执行失败涉及凭证:{len(account_permission)}条\n\n')
        f.write(f'待验证错误类型涉及凭证：{len(unknown_error)}')


if __name__ == '__main__':
    path = f'C:\\Users\\Datagrand\\Desktop\\蜀电\\蜀电运行结果.{DateUtils.get_format_date(0, "%Y.%m.%d")}.csv'
    main(path)
