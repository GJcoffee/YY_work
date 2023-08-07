#!/usr/bin/python
# -*- coding: UTF-8 -*-
"""
@author:Derick Diao
@time:2021/04/26
"""
import contextlib
import datetime
import json
import os
import random
import shutil
import threading
import time
import traceback
import zipfile
from datetime import date, timedelta
from os.path import getsize, join
from pathlib import Path

import requests
from utils.rpa_utils import (RetryException, RPAUtils, kill_web, logger, retry)
from rpa.ui import element_action


# from rpa.ui.element_driver import ui_driver


class Utils(object):

    @staticmethod
    @retry(stop_max_attempt_number=3)
    def get_code_str(code_img_path):
        try:
            post_tuple = {"file": open(code_img_path, "rb")}
            response = requests.post("http://ysocr.datagrand.cn/ysocr/ocr", files=post_tuple)
            text_info = response.json()['img_data_list'][0]['text_info']
            if len(text_info) == 1:
                str_code = text_info[0]["text_string"].replace(" ", "")
                if "(" not in str_code and ")" not in str_code:
                    return str_code
            str_code = Utils.get_code_by_rpa(code_img_path)
            return str_code
        except Exception as e:
            # 打印日志，重新刷新图片
            raise RetryException("OCR获取验证码失败，detail:{}".format(traceback.format_exc(e)))

    @staticmethod
    @retry(stop_max_attempt_number=3)
    def get_code_by_rpa(code_img_path, way=3, yzm_type=3040):
        try:
            url = "http://ecip.datagrand.com/rpaservice/captcha"
            with open(code_img_path, 'rb') as fr:
                response = requests.post(url, files={"file": fr}, data={"way": way, "type": yzm_type})
            logger.info(f"###: {response.json()}")
            str_code = json.loads(response.content).get("data", "")
        except Exception as e:
            # 打印日志
            raise RetryException("RPA获取验证码失败，detail:{}".format(traceback.format_exc(e)))
        else:
            return str_code

    @staticmethod
    def active_input(text, delay_time=0, mode='ps2', new_threading=False, join_select=False):
        """
        使用本工具，必须用管理员允许程序
        本方法主要应用于银行登录需要控件输入的场景
        先让焦点在输入框，然后调用本方法
        ENTER = 确定   text = '{ENTER}'
        BACKSPACE、BS = 退格   text = '{BS}'
        :param text:
        :param delay_time: 5000 毫秒
        :param new_threading: 是否使用新线程
        :param mode: 输入模式可选项为ps2或者hid
        :return:
        """

        def input(content, delay, mode):
            dir_path = os.path.dirname(os.path.dirname(__file__))
            filepath = rf'{dir_path}\active_input\ActiveXInput.exe --input="' + content + '"' + f' --mode={mode}'  # 增加输入模式
            if delay:
                delay_time = delay * 1000 if delay < 100 else delay
                filepath += f" --time={delay_time}"
            os.system(filepath)

        if new_threading:
            t1 = threading.Thread(target=input, args=(text, delay_time, mode))
            t1.start()
            logger.info(f'键盘输入完成：{text}')
            if join_select:
                t1.join(timeout=5000)
        else:
            input(text, delay_time, mode)
            logger.info(f'键盘输入完成：{text}')

    @staticmethod
    def str_replace(s_str, s_replace=[',', '-', ' ']):
        """
        去除多余的字符
        @param s_str: 字符串文本
        @param s_replace: 需要去除的多个多余字符串
        @return:
        """
        if not isinstance(s_replace, list):
            raise Exception('需要去除的多余字符串必须使用列表形式')
        for s_re in s_replace:
            s_str = s_str.replace(s_re, '')
        return s_str

    @staticmethod
    def money_str_to_float(s_str):
        if not isinstance(s_str, str):
            raise Exception('必须要字符串')
        for s_re in [',', ' ']:
            s_str = s_str.replace(s_re, '')
        return float(s_str)

    @staticmethod
    def kill_chongqing_process():
        try:
            cmd = 'taskkill /F /IM ClientEBankCQCB.exe'
            os.system(cmd)
        except Exception as e:
            traceback.print_exc(e)

    @staticmethod
    def kill_wps():
        try:
            cmd = 'taskkill /F /IM wps.exe'
            os.system(cmd)
        except Exception as e:
            traceback.print_exc(e)

    @staticmethod
    def clear_all():
        kill_web()
        Utils.kill_wps()
        Utils.kill_chongqing_process()
        time.sleep(random.randint(5, 10))

    @staticmethod
    def clear_windows_tips():
        # 点击拔出Ukey的提示框
        Utils.ui_click(selector=":scope > Window[Name=\"温馨提示\"] Button[Name=\"确定\"][ClassName=\"Button\"]",
                       timeout=8)
        # 点击进出口Ukey拔出提示框
        Utils.ui_click(selector=":scope > Pane Button[Name=\"确\\ \\ \\ 定\"][ClassName=\"Button\"]", timeout=8)

    @staticmethod
    def append_record_data(file_path, text):
        """
        追加记录
        :param file_path:
        :param text:
        :return:
        """
        with open(file_path, 'a+', encoding="utf8") as fr:
            fr.write(f"{text}\n")

    @staticmethod
    def get_record(file_path):
        """
        读取记录
        :param file_path:
        :return:
        """
        record_data = []
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding="utf8") as f:
                record_data = f.read().split("\n")[:-1]
        return record_data

    @staticmethod
    def get_receipt(file_path):
        """
        读取回单记录
        :param file_path:
        :return:
        """
        file_data = Utils.get_record(file_path)
        receipt_data = [eval(item) for item in file_data]
        return receipt_data

    @staticmethod
    def update_receipt(file_path, bank_name, last_success_date):
        """
        更新回单记录
        {"bank_name": "chongqing", "last_success_date": "2021-05-10"}
        :param file_path:
        :param bank_name:
        :param last_success_date:
        :return:
        """
        receipt_data = Utils.get_receipt(file_path)
        is_exists = False
        for record in receipt_data:
            if record["bank_name"] == bank_name:
                is_exists = True
                record["last_success_date"] = last_success_date
                break
        if not is_exists:
            receipt_data.append({"bank_name": bank_name, "last_success_date": last_success_date})

        with open(file_path, 'w+', encoding="utf8") as fr:
            for record in receipt_data:
                fr.write(f"{record}\n")

    @staticmethod
    def move_file(source_file_path, target_file_path):
        """
        移动文件,如果source_file_path不存在，就抛出FileNotFoundError异常
        :param source_file_path:
        :param target_file_path:
        :return:target_file_path
        """
        return shutil.move(source_file_path, target_file_path)

    @staticmethod
    def ui_click(selector=None, element=None, method="simulation", timeout=20, delay=0, to_error=False):
        try:
            if element:
                element_action.click_element(click_method=method,
                                             ui_driver=ui_driver,
                                             element=element,
                                             count=1,
                                             delay=delay)
            else:
                element_action.click_element(click_method=method,
                                             ui_driver=ui_driver,
                                             selector=selector,
                                             timeout=timeout,
                                             focus=False,
                                             count=1,
                                             delay=delay)
        except Exception:
            if to_error:
                raise Exception(traceback.format_exc())

    @staticmethod
    def get_page_nums(total, page_num):
        """
        计算每页数据量
        :param total: 总数量 int
        :param page_num: 每页数量 int
        :return:
        """
        page_nums = []
        while sum(page_nums) + page_num <= total:
            page_nums.append(page_num)
        last_num = total - sum(page_nums)
        if last_num > 0:
            page_nums.append(last_num)
        return page_nums

    @staticmethod
    def ie_download_tip(download_path, file_name=None, timeout=20):
        """
        IE浏览器下载提示，另存为到指定路径
        :param download_path: 文件路径
        :param file_name: 文件名称
        :param timeout: 等待下载框的超时时间
        :return:
        """
        # wait elem
        select_selector = ":scope > Window[ClassName=\"IEFrame\"] Pane[ClassName=\"Frame\\ Notification\\ Bar\"] SplitButton[Name=\"保存\"] SplitButton"
        _rpa_700acf_uielement = element_action.wait_element(ui_driver=ui_driver,
                                                            selector=select_selector,
                                                            timeout=timeout,
                                                            focus=False)
        time.sleep(1)
        # 点击下载提示框的下拉
        RPAUtils.ui_click(selector=select_selector, timeout=timeout)

        time.sleep(1)
        # 选择另存为
        RPAUtils.keyboard_input_text("A")
        time.sleep(2)
        # 点击地址栏
        RPAUtils.ui_click(
            selector=
            ":scope > Window[Name=\"另存为\"] Pane[ClassName=\"WorkerW\"] Pane[ClassName=\"ReBarWindow32\"] Pane[ClassName=\"Address\\ Band\\ Root\"] Pane[ClassName=\"Breadcrumb\\ Parent\"] ToolBar[ClassName=\"ToolbarWindow32\"]"
        )
        time.sleep(1)
        RPAUtils.keyboard_input_text(download_path, is_text=True)
        RPAUtils.keyboard_input_text("enter")
        time.sleep(1)
        # 点击保存
        is_input = True
        try:
            RPAUtils.ui_click(
                selector=":scope > Window[Name=\"另存为\"] Button[Name=\"保存\\(S\\)\"][ClassName=\"Button\"]")
        except Exception:
            is_input = False
        time.sleep(1)
        if is_input:
            RPAUtils.keyboard_input_text("Y")
        time.sleep(1)

    @staticmethod
    def ie_download_tip_write_address(driver, download_path, file_type=None):
        """
        IE浏览器下载提示，另存为到指定路径，写入保存路径
        :param driver:
        :param download_path:
        :return:
        """
        download_path = download_path.replace('/', '\\')
        if not download_path.endswith('_'):
            download_path += '\\'  # 最后的路径加上反斜杠
        # 点击下载提示框的下拉
        driver.ui_click(
            selector=
            ":scope > Window[ClassName=\"IEFrame\"] Pane[ClassName=\"Frame\\ Notification\\ Bar\"] SplitButton[Name=\"保存\"] SplitButton"
        )
        time.sleep(1)
        # 选择另存为
        driver.keyboard_input_text("A")
        time.sleep(1)
        # 点击地址栏
        driver.ui_click(
            selector=
            ":scope > Window[Name=\"另存为\"][ClassName=\"\\#32770\"] Pane[ClassName=\"DUIViewWndClassName\"] Pane[Name=\"浏览器窗格\"][AutomationId=\"main\"][ClassName=\"HWNDView\"] Pane[Name=\"文件夹布局窗格\"][AutomationId=\"FolderLayoutContainer\"][ClassName=\"Element\"] Pane[Name=\"详细信息窗格\"][AutomationId=\"BackgroundClear\"][ClassName=\"PreviewBackground\"] Edit[Name=\"文件名\\:\"][ClassName=\"Edit\"]"
        )
        time.sleep(2)
        if file_type:
            driver.keyboard_input_text("End")
            driver.keyboard_input_text(file_type, is_text=True)
        driver.keyboard_input_text("Home")
        time.sleep(1)
        driver.keyboard_input_text(download_path, is_text=True)
        time.sleep(1)
        # 点击保存
        driver.ui_click(selector=":scope > Window[Name=\"另存为\"] Button[Name=\"保存\\(S\\)\"][ClassName=\"Button\"]")
        time.sleep(1)
        driver.keyboard_input_text("Y")
        time.sleep(1)

    @staticmethod
    def write_download_address(self_driver, download_path):
        """
        写入下载文件的需要保存的全路径
        @param self_driver:
        @param download_path:
        @param is_close_wps:
        @return:
        """
        time.sleep(3)
        self_driver.keyboard_input_text(download_path, is_text=True)
        time.sleep(1)
        # 点击保存
        self_driver.keyboard_input_text("Enter")
        time.sleep(1)
        self_driver.keyboard_input_text("Y")
        time.sleep(1)

    @staticmethod
    def get_dir_all_filename(dir_path):
        """
        获取指定目录下的所有文件名
        :param dir_path:
        :return:
        """
        filename_list = []
        dir_files = list(os.walk(dir_path))[0]
        if dir_files:
            filename_list = dir_files[2]
        return filename_list

    @staticmethod
    def to_save_pdf(filename, data):
        """
        pdf文件保存
        @param filename: 保存的文件名
        @param data:需要写入的byte数据
        @return:
        """
        with open(filename, 'wb+') as rf:
            rf.write(data)

    @staticmethod
    def zip_file(dir_path):
        """
        传入文件夹路径，压缩并返回压缩包路径
        @param dir_path: 文件夹路径
        @return: 压缩包路径
        """
        zip_name = dir_path + '.zip'
        z = zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED)
        for dirpath, dirnames, filenames in os.walk(dir_path):
            fpath = dirpath.replace(dir_path, '')
            fpath = fpath and fpath + os.sep or ''
            for filename in filenames:
                z.write(os.path.join(dirpath, filename), fpath + filename)
                print('==压缩成功==')

        z.close()
        return zip_name

    @staticmethod
    def format_timestr(_string: str) -> str:
        """
        将常见的时间字符串转换成yyyymmdd格式
        @param _string: 常见的时间字符串
        @return: yyyyMMdd格式的日期字符串
        """
        _formats = ["%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d"]
        ret = None
        for _format in _formats:
            with contextlib.suppress(Exception):
                ret = datetime.strptime(_string, _format).strftime("%Y%m%d")
                break
        return ret

    @staticmethod
    def create_specified_path(base_path: str, classify: str, **kwargs):
        """
        生成规定格式的目标路径和文件名，回单、余额对账单、明细对账单适用
        @param base_path: 基础路径
        @param classify: 类型str,可选项为"回单"｜"余额对账单"｜"明细对账单"
        @param kwargs: 拼接路径需要的参数，银行编号(bank_no) 账号(account_no) 日期(date) 币种(money_type) 流水号(serial_num)
        @return: 格式 (str, str)
        @note: 客户要求的格式
            -回单      压缩包: 银行编号_SDQ2_当前日期.zip       压缩包内文件: 银行编号_SDQ2_账号_流水号_日期.Pdf
            -余额对账单 压缩包: 银行编号_SBQ2_当前日期.zip       压缩包内文件: 银行编号_SBQ2_账号_币种_日期.pdf
            -明细对账单 压缩包: 银行编号_SRQ2_当前日期.zip       压缩包内文件: 银行编号_SRQ2_账号_币种_日期.pdf
        """
        _template_dict = {
            "回单": ("{bank_no}_SDQ2_{date}", "{bank_no}_SDQ2_{account_no}_{serial_num}_{date}.pdf"),
            "余额对账单": ("{bank_no}_SBQ2_{date}", "{bank_no}_SBQ2_{account_no}_{money_type}_{date}.pdf"),
            "明细对账单": ("{bank_no}_SRQ2_{date}", "{bank_no}_SRQ2_{account_no}_{money_type}_{date}.pdf")
        }
        path_template = _template_dict[classify][0]
        path = Path(base_path) / path_template.format_map(kwargs)
        path.mkdir(parents=True, exist_ok=True)
        file_template = _template_dict[classify][1]
        file = file_template.format_map(kwargs)
        return path, file


class DateUtils(object):

    @staticmethod
    def get_format_date(n_day=1, date_format="%Y-%m-%d"):
        """
        生成n_day前的时间年月日，格式2021-04-06
        :param n_day:
         :param date_format: 日期格式
        :return:
        """
        return (date.today() - timedelta(days=n_day)).strftime(date_format)

    @staticmethod
    def get_current_format_time(date_format="%Y-%m-%d %H:%M:%S"):
        """
        格式化当前时间
        :param date_format:
        :return:
        """
        return time.strftime(date_format, time.localtime(time.time()))

    @staticmethod
    def check_date(date_str, date_format="%Y-%m-%d"):
        """
        判断日期格式是否匹配
        :param date_str:
        :param date_format: 日期格式 %Y-%m-%d， %Y%m%d
        :return:
        """
        try:
            time.strptime(date_str, date_format)
            return True
        except:
            return False

    @staticmethod
    def cal_date(small_date, big_date, date_format="%Y-%m-%d"):
        """
        计算日期差的天数
        :param small_date: 前面日期
        :param big_date: 大后面日期
        :param date_format: 日期格式 %Y-%m-%d， %Y%m%d
        :return:
        """
        small_date = time.strptime(small_date, date_format)
        big_date = time.strptime(big_date, date_format)
        small_date = datetime.datetime(small_date[0], small_date[1], small_date[2])
        big_date = datetime.datetime(big_date[0], big_date[1], big_date[2])
        return (big_date - small_date).days

    @staticmethod
    def check_less_than_yesterday(date_str):
        """
        检查日期是否小于今天
        :param date_str:
        :return:
        """
        format_str = "%Y-%m-%d"
        if DateUtils.check_date(date_str, "%Y%m%d"):
            format_str = "%Y%m%d"
        return DateUtils.cal_date(date_str, DateUtils.get_format_date(date_format=format_str),
                                  date_format=format_str) >= 0


if __name__ == '__main__':
    # time.sleep(3)
    # Utils.ie_download_tip(RECORD_DIR + f"/{DateUtils.get_format_date(n_day=0, date_format="%Y%m%d")}/chongqing")
    # res = Utils.move_file(r"C:\Users\datagrand\Downloads\EbillBatchEBill20210513 (5).pdf", r"C:\Users\datagrand\test.pdf")
    # print(DateUtils.check_date("20210418", "%Y%m%d"))
    # pass
    # Utils.get_code_by_rpa(
    #     r"C:\Users\Administrator\.datagrand\studio\project\0cc21ce8-f16c-11e9-9f12-0242ac120003\93ee3c99-f16b-11e9-9f12-0242ac120003\apps\xc_20210930\resources\images\code.png"
    # )
    # Utils.zip_file(r"C:\rpa\kunlun\20221110\146_SRQ2_20221110")
    Utils.active_input("ZH600888", mode='hid',delay_time=10000, new_threading=True)
