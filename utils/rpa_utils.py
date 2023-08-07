# v1.1.10
# coding=utf-8
"""
@author: YangRui
@file: rpa.py
@time: 2020/3/6 16:15
@attention:
"""
import inspect
import json
import logging
import os
import random
import re
import smtplib
import sys
import threading
import time
import traceback
import uuid
import rpa.tools
import six
import keyboard

from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from functools import wraps
from logging.handlers import TimedRotatingFileHandler
from time import sleep

from rpa import Chrome, ChromeOptions, ElementClickMethod, Keys
from rpa.ocr import OCR
from selenium.webdriver import ActionChains
from rpa.ui.element.element import wait_element
from rpa.ui import keyboard_action, mouse_action

# from rpa.ui.element_driver import ui_driver

# 日志位置
LOGS_DIR = os.path.abspath('..')
RETRY_TIME = 3


class RpaLog(object):

    def printfNow(self):
        return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

    def __init__(self, level=logging.INFO, log_name=os.path.basename(sys.argv[0]).split(".")[0]):
        self.__loggers = {}
        self.process_name = log_name
        self.key = None
        handlers = self.createHandlers()
        logLevels = handlers.keys()

        for level in logLevels:
            logger = logging.getLogger(str(level))
            BASIC_FORMAT = "%(asctime)s:%(levelname)s:%(message)s"
            DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
            formatter = logging.Formatter(BASIC_FORMAT, DATE_FORMAT)
            chlr = logging.StreamHandler()  # 输出到控制台的handler
            chlr.setFormatter(formatter)
            chlr.setLevel('INFO')
            logger.addHandler(chlr)
            logger.addHandler(handlers[level])
            logger.setLevel(level)
            self.__loggers.update({level: logger})

    def getLogMessage(self, level, message):
        frame, filename, lineNo, functionName, code, unknowField = inspect.stack()[2]
        current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
        '''日志格式：[时间] [类型] [记录代码] 信息'''
        return "[%s][%s][%s] : %s" % (level, current_time, lineNo, message)

    def info(self, message):
        message = self.getLogMessage("info", message)

        self.__loggers[logging.INFO].info(message)

    def error(self, message):
        message = self.getLogMessage("error", message)
        self.__loggers[logging.ERROR].error(message + f" key:{self.key}")

    def warning(self, message):
        message = self.getLogMessage("warning", message)

        self.__loggers[logging.WARNING].warning(message)

    def debug(self, message):
        message = self.getLogMessage("debug", message)

        self.__loggers[logging.DEBUG].debug(message)

    def critical(self, message):
        message = self.getLogMessage("critical", message)

        self.__loggers[logging.CRITICAL].critical(message)

    def createHandlers(self):
        process_name = self.process_name
        dir = LOGS_DIR
        log_path = "log" + "\\{}".format(self.process_name)
        if not os.path.exists(os.path.join(dir, log_path)):
            os.makedirs(os.path.join(dir, log_path))
        day = datetime.now().strftime("%Y-%m-%d")
        handlers = {
            logging.INFO: os.path.join(dir, f'{log_path}\\{process_name}-info.log.{day}.log'),
            logging.ERROR: os.path.join(dir, f'{log_path}\\{process_name}-error.log.{day}.log'),
        }
        logLevels = handlers.keys()

        for level in logLevels:
            if level == 20:
                path = os.path.abspath(handlers[level])
                handlers[level] = TimedRotatingFileHandler(path, when="D", backupCount=60, encoding='utf-8', interval=1)
                handlers[level].suffix = "%Y-%m-%d.log"
                handlers[level].extMatch = r"^\d{4}-\d{2}-\d{2}.log$"
                handlers[level].extMatch = re.compile(handlers[level].extMatch)
            else:
                path = os.path.abspath(handlers[level])
                handlers[level] = TimedRotatingFileHandler(path, when="D", backupCount=60, encoding='utf-8', interval=1)
                handlers[level].suffix = "%Y-%m-%d.log"
                handlers[level].extMatch = r"^\d{4}-\d{2}-\d{2}.log$"
                handlers[level].extMatch = re.compile(handlers[level].extMatch)
        return handlers


logger = RpaLog()
"""
RetryException :可重试异常
EndException: 不可重试异常
子类:
RPA 异常类  rpa抛出的异常
DATA 异常类  数据导致的异常
Project 异常类  未知的异常
"""


class RetryException(Exception):
    """可重试异常"""

    def __init__(self, msg):
        self.msg = msg

    def __str__(self):
        return self.msg


class EndException(Exception):
    """不可重试异常"""

    def __init__(self, msg):
        self.msg = msg

    def __str__(self):
        return self.msg


class DataException(EndException):
    """因操作的数据原因导致的异常"""

    def __init__(self, msg):
        self.msg = msg

    def __str__(self):
        return self.msg


class RpaException(RetryException):
    """因网页错误导致错误不能进行"""

    def __init__(self, msg=None):
        self.msg = msg

    def __str__(self):
        return self.msg


class ProjectException(EndException):
    """未知异常,但导致整个项目无法进行的异常,抛出整个程序停止"""

    def __init__(self, msg):
        self.msg = msg

    def __str__(self):
        return self.msg


def func_notes(func):
    doc = func.__doc__
    if not doc:
        return ""
    data = doc.split('\n')
    return data[1].strip()


def catch_rpa_except(func):
    """
    实现对RPA函数的异常捕获
    :param func:
    :return:
    """

    def inner(self, *args, **kwargs):
        start_time = time.time()
        try:
            result = func(self, *args, **kwargs)
            logger.info("{}:{} , 耗时:{:.2f} , 参数:{}{}".format(func.__name__, func_notes(func),
                                                                 time.time() - start_time, args, kwargs))
            return result
        except Exception as e:
            logger.error("key:{},{}:{} , 耗时:{:.2f} , 参数:{}{} , 错误信息{}".format(logger.key, func.__name__,
                                                                                      func_notes(func),
                                                                                      time.time() - start_time, args,
                                                                                      kwargs,
                                                                                      traceback.format_exc()))
            raise RpaException(str(e)) from e

    return inner


class RPAMetaclass(type):
    """基类,对RPA函数分别进行异常捕获"""

    def __new__(cls, name, bases, attrs):
        for k, v in attrs.items():
            if not hasattr(v, '__call__'):
                continue
            attrs[k] = catch_rpa_except(v)
        return type.__new__(cls, name, bases, attrs)


def logger_time(location):
    """
    打印方法执行时间
    :param location:定位
    :return:
    :note:
        - 打印方法执行时间
    """

    def inner(func):
        @wraps(func)
        def inner_echo(*args, **kwargs):
            start_time = time.time()
            res = func(*args, **kwargs)
            logger.info("{}, 耗时:{:.2f}S, {}, {}".format(location,
                                                          time.time() - start_time, func.__name__, func_notes(func)))
            return res

        return inner_echo

    return inner


def kill_wps(func):
    """
    杀死对应wps的进程
    :return:
    """

    def inner_func(*args, **kwargs):
        try:
            cmd = 'taskkill /F /IM wps.exe'
            os.system(cmd)
            res = func(*args, **kwargs)
            return res
        except Exception:
            pass

    return inner_func


def kill_web():
    """
    杀死对应的进程
    :return:
    """
    try:
        cmd = 'taskkill /F /IM chrome.exe'
        cmd1 = 'taskkill /F /IM iexplore.exe '
        os.system(cmd)
        os.system(cmd1)
    except Exception:
        pass


class Singleton:
    """
    单例装饰器。
    """
    __cls = dict()

    def __init__(self, cls):
        self.__key = cls

    def __call__(self, *args, **kwargs):
        if self.__key not in self.cls:
            self[self.__key] = self.__key(*args, **kwargs)
        return self[self.__key]

    def __setitem__(self, key, value):
        self.cls[key] = value

    def __getitem__(self, item):
        return self.cls[item]

    @property
    def cls(self):
        return self.__cls

    @cls.setter
    def cls(self, cls):
        self.__cls = cls


class RPAUtils(object, metaclass=RPAMetaclass):
    """
    包含方法:
        点击元素  click
        图片点击  image2roiClick
        填写input框   input
        处理select下拉框  options
        判断元素是否存在   visible
        跳转网页   skip_page
        获取元素文本  get_text
        获取元素html源码  get_html
        获取元素的value  get_value
        执行js  implement_js
        刷新页面 refresh
        上传文件  upload
        设置元素的属性 set_attribute
        跳转到新的页面  switch_page
        向网页发送按键  send_keys
        鼠标操作   mouse_action
        网页截屏  screen_picture
        捕获新打开的界面 catch_page
        根据图片定位元素点击 screenshots_click
        关闭网页alert弹出框  confirm
        获取当前页面的源码   page_source
        弹出windows提示框   windows_tips
        获取cookie     get_cookies
        调试模式的chrome   debug_browser
        弹出windows的对话框 windows_dialogue
        回退网页 back
        退出浏览器  out
    """

    def __init__(self, process, browser=None, overtime=3, **kwargs):
        self.process_name = process
        self.overtime = overtime
        self.browser_type = Chrome if not browser else browser
        self.url = None
        self.browser = None
        self.driver = None
        self.is_open = False
        self.process_key = uuid.uuid4()

    def _init_browser(self, url, driver_options=None, prefs=None, **kwargs):
        """
        初始化浏览器,设置浏览器参数,打开网页
        @param url: 打开网址
        @param browser_type: 启动浏览器类型
        @param driver_options: 浏览器选项设置,格式为{'prefs':{'safebrowsing.enabled': True}} <dict>
        @return:浏览器对象和操作页面对象
        @note:
        """
        self.url = url
        if self.browser_type == Chrome and driver_options:
            option = ChromeOptions()
            option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            if driver_options:
                prefs.update(driver_options)
            option.add_experimental_option('prefs', prefs)
            browser = self.browser_type(option)
        else:
            browser = self.browser_type()
        browser.max_window()
        driver = browser.create(self.url)
        self.browser = browser
        self.driver = driver
        self.is_open = True

    def init_browser(self, url, driver_options=None, headless=False, google_log=False, **kwargs):
        """
        初始化浏览器,设置浏览器参数,打开网页
        @param url: 打开网址
        @param browser_type: 启动浏览器类型
        @param driver_options: 浏览器选项设置,格式为{'prefs':{'safebrowsing.enabled': True}} <dict>
        @return:浏览器对象和操作页面对象
        @note:
        """
        self.url = url
        if self.browser_type == Chrome:
            option = ChromeOptions()
            if headless == True:
                option.headless = True
            prefs = {'safebrowsing.enabled': True}
            if driver_options:
                prefs.update(driver_options)
            if google_log:
                option.add_experimental_option('w3c', False)
                option.set_capability("loggingPrefs", {'performance': 'ALL'})
            option.add_experimental_option('prefs', prefs)
            browser = self.browser_type(option)
        else:
            browser = self.browser_type()
        if headless == False:
            browser.max_window()
        driver = browser.create(self.url)
        self.browser = browser
        self.driver = driver
        self.is_open = True

    def catch_ie(self, url, ie_path=None, **kwargs):
        """
        捕获IE
        :return:
        """
        # 启动IE
        # import rpa.win32._tools as system
        # system.start_program(ie_path)
        # sleep(3)
        from rpa import IE, IEOptions
        option = IEOptions()
        browser = IE(option)
        _rpa_bdd705_webpage = browser.create("about:blank")
        driver = browser.web_driver
        driver.command_executor._commands['attachToBrowser'] = ('POST', '/session/$sessionId/attach')
        driver.execute('attachToBrowser', {'url': url})
        new_page = browser.catch(url, regex=False)
        _rpa_bdd705_webpage.close()
        self.driver = new_page
        self.browser = browser

    def catch_chrome(self, url=None, title=None, regex=False, **kwargs):
        """
        捕获已经打开的chrome浏览器
        :param url:
        :param kwargs:
        :return:
        """
        option = ChromeOptions()
        option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        rpa_a66883_browser = Chrome(option)
        if url:
            rpa_942b27_webpage = rpa_a66883_browser.catch(url, regex=regex)
        elif title:
            rpa_942b27_webpage = rpa_a66883_browser.catch_by_title(url, regex=regex)
        self.driver = rpa_942b27_webpage

    def debug_browser(self, url, chrome_path, **kwargs):
        """
        打开调试模式的浏览器(谷歌浏览器)
        :param url:需要捕获的url
        :param chrome_path:谷歌安装路径
        :return:
        :note:
            -.进入谷歌的安装目录(含有chrome.exe的目录),进入cmd,执行命令:chrome.exe --remote-debugging-port=9222 --disable-popup-blocking
            -.使用弹出的浏览器去执行操作,操作到需要捕获的页面为止
        """
        t1 = threading.Thread(target=self.open_browser, args=(chrome_path,))
        t1.start()
        self.url = url
        option = ChromeOptions()
        option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        chrome = Chrome(option)
        driver = chrome.catch(self.url)
        self.browser = chrome
        self.driver = driver

    def get_elements(self, element="all", id="", css="", xpath=None, frame=None, wait_time=None, **kwargs):
        """
        获取所有元素
        :param id:
        :param css:
        :param xpath:
        :param frame:
        :param wait_time:
        :param kwargs:
        :return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        if isinstance(wait_time, str):
            wait_time = int(wait_time)
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        return self.driver.get_elements(element_name=element, timeout=wait_time)

    def click(self,
              element,
              method=ElementClickMethod.MOUSE_CLICK,
              id="",
              css="",
              xpath=None,
              frame=None,
              wait_time=None,
              double=False,
              **kwargs):
        """
        点击元素
        @param element: 元素名 <str>
        @param method: 点击方法,默认为ElementClickMethod.MOUSE_CLICK <function>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        if isinstance(wait_time, str):
            wait_time = int(wait_time)
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        if double:
            self.driver.get_element(element, timeout=wait_time).double_click()
        else:
            self.driver.get_element(element, timeout=wait_time).click(method=method)
        # 防止息屏
        # sleep(1.5)
        # import rpa.win32._keyboard_action as keyboard
        # keyboard.key_send("down")

    def image2roiClick(self,
                       localImagePath,
                       captureRange=None,
                       accuracy=0.8,
                       buttonType="left",
                       clickType="single",
                       clickPosition="center",
                       horizontalOffset=0,
                       verticalOffset=0,
                       timeout=30):
        """
        图片点击，OCR定位传入图片位置进行点击
        @param localImagePath: 图片的绝对路径
        @return: None
        """
        OCR().image2roiClick(localImagePath=localImagePath,
                             captureRange=captureRange,
                             accuracy=accuracy,
                             buttonType=buttonType,
                             clickType=clickType,
                             clickPosition=clickPosition,
                             horizontalOffset=horizontalOffset,
                             verticalOffset=verticalOffset,
                             timeout=timeout)

    def input(self, element, content, xpath=None, frame=None, id="", css="", wait_time=None, **kwargs):
        """
        填写输入框
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        self.driver.get_element(element, timeout=wait_time).input(content, simulate=True)

    def dormancy(self, num, **kwargs):
        """
        睡几秒
        :param num:
        :param kwargs:
        :return:
        """
        if isinstance(num, str):
            num = int(num)
        time.sleep(num)

    def hover_element(self, xpath_or_id):
        """
        鼠标移动到某一个元素上边
        """
        element = self.get_elements(xpath_or_id)
        ActionChains(self.driver).move_to_element(element).perform()

    def options(self, element, content, xpath=None, frame=None, id="", css="", wait_time=None, **kwargs):
        """
        下拉列表
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        :return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        if isinstance(content, int):
            self.driver.get_element(element, timeout=wait_time).option_by_index(content)
        else:
            self.driver.get_element(element, timeout=wait_time).option(content)

    def visible(self, element, xpath=None, frame=None, id="", css="", wait_time=None, is_raies=None, **kwargs):
        """
        元素是否可见
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @param is_raies:  是否抛出异常 <bool>
        :return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        try:
            self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
            self.driver.switch_to_frame_by_path(frame)
            is_boolean = self.driver.get_element(element, timeout=wait_time).is_visible()
        except Exception:
            if is_raies:
                raise Exception("element not found")
            is_boolean = False
        return is_boolean

    def wait_element(self, element, xpath=None, frame=None, id="", css="", wait_time=None, **kwargs):
        """
        等待元素
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @param wait_time:等待时间
        @return: 返回True就是存在
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        res = self.driver.wait_element_loaded(element, timeout=wait_time)
        return res

    def skip_page(self, url, **kwargs):
        """
        跳转网页
        :param url:
        :return:
        """
        self.driver.navigate(url)

    def open_page(self, url, **kwargs):
        """
        跳转网页
        :param url:
        :return:
        """
        self.driver = self.browser.create(url)

    def get_text(self, element, xpath=None, frame=None, id="", css="", wait_time=None, to_error=False, **kwargs):
        """
        获取元素文本内容
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param to_error:  是否抛出异常
        @param frame:  元素所在frame  <list>
        :return: 文本内容 <str>
        """
        try:
            xpath = xpath if xpath else []
            frame = frame if frame else []
            wait_time = wait_time if wait_time else self.overtime
            self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
            self.driver.switch_to_frame_by_path(frame)
            elem = self.driver.get_element(element, timeout=wait_time)
            result = elem.text()
        except Exception:
            if to_error:
                import traceback
                raise Exception(traceback.format_exc())
            else:
                result = ""
        return result

    def get_html(self, element, xpath=None, frame=None, id="", css="", wait_time=None, to_error=False, **kwargs):
        """
        获取元素html代码
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param to_error:  是否抛出异常
        @param frame:  元素所在frame  <list>
        :return: 文本内容 <str>    <li>123123<li/>
        """
        try:
            xpath = xpath if xpath else []
            frame = frame if frame else []
            wait_time = wait_time if wait_time else self.overtime
            self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
            self.driver.switch_to_frame_by_path(frame)
            result = self.driver.get_element(element, timeout=wait_time).html()
        except Exception:
            if to_error:
                import traceback
                raise Exception(traceback.format_exc())
            else:
                result = ""
        return result

    def get_value(self, element, xpath=None, frame=None, id="", css="", attr=None, wait_time=None, **kwargs):
        """
        获取元素的值
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        :return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        if not attr:
            result = self.driver.get_element(element, timeout=wait_time).value()
            if not result:
                result = self.driver.get_element(element, timeout=wait_time).text()
        else:
            result = self.driver.get_element(element, timeout=wait_time).get_attr(attr)
        return result

    def implement_js(self, js, **kwargs):
        """
        执行js
        @param js:执行的js语句
        """
        self.driver.execute_js(js)

    def refresh(self, **kwargs):
        """
        刷新网页
        @return:
        @note:
        """
        self.driver.reload()

    def upload(self, element, file_path, xpath=None, frame=None, id="", css="", wait_time=None, **kwargs):
        """
        上传文件
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        self.driver.get_element(element, timeout=wait_time).upload(file_path)

    def set_attribute(self, element, attribute, value, xpath=None, frame=None, id="", css="", wait_time=None, **kwargs):
        """
        设置元素属性
        @param element: 元素名 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @param attribute: 元素的属性  <str>
        @param value: 设置的属性值  <str>
        @return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        wait_time = wait_time if wait_time else self.overtime
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        self.driver.get_element(element, timeout=wait_time).set_attr(attribute, value)

    def switch_page(self, parameter, is_title=True, regex=False, **kwargs):
        """
        切换页面
        @param is_title: 是否根据标题模糊匹配,默认为True
        @param parameter: title or url
        @param regex: 是否正则匹配,默认为False
        @return: 捕获的page
        @note:
        """
        if is_title:
            new_driver = self.browser.catch_by_title(parameter, regex=regex)
        else:
            new_driver = self.browser.catch(parameter, regex=regex)
        self.driver = new_driver
        # return new_driver

    def catch_page(self, url, regex=False, **kwargs):
        """
        捕获打开的新页面
        @param url:新页面的地址
        """
        self.driver = self.browser.catch(url, regex=regex)

    def page_source(self, **kwargs):
        """
        获取当前的网页源码
        @param kwargs:
        @return:
        """
        data = self.driver.web_driver.page_source
        return data

    def send_keys(self, content=None, key=None, **kwargs):
        """
        发送按键
        @param content: 发送文本
        @param key: 发送按键,使用rpa包中的Keys类中包含的属性
        @return:
        """
        if content:
            self.driver.send_keys(content)
        self.driver.send_keys(getattr(Keys, key))

    def mouse_action(self,
                     element=None,
                     action_type=0,
                     times=0,
                     xpath=None,
                     frame=None,
                     id="",
                     css="",
                     wait_time=None,
                     **kwargs):
        """
        鼠标操作:默认鼠标单击
        @param element: 移入或移出元素 <str>
        @param element: 元素名 <str>

        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @param action_type: 操作类型  0:鼠标点击  1:鼠标移入  2:鼠标移出 <int>
        @param times: 点击类型  0:单击   1:双击   2:右击 <int>
        @return:
        """
        if element:
            xpath = xpath if xpath else []
            frame = frame if frame else []
            wait_time = wait_time if wait_time else self.overtime
            self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
            self.driver.switch_to_frame_by_path(frame)
            if action_type == 1:
                self.driver.get_element(element, timeout=wait_time).mouse_move_in()
            elif action_type == 2:
                self.driver.get_element(element, timeout=wait_time).mouse_move_out()
        if not element and action_type == 0:
            if times == 0:
                self.driver.mouse_click(is_double=False)
            elif times == 1:
                self.driver.mouse_click(is_double=True)
            else:
                self.driver.mouse_context_click()

    def screen_picture(self, save_path, element=None, xpath=None, frame=None, id="", css="", wait_time=None, **kwargs):
        """
        网页截图
        @param save_path: 保存路径
        @param element: 元素名 是否对网页元素截图,默认不传,对整个浏览器截屏 <str>
        @param id:  元素id  <str>
        @param css:  元素css路径 <str>
        @param xpath:  元素xpath定位 <list>
        @param frame:  元素所在frame  <list>
        @return:
        """
        xpath = xpath if xpath else []
        frame = frame if frame else []
        if not element:
            self.driver.screen_shot(save_path)
        self.driver.add_element(element, id=id, css=css, xpath=xpath, frame=frame)
        self.driver.switch_to_frame_by_path(frame)
        self.driver.get_element(element).screen_shot(save_path)

    def screenshots_click(self, save_path, target_path, position="center", button="left", **kwargs):
        """
        根据图片定位元素点击
        @param save_path: 全屏截图保存路径 <str>
        @param target_path: 目标元素截图路径  <str>
        @param position: 图片位置类型  center/left/right/left_right  <str>
        @param button: win32点击方式(鼠标左击/右击)  <str>
        @return:
        """
        self.driver.screen_shot(save_path)
        from rpa.ocr.ocr import OCR
        coordinate = OCR().image2roi(save_path, target_path, position)
        import rpa.win32
        rpa.win32.mouse_click(x=coordinate[0], y=coordinate[1], button=button, count=1, duration=0.1)

    def confirm(self, is_sure=True, **kwargs):
        """
        关闭浏览器alert弹出框
        :return:
        """
        try:
            confirm = self.driver.web_driver.switch_to.alert
            if is_sure:
                confirm.accept()
            else:
                confirm.dismiss()
        except Exception:
            print("未出现弹出框！")

    def get_cookies(self, **kwargs):
        """
        获取cookies
        :return: <dict>
        """
        cookie = self.driver.web_driver.get_cookies()
        return cookie

    def close(self, **kwargs):
        """
        关闭页面
        :return:
        """
        self.driver.close()

    def out(self, **kwargs):
        """
        关闭浏览器
        :return:
        """
        sleep(2)
        self.browser.quit()
        self.is_open = False

    def back(self, **kwargs):
        """
        回退网页
        :return:
        """
        sleep(2)
        self.browser.web_driver.back()

    def get_google_log(self):
        """
        获取谷歌日志
        :return:
        """
        clog = self.driver.web_driver.get_log('performance')
        logs = [json.loads(log['message'])['message'] for log in clog]
        return logs

    def execute_cdp_cmd(self, cmd, cmd_args):
        """
        使用web_driver执行cmd命令
            execute_cdp_cmd("Network.getAllCookies", {'requestId': "223rr4r4r4"}) 获取所有cookie
            execute_cdp_cmd("Network.getRequestPostData ", {'requestId': "223rr4r4r4"}) 获取请求的postdata
            execute_cdp_cmd("Network.getResponseBody ", {'requestId': "223rr4r4r4"}) 获取ResponseBody
        :param cmd:
        :param cmd_args:
        :return:
        """
        return self.driver.web_driver.execute_cdp_cmd(cmd, cmd_args)

    @staticmethod
    def windows_tips(message, tips, **kwargs):
        """
        弹出windows提示框
        :param message:提示信息
        :param tips:提示类型
        :return:
        """
        import win32api
        import win32con
        win32api.MessageBox(0, message, tips, win32con.MB_OK)

    @staticmethod
    def windows_dialogue(message, tips, **kwargs):
        """
        弹出windows的对话框
        :param message: 提示信息
        :param tips:提示类型
        :return:0/1
        """
        import win32api
        import win32con
        result = win32api.MessageBox(0, message, tips, win32con.MB_YESNO)
        if result == 6:
            return 1
        else:
            return 0

    @staticmethod
    def open_browser(chrome_path: str):
        """
        调试模式下打开浏览器,用于捕获已经打开的浏览器
        :param chrome_path: 谷歌浏览器安装目录(包含chrome.exe的目录)
        :return:
        """
        import os
        cmd = 'taskkill /F /IM chrome.exe'
        try:
            os.system(cmd)
        except Exception:
            print("no process kill")
        os.chdir(chrome_path)
        os.system("chrome.exe --remote-debugging-port=9222 --disable-popup-blocking")

    @staticmethod
    def send_email(title, user, password, receive, server, port, content, file_path, file_name, **kwargs):
        """
        发送带附件的邮件
        :param title: 标题
        :param user: 发送者
        :param password: SMTP服务密码
        :param receive: 收件人
        :param server: SMTP服务地址
        :param port: SMTP端口
        :param content: 邮件内容
        :param file_path: 文件路径
        :param file_name: 文件名
        :return:
        :note:163邮箱发送频率高会返回554错误
        """
        msg = MIMEMultipart()
        msg["Subject"] = title
        msg["From"] = user
        msg["To"] = receive
        part = MIMEText(content)
        msg.attach(part)
        part = MIMEApplication(open(file_path, 'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename=file_name)
        msg.attach(part)
        s = smtplib.SMTP(server, port, timeout=60)
        s.login(user, password)
        s.sendmail(user, receive, msg.as_string())
        s.close()

    @staticmethod
    def ui_click(selector=None, element=None, method="simulation", timeout=20, delay=0, to_error=True, count=1):
        try:
            if element:
                element_action.click_element(click_method=method,
                                             ui_driver=ui_driver,
                                             element=element,
                                             count=count,
                                             delay=delay)
            else:
                element_action.click_element(click_method=method,
                                             ui_driver=ui_driver,
                                             selector=selector,
                                             timeout=timeout,
                                             focus=False,
                                             count=count,
                                             delay=delay)
        except Exception:
            if to_error:
                raise Exception(traceback.format_exc())
            else:
                logger.info("桌面元素点击异常！")

    @staticmethod
    def ui_input(selector=None, element=None, content="", timeout=60):
        if selector:
            element_action.set_edit_text(text=content,
                                         ui_driver=ui_driver,
                                         selector=selector,
                                         timeout=timeout,
                                         focus=False)
        else:
            element_action.set_edit_text(text=content, ui_driver=ui_driver, element=element)

    @staticmethod
    def ui_get_element_attributes(attributes=("Name",), element=None, selector=None, timeout=60):
        """

        :param attributes: ["Name"]
        :param element:
        :param selector:
        :param timeout:
        :return:
        list or text
        """
        if element:
            res = element_action.get_element_attributes(attributes=attributes, ui_driver=ui_driver, element=element)
        else:
            res = element_action.get_element_attributes(attributes=attributes,
                                                        ui_driver=ui_driver,
                                                        selector=selector,
                                                        timeout=timeout,
                                                        focus=False)
        return res

    @staticmethod
    def ui_get_element(selector=None, index=None):
        """
        获取元素
        :param selector:
        :param index: 获取的元素下标，默认获取所有元素
        :return:
        [] or element
        """
        res = None
        if index == 0:
            res = element_action.query_element(ui_driver=ui_driver, selector=selector)
        elif index:
            elements = element_action.query_element_all(ui_driver=ui_driver, selector=selector)
            try:
                res = elements[index]
            except:
                logger.info("未找到指定下标元素！")
        else:
            res = element_action.query_element_all(ui_driver=ui_driver, selector=selector)
        return res

    @staticmethod
    def ui_wait_element(selector="", timeout=60, to_error=False):
        try:
            element_action.wait_element(ui_driver=ui_driver, selector=selector, timeout=timeout, focus=False)
        except Exception:
            if to_error:
                raise Exception(traceback.format_exc())
            else:
                pass

    @staticmethod
    def keyboard_input_text(content, is_text=False, delay=0):
        """键盘输入快捷键或者文本"""
        # from rpa.ui import keyboard_action
        if is_text:
            keyboard_action.keyboard_write(text=content, delay=delay, auxiliary_key="")
        else:
            keyboard_action.keyboard_send(hotkey=content, auxiliary_key="")

    @staticmethod
    def pause_show_tip(title="暂停流程", info="程序运行中，点击按钮继续", button_text="点击继续运行", timeout=0):
        """
        流程暂停并进行提示
        :param title:
        :param info:
        :param button_text:
        :param timeout: 超时后自动关闭的时间，0代表永久不自动关闭
        :return:
        """
        rpa.tools.Dialog(title=title, info=info, button_text=button_text, timeout=timeout)


MAX_WAIT = 1073741823


def retry(*dargs, **dkw):
    """
    Decorator function that instantiates the Retrying object
    @param *dargs: positional arguments passed to Retrying object
    @param **dkw: keyword arguments passed to the Retrying object
    """
    # support both @retry and @retry() as valid syntax
    if len(dargs) == 1 and callable(dargs[0]):  # 当用法为@retry不带括号时走这条路径,dargs[0]为retry注解的函数,返回函数对象wrapped_f

        def wrap_simple(f):

            @six.wraps(f)
            def wrapped_f(*args, **kw):
                return Retrying().call(f, *args, **kw)

            return wrapped_f

        return wrap_simple(dargs[0])

    else:  # 当用法为@retry()带括号时走这条路径，返回函数对象wrapped_f

        def wrap(f):

            @six.wraps(f)
            def wrapped_f(*args, **kw):
                return Retrying(*dargs, **dkw).call(f, *args, **kw)

            return wrapped_f

        return wrap


class Retrying(object):

    def __init__(
            self,
            stop=None,
            wait=None,
            stop_max_attempt_number=RETRY_TIME,  # 最大重试次数
            stop_max_delay=None,  # 最大重试时间
            wait_fixed=None,  # 等待时间
            wait_random_min=None,
            wait_random_max=None,  # 随机等待时间区间
            wait_incrementing_start=None,
            wait_incrementing_increment=None,  # 递增等待
            wait_exponential_multiplier=None,
            wait_exponential_max=None,
            retry_on_exception=None,  # 自定义校验捕获到的异常是否需要重试
            retry_on_result=None,  # 自定义函数校验执行到的结果是否需要重试
            wrap_exception=False,  # RetryError
            stop_func=None,  # 自定义停止函数
            wait_func=None,  # 自定义等待时间函数
            wait_jitter_max=None,  # 等待抖动值
            before_retry=None,  # 充实之前执行函数
            before_retry_kwargs=None,
    ):

        self._stop_max_attempt_number = 3 if stop_max_attempt_number is None else stop_max_attempt_number  # 在停止之前尝试的最大次数
        self._stop_max_delay = 100 if stop_max_delay is None else stop_max_delay  # 重试两个函数之前的最大延迟时间
        self._wait_fixed = 1000 if wait_fixed is None else wait_fixed  # 两次调用方法期间固定停留时长
        self._wait_random_min = 0 if wait_random_min is None else wait_random_min  # 在两次调用方法停留时长，停留最短时间
        self._wait_random_max = 1000 if wait_random_max is None else wait_random_max  # 在两次调用方法停留时长，停留最长时间
        self._wait_incrementing_start = 0 if wait_incrementing_start is None else wait_incrementing_start  # 每调用一次则会增加的时长
        self._wait_incrementing_increment = 100 if wait_incrementing_increment is None else wait_incrementing_increment  # 每次调用会增长的时间
        self._wait_exponential_multiplier = 1 if wait_exponential_multiplier is None else wait_exponential_multiplier  # 以指数的形式产生两次retrying之间的停留时间，产生的值为2^previous_attempt_number * wait_exponential_multiplier
        self._wait_exponential_max = MAX_WAIT if wait_exponential_max is None else wait_exponential_max
        self._wait_jitter_max = 0 if wait_jitter_max is None else wait_jitter_max
        self._before_retry = 0 if before_retry is None else before_retry
        self._before_retry_kwargs = 0 if before_retry_kwargs is None else before_retry_kwargs

        # TODO add chaining of stop behaviors
        # stop behavior
        stop_funcs = []  # 根据重试次数和延迟判断是否应该停止重试
        if stop_max_attempt_number is not None:  # 重试次数函数
            stop_funcs.append(self.stop_after_attempt)

        if stop_max_delay is not None:  # 最大重试时间函数
            stop_funcs.append(self.stop_after_delay)

        if stop_func is not None:
            self.stop = stop_func

        elif stop is None:  # 执行次数和延迟任何一个达到限制则停止
            self.stop = lambda attempts, delay: any(f(attempts, delay) for f in stop_funcs)

        else:
            self.stop = getattr(self, stop)

        # TODO add chaining of wait behaviors
        # wait behavior 延迟时间计算
        wait_funcs = [lambda *args, **kwargs: 0]
        if wait_fixed is not None:
            wait_funcs.append(self.fixed_sleep)

        if wait_random_min is not None or wait_random_max is not None:
            wait_funcs.append(self.random_sleep)  # 从最大等待时间和最小等待时间种随机

        if wait_incrementing_start is not None or wait_incrementing_increment is not None:
            wait_funcs.append(self.incrementing_sleep)  # 递增增加时间

        if wait_exponential_multiplier is not None or wait_exponential_max is not None:
            wait_funcs.append(self.exponential_sleep)  # 指数增加

        if wait_func is not None:
            self.wait = wait_func

        elif wait is None:  # 返回几个函数的最大值,作为等待时间
            self.wait = lambda attempts, delay: max(f(attempts, delay) for f in wait_funcs)

        else:
            self.wait = getattr(self, wait)

        # retry on exception filter
        if retry_on_exception is None:
            self._retry_on_exception = self.always_reject
        else:
            self._retry_on_exception = retry_on_exception

        # TODO simplify retrying by Exception types
        # retry on result filter
        if retry_on_result is None:
            self._retry_on_result = self.never_reject
        else:
            self._retry_on_result = retry_on_result

        self._wrap_exception = wrap_exception

        if before_retry is None:
            self._before_retry = self.before_on_retry
        else:
            self._before_retry = before_retry

    def before_on_retry(self, *args, **kwargs):
        return 1

    def stop_after_attempt(self, previous_attempt_number, delay_since_first_attempt_ms):
        """Stop after the previous attempt >= stop_max_attempt_number."""
        return previous_attempt_number >= self._stop_max_attempt_number

    def stop_after_delay(self, previous_attempt_number, delay_since_first_attempt_ms):
        """Stop after the time from the first attempt >= stop_max_delay."""
        return delay_since_first_attempt_ms >= self._stop_max_delay

    def no_sleep(self, previous_attempt_number, delay_since_first_attempt_ms):
        """Don't sleep at all before retrying."""
        return 0

    def fixed_sleep(self, previous_attempt_number, delay_since_first_attempt_ms):
        """Sleep a fixed amount of time between each retry."""
        return self._wait_fixed

    def random_sleep(self, previous_attempt_number, delay_since_first_attempt_ms):
        """Sleep a random amount of time between wait_random_min and wait_random_max"""
        return random.randint(self._wait_random_min, self._wait_random_max)

    def incrementing_sleep(self, previous_attempt_number, delay_since_first_attempt_ms):
        """
        递增等待
        Sleep an incremental amount of time after each attempt, starting at
        wait_incrementing_start and incrementing by wait_incrementing_increment
        """
        result = self._wait_incrementing_start + (self._wait_incrementing_increment * (previous_attempt_number - 1))
        if result < 0:
            result = 0
        return result

    def exponential_sleep(self, previous_attempt_number, delay_since_first_attempt_ms):
        # 指定指数增加等待
        exp = 2 ** previous_attempt_number
        result = self._wait_exponential_multiplier * exp
        if result > self._wait_exponential_max:
            result = self._wait_exponential_max
        if result < 0:
            result = 0
        return result

    def never_reject(self, result):
        """默认返回0则重试"""
        return False

    def always_reject(self, result):
        """默认判断是否是EndException"""
        return not isinstance(result, EndException)

    def should_reject(self, attempt):
        """
        根据每次执行结果或异常类型判断是否应该停止
        :param attempt:
        :return:
        """
        reject = False
        if attempt.has_exception:  # 假如异常在retry_on_exception参数中返回True，则重试,默认不传异常参数时，发生异常一直重试
            reject |= self._retry_on_exception(attempt.value[1])
        else:  # 假如函数返回结果在retry_on_result参数函数中为True，则重试
            reject |= self._retry_on_result(attempt.value)

        return reject

    def call(self, fn, *args, **kwargs):
        """
        实现重试
        :param fn: 执行函数
        :param args:
        :param kwargs:
        :return:
        """
        start_time = int(round(time.time() * 1000))
        attempt_number = 1  # 重试次数
        while True:
            try:
                attempt = Attempt(fn(*args, **kwargs), attempt_number, False)
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                error_msg = repr(e)
                logger.error("文件:{},{}行,出现异常:{}".format(e.__traceback__.tb_frame.f_globals["__file__"],
                                                               exc_tb.tb_next.tb_lineno, error_msg))
                tb = sys.exc_info()  # 获取异常信息
                attempt = Attempt(tb, attempt_number, True)  # 抛出异常

            if not self.should_reject(attempt):
                return attempt.get(self._wrap_exception)

            delay_since_first_attempt_ms = int(round(time.time() * 1000)) - start_time
            if self.stop(attempt_number, delay_since_first_attempt_ms):  # 根据重试次数和函数已经执行时间判断是否应该停止
                if not self._wrap_exception and attempt.has_exception:
                    # get() on an attempt with an exception should cause it to be raised, but raise just in case
                    raise attempt.get()
                else:
                    raise RetryError(attempt)
            else:  # 不停止则等待一定时间，延迟时间根据wait函数返回值和_wait_jitter_max计算
                logger.error("异常第{}次重试".format(attempt_number))
                sleep = self.wait(attempt_number, delay_since_first_attempt_ms)
                if self._wait_jitter_max:
                    jitter = random.random() * self._wait_jitter_max
                    sleep = sleep + max(0, jitter)
                time.sleep(sleep / 1000.0)
                if self._before_retry_kwargs:
                    self._before_retry(self._before_retry_kwargs)
                else:
                    self._before_retry()
            attempt_number += 1


class Attempt(object):
    """
    An Attempt encapsulates a call to a target function that may end as a
    normal return value from the function or an Exception depending on what
    occurred during the execution.
    #Attempt将函数执行结果或者异常信息以及执行次数作为内部状态,用True或False标记是内部存的值正常执行结果还是异常
    """

    def __init__(self, value, attempt_number, has_exception):
        self.value = value  # 函数的执行结果或是执行异常信息
        self.attempt_number = attempt_number  # 尝试次数
        self.has_exception = has_exception  # 是否出现异常

    def get(self, wrap_exception=False):
        """
        Return the return value of this Attempt instance or raise an Exception.
        If wrap_exception is true, this Attempt is wrapped inside of a
        RetryError before being raised.

        """
        if self.has_exception:
            if wrap_exception:
                raise RetryError(self)  # 对异常用RetryError包裹
            else:
                six.reraise(self.value[0], self.value[1], self.value[2])  # 重新抛出异常
        else:
            return self.value

    def __repr__(self):
        if self.has_exception:
            return "Attempts: {0}, Error:\n{1}".format(self.attempt_number, "".join(traceback.format_tb(self.value[2])))
        else:
            return "Attempts: {0}, Value: {1}".format(self.attempt_number, self.value)


class RetryError(Exception):
    """
    A RetryError encapsulates the last Attempt instance right before giving up.
    """

    def __init__(self, last_attempt):
        self.last_attempt = last_attempt

    def __str__(self):
        return "RetryError[{0}]".format(self.last_attempt)


class RpaArray(list):
    """定义一个数组,用来装流程"""

    def __init__(self):
        super(RpaArray, self).__init__()
        self.sequence = []

    def add_element(self, **kwargs):
        """添加element对象"""
        element = Element(**kwargs)
        self.sequence.append(element.element)
        self.append({element.element: element})

    def add_func(self, function, **kwargs):
        """添加函数"""
        if not callable(function):
            raise DataException("object is not a function")
        self.sequence.append(function.__name__)
        self.append({function.__name__: (function, kwargs)})

    def change_index(self, object_name, target_name):
        """调换索引"""
        if object_name not in self.sequence:
            raise DataException("object_name is not in array")
        index = self.sequence.index(object_name)
        data = self.pop(index)
        name = self.sequence.pop(index)
        target_index = self.sequence.index(target_name)
        self.sequence.insert(target_index, name)
        self.insert(target_index, data)

    def change_parameter(self, object_name, **kwargs):
        """切换参数"""
        if object_name not in self.sequence:
            raise DataException("object_name is not in array")
        index = self.sequence.index(object_name)
        self[index] = Element(**kwargs)

    def del_element(self, object_name):
        """删除步骤"""
        if object_name not in self.sequence:
            raise DataException("object_name is not in array")
        index = self.sequence.index(object_name)
        self.sequence.pop(index)
        print("弹出的index:{}".format(index))
        self.pop(index)


class Element(object):
    """实例化流程步骤"""

    def __init__(self, **kwargs):
        """
        初始化函数,将所有的参数初始化为属性
        :param kwargs:
        """
        for k, v in kwargs.items():
            self.__setattr__(k, v)
        self.custom = False  # 是否为自定义方法
        self.verification_data()

    def verification_data(self):
        """
        验证填写参数类型和属性是否正确
        :return:
        """
        for name, value in vars(self).items():
            if name == "custom":
                continue
            if name == "content" or name == "content_type":
                if vars(self).get(name):
                    if not isinstance(vars(self).get("content"), eval(vars(self).get("content_type"))):
                        raise DataException("填写参数类型错误!")
            elif name == "func":
                if callable(vars(self).get("func")):
                    self.__setattr__("custom", True)  # 自定义方法添加
                elif not Element.verification_attr(RPAUtils, vars(self).get("func")):
                    raise DataException("RPA方法错误")
            else:
                if not value:
                    raise DataException("传参为空!")

    @staticmethod
    def verification_attr(obj, name):
        """验证函数参数是否在对象上"""
        try:
            getattr(obj, name)
            return 1
        except AttributeError:
            return 0


class Trans(object):

    def __init__(self,
                 process,
                 func_list,
                 pause=0.5,
                 browser=None,
                 need=None,
                 unwanted=None,
                 call_back=None,
                 verification=None,
                 args_dict=None,
                 otherwise=None):
        """
        初始化
        :param process: 流程名
        :param pause: 流程执行间隔
        :param need: 重试时需要保留的步骤
        :param unwanted: 重试时不需要保留的步骤
        :param call_back: 回调函数
        :param verification: 验证函数,当前步骤执行完成后验证是否成功的函数
        :param args_dict: 传参
        :param debug: 不可重试异常是否关闭相关应用
        """
        self.ranks = RpaArray()  # [{"element_name",obj}]
        self.pause = pause
        self.func_list = func_list
        self.need = need
        self.unwanted = unwanted
        self.call_back = call_back
        self.verification = verification
        self.driver = RPAUtils(process=process, browser=browser)
        self.browser = self.driver.browser
        self.process = process
        self.result = {"driver": self.driver}  # 接收函数的返回值,传递必要参数{web_driver}
        self.args_dict = args_dict
        self.debug = True
        if otherwise:
            self.result.update(otherwise)
        self.register()

    def run(self):
        """
        运行整个流程
        :return:
        """
        for action in self.ranks:
            element_function = list(action.values())[0]  # 获取Element对象
            if list(action.values())[0].custom:
                function = getattr(element_function, "func")
            else:
                function = getattr(self.driver, element_function.func)
            kwargs = vars(element_function)  # 获取所有的参数
            kwargs.update({"driver": self.driver, "result": self.result})
            try:
                value = function(**kwargs)
            except Exception as e:
                # 关闭浏览器
                if self.driver.is_open:
                    self.driver.out()
                # 记录错误步骤
                import sys
                info = sys.exc_info()
                self.result[element_function.element] = "error"
                logger.info(info)
                raise info[0](e)
            setattr(element_function, "result", value)  # 添加结果集属性
            self.result[element_function.element] = element_function
            time.sleep(self.pause)
        # 验证流程是否成功的函数
        if self.verification:
            self.verification()
        return self.result

    def register(self):
        """
        将流程注册到动作集中去
        :return:
        """
        for action in self.func_list:
            element = action.get("element", None)
            if element and self.args_dict.get(element, None):
                action.update(self.args_dict.get(element))
            try:
                self.ranks.add_element(**action)
            except Exception as e:
                print(element)
                raise EndException(f"action collections register error:{repr(e)}")

    def change_process(self, number, result):
        """删除或保留一些流程"""
        if self.need:
            this_need = self.need[number - 1]
            if not set(this_need) <= set(self.ranks.sequence):
                raise EndException("process need error!")
            for process in self.ranks.sequence:
                if process not in this_need:
                    self.ranks.del_element(process)
        if self.unwanted:
            this_unwanted = self.unwanted[number - 1]
            if not set(this_unwanted) <= set(self.ranks.sequence):
                raise EndException("process unwanted error!")
            for process in this_unwanted:
                self.ranks.del_element(process)

    def clear(self):
        """
        清空
        :return:
        """
        self.ranks.clear()


def click_element(element_name, selector, is_msg=False, timeout=60, is_raise=True, times=1):
    """
    点击元素
    :param element_name:
    :param selector:
    :param is_msg:
    :param timeout:
    :param is_raise:
    :return:
    """
    try:
        element = wait_element(selector=selector, timeout=timeout)
        if is_msg:
            element.invoke()
        else:
            for _ in range(times):
                element.click()
        logger.info(f"点击元素:{element_name}成功！")
    except Exception as e:
        logger.error(f"点击元素:{element_name}失败,error_mgs:{traceback.format_exc()}")
        if is_raise: raise e


def click_next_element(element_name, selector, is_msg=False, timeout=360, is_raise=True, times=1):
    """
    点击元素
    :param element_name:
    :param selector:
    :param is_msg:
    :param timeout:
    :param is_raise:
    :return:
    """
    try:
        element = wait_element(selector=selector, timeout=timeout).next_sibling
        if is_msg:
            element.invoke()
        else:
            for _ in range(times):
                element.click()
        logger.info(f"点击元素:{element_name}成功！")
    except Exception as e:
        logger.error(f"点击元素:{element_name}失败,error_mgs:{traceback.format_exc()}")
        if is_raise: raise e


def input_element(element_name, selector, value, is_msg=False, timeout=60, delay_time=0):
    """
    填写填入框内容
    :param element_name:
    :param selector:
    :param value:
    :param is_msg:系统信息
    :return:
    """
    try:
        element = wait_element(selector=selector, timeout=timeout)
        element.click()
        if is_msg:
            element.set_text(text=value)
        else:
            time.sleep(0.5)
            keyboard.write(value, delay=delay_time)
        logger.info(f'元素:{element_name}填写成功,填写值:{value}')
    except Exception as e:
        logger.error(f'元素:{element_name}填写失败,填写值:{value}.error_msg:{traceback.format_exc()}')
        raise e


def input_next_element(element_name, selector, value, is_msg=False, timeout=60, delay_time=0):
    """
    填写填入框内容
    :param element_name:
    :param selector:
    :param value:
    :param is_msg:系统信息
    :return:
    """
    try:
        element = wait_element(selector=selector, timeout=timeout).next_sibling
        element.click()
        if is_msg:
            element.set_text(text=value)
        else:
            time.sleep(0.5)
            keyboard.write(value, delay=delay_time)
        logger.info(f'元素:{element_name}填写成功,填写值:{value}')
    except Exception as e:
        logger.error(f'元素:{element_name}填写失败,填写值:{value}.error_msg:{traceback.format_exc()}')
        raise e


def click_image(element_name, image_path, double_click=False, position='center', accuracy=0.8):
    """
    图片点击
    :param element_name: 元素名称
    :param image_path: 点击图片
    :param double_click: 双击
    :param position: 点击位置
    :param accuracy: 识别准确率
    :return:
    """
    try:
        click_type = "double" if double_click else "single"
        OCR().image2roiClick(localImagePath=image_path, captureRange=None,
                             accuracy=0.8, buttonType="left", clickType="single", clickPosition="center",
                             horizontalOffset=0, verticalOffset=0, timeout=30)
        # OCR().image2roiClick(localImagePath=image_path, captureRange=None, accuracy=accuracy, buttonType="left",
        #                      clickType=click_type, clickPosition=position,
        #                      horizontalOffset=0, verticalOffset=0, timeout=30)
        logger.info(f"点击元素:{element_name}成功！")
    except Exception as e:
        logger.error(f"点击元素:{element_name}失败！")
        raise e


def send_keys(keyword, times, wait=0.0):
    """
    发送按键
    :return:
    """
    for _ in range(times):
        keyboard_action.keyboard_send(hotkey=keyword, auxiliary_key="")
        time.sleep(wait)
    logger.info(f"发送快捷键:{keyword} 成功。")


def mouse_click(element_name, x, y, absolute=True, mouse_button='left', count=1, delay=0, aux_key=""):
    try:

        mouse_action.mouse_click(
            x=x,
            y=y,
            absolute=absolute,
            mouse_button=mouse_button,
            count=count,
            delay=delay,
            auxiliary_key=aux_key
        )
    except Exception as e:
        logger.info(f"点击{element_name}失败")
        raise e
    else:
        logger.info(f"点击{element_name}成功")


if __name__ == '__main__':
    from rpa.ui import element_action
    from rpa.ui.element_driver import ui_driver

    pass
