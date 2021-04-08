"""
@file: uitest.py
@author: rrh
@time: 2021/4/8 10:47 下午
"""

# -*- coding: utf-8 -*-
"""
@version: 1.3
@v1.2 2020-12-2 加上了动作.txt的换行符处理
@v1.3 2020-12-3 加上了读入阈值和动作文档，如：UI工具包.exe 0.85 动作.txt。
                若无参数，默认为0.8和动作.txt。注意：阈值和文档顺序不能变，输入参数<=2个
                总之，参数有如下搭配：空 空、0.8 空、0.8 动作.txt
@author: lt
@software: Python
@file: UI工具包.py
@time: 2020-12-1
"""

import os
from os import path
import subprocess
import sys
import time
import cv2
import aircv as ac
import win32gui
import pyautogui
import pyperclip
import re
import numpy as np

from PIL import ImageGrab
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import *
import winshell

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # 获取项目的绝对路径

ACTIONCN2EN = {"匹配": "match", "输入内容": "input_type", "复制入内容": "input_copy", "左键单击": "click_once",
               "左键双击": "click_twice", "Enter键单击": "click_enter_once", "等待出现": "wait", "全选": "select_all",
               '点击按键': 'type_write', '检测结果': 'check_result', '鼠标滚动': 'scroll'}


def time_sleep(func):
    def func_wraper(*args, **kwargs):
        time.sleep(0.5)
        res = func(*args, **kwargs)
        time.sleep(0.5)
        return res

    return func_wraper


def check_result(target):
    match_res = None
    for i in range(60 * 2):
        scroll(-100)
        try:
            match_res = match(target)
            break
        except:
            match_res = None
        time.sleep(1)
    if not match_res:
        print('巡检失败')
        return False
    print('巡检成功')
    return True


##target:QQ-快捷方式，就会去找-->该文件所在目录的/图标与控件/QQ/快捷方式.png的123
def match(target, ignore_err=False):
    global ROOT_DIR
    # global threshold

    if '-' in target:
        target_path = target.split("-")[0]
        target_name = target.split("-")[1]

        appname = ROOT_DIR + '/testimage/' + target_path + '/' + target_name + '.png'
        appname_2 = ROOT_DIR + '/testimage/' + target_path + '/' + target_name + '_2.png'
        appname_3 = ROOT_DIR + '/testimage/' + target_path + '/' + target_name + '_3.png'
    else:
        appname = ROOT_DIR + '/testimage/' + target + '.png'
        appname_2 = ROOT_DIR + '/testimage/' + target + '_2.png'
        appname_3 = ROOT_DIR + '/testimage/' + target + '_3.png'

    ##根据桌面路径进行截屏
    ##获取当前桌面路径
    path = os.path.join(os.path.expanduser('~'), "Desktop")
    hwnd = win32gui.FindWindow(None, path)
    app = QApplication(sys.argv)
    screen = QApplication.primaryScreen()
    img = screen.grabWindow(hwnd).toImage()
    # img = ImageGrab.grab()

    ##保存截屏
    root_desk = ROOT_DIR + '/desktop.png'
    img.save(root_desk)
    ##等待图片保存
    time.sleep(0.5)
    # 支持图片名为中英文名，也支持路径中英文
    imsrc = cv2.imdecode(np.fromfile(root_desk, dtype=np.uint8), -1)
    imobj = cv2.imdecode(np.fromfile(appname, dtype=np.uint8), -1)

    # 匹配图标位置
    pos = ac.find_template(imsrc, imobj, threshold)

    if pos == None and os.path.exists(appname_2):
        print(u"第一张图标没匹配到，2秒后匹配第二张：")
        time.sleep(2)
        print(appname_2)
        imobj_2 = cv2.imdecode(np.fromfile(appname_2, dtype=np.uint8), -1)
        pos = ac.find_template(imsrc, imobj_2, threshold)
    if pos == None and os.path.exists(appname_3):
        print(u"第二张图标没匹配到，2秒后匹配第三张：")
        time.sleep(2)
        print(appname_3)
        imobj_3 = cv2.imdecode(np.fromfile(appname_3, dtype=np.uint8), -1)
        pos = ac.find_template(imsrc, imobj_3, threshold)

    # 如果第三张还未匹配到，用另一种方法重新截图
    if pos == None:
        print(u"2秒后重新截全屏...")
        time.sleep(2)
        ##保存截屏
        root_desk = ROOT_DIR + '/desktop_2.png'

        img = ImageGrab.grab()
        img.save(root_desk)
        ##等待图片保存
        time.sleep(0.5)
        # 支持图片名为中英文名，也支持路径中英文
        imsrc = cv2.imdecode(np.fromfile(root_desk, dtype=np.uint8), -1)
        imobj = cv2.imdecode(np.fromfile(appname, dtype=np.uint8), -1)
        # 匹配图标位置
        pos = ac.find_template(imsrc, imobj, threshold)
        if pos == None and os.path.exists(appname_2):
            print(u"第一张图标没匹配到，2秒后匹配第二张：")
            time.sleep(2)
            print(appname_2)
            imobj_2 = cv2.imdecode(np.fromfile(appname_2, dtype=np.uint8), -1)
            # imobj_2 = ac.imread(appname_2)
            pos = ac.find_template(imsrc, imobj_2, threshold)
        if pos == None and os.path.exists(appname_3):
            print(u"第二张图标没匹配到，2秒后匹配第三张：")
            time.sleep(2)
            print(appname_3)
            imobj_3 = cv2.imdecode(np.fromfile(appname_3, dtype=np.uint8), -1)
            # imobj_3 = ac.imread(appname_3)
            pos = ac.find_template(imsrc, imobj_3, threshold)

        if pos == None:
            print(u"第二种截图方式也没有匹配到：" + target)
            if not ignore_err:
                raise Exception('匹配失败')

    elif pos != None:
        x, y = pos['result']
        print(x, y, 'pos')
        pyautogui.moveTo(x=x, y=y, duration=0.5)
        print(u"匹配成功：{}，坐标位置为：{}".format(target, pyautogui.position()))
        time.sleep(0.5)
        return x, y


@time_sleep
def type_write(key):
    pyautogui.typewrite([key])


@time_sleep
def click_once():
    pyautogui.click()


@time_sleep
def click_twice():
    pyautogui.doubleClick()


@time_sleep
def scroll(move_length=-1000):
    time.sleep(0.5)
    pyautogui.scroll(move_length)
    time.sleep(0.5)


@time_sleep
def click_enter_once():
    pyautogui.typewrite(['enter'], 0.25)


def get_access_code(code_server):
    import requests
    try:
        res = requests.get(code_server).text
    except Exception as e:
        print(e)
        return ''
    res_list = list(filter(None, res.split('\n')))
    result_index = res_list.index('  <div class="message">')
    return res_list[result_index + 1].strip()


def input_type(msg=''):
    code_server = 'http://200.200.1.74:45678/api/v1/cloud/debug_qr_v2'
    code = get_access_code(code_server)
    if not msg:
        msg = code
    if not msg:
        raise Exception('写入内容不能为空，请传入写入内容')
    time.sleep(0.5)
    if isinstance(msg, float):
        msg = str(int(msg))
    pyautogui.typewrite(str(msg), 0.2)
    time.sleep(0.5)


@time_sleep
def input_copy(msg=""):
    time.sleep(0.5)
    pyperclip.copy(msg)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')


@time_sleep
def select_all():
    pyautogui.hotkey('ctrlleft', 'a')


def wait(target):
    ignore_err = False
    if ',' in target:
        target, ignore_err = target.split(',')
        ignore_err = int(ignore_err)
    for i in range(60 * 2):
        try:
            match_res = match(target, ignore_err=ignore_err)
            if match_res:
                # 如果匹配到了，直接结束
                return
            time.sleep(1)
        except:
            # 如果没匹配到，等一秒继续检测
            time.sleep(1)
    if not ignore_err:
        raise Exception('长时间没有出现预期的UI')


def create_shortcut_to_desktop(target, start_in, title):
    """创建快捷方式"""
    s = os.path.basename(target)
    fname = os.path.splitext(s)[0]
    winshell.CreateShortcut(Path=os.path.join(winshell.desktop(), fname + '.lnk'), Target=target, Icon=(target, 0),
                            Description=title, StartIn=start_in)


if __name__ == '__main__':
    target_path = ROOT_DIR + '\\test_exe\\acheck_full\\acheck\\aCheck.exe'
    start_in = ROOT_DIR + '\\test_exe\\acheck_full\\acheck'
    create_shortcut_to_desktop(target_path, start_in, 'acheck')
    time.sleep(3)

    threshold = 0.85

    file_name = 'action.txt'
    file = ROOT_DIR + '\\' + file_name

    try:
        with open(file, encoding='utf-8') as f:
            action_str = f.read()

        # 保证'动作.txt'里，动作最后的换行符无影响
        action_str = action_str.strip('\n')
        # 保证动作之间的单个换行符跟->等效。若有多个换行符，放到if act=''里
        action_str = action_str.replace('\n', '->')
        action_list = action_str.split('->')
        for act in action_list:
            if act == '':  # 当同时有多个换行符时，act就是空，在这里排除
                pass
            else:
                try:  # 防止同一Dj99ZzZz个action里，有一个动作异常就不执行其他动作
                    act.encode('utf-8')
                    ##提取action具体动作的括号里的内容，如match(QQ-单发输入框)
                    p1 = re.compile(r'[(](.*?)[)]', re.S)
                    params = re.findall(p1, act)  # 查找后返回list，都放在[0]里，()内无东西，即无参数，返回空列表
                    act = act.split('(')[0]

                    if params == []:
                        print(u"调用方法为：{}".format(act))
                        eval(ACTIONCN2EN[act])()
                    else:
                        params = params[0]
                        if params == 'arg1':
                            try:
                                params = sys.argv[1]
                            except IndexError as ex:
                                raise Exception('请传第1个参数')
                        if params == 'arg2':
                            try:
                                params = sys.argv[2]
                            except IndexError as ex:
                                raise Exception('请传第2个参数')
                        print(u"调用方法为：{}({})".format(act, params))
                        eval(ACTIONCN2EN[act])(params)  # 参数可以有多个，参数之间用,分隔。目前只支持一个参数
                except Exception as e:
                    print(e)
                    break
    except Exception as e:
        print(e)
