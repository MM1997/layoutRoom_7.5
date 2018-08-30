from typing import List
from lib.public import saveimage, matchimg, floattoint, WebRequests
import time
import pyautogui
from lib.pic import *
from pywinauto.base_wrapper import BaseWrapper
from lib.pic import *
import jsonData

def match1(findpic=DREAMROOM):
    """
    :type appwindow:BaseWrapper
    :param findpic:需要查找的图片
    :param timeout: 超时时间
    :param checkfreq:检查频率
    :return:检查当前图片是否存在
    :rtype:bool
    """
    search_match = matchimg(picturename, findpic)
    print(search_match)


def check1(findpic=DREAMROOM):
    """
    :type appwindow:BaseWrapper
    :param findpic:需要查找的图片
    :param timeout: 超时时间
    :param checkfreq:检查频率
    :return:检查当前图片是否存在
    :rtype:bool
    """

    search_match = matchimg(picturename, findpic)
    if search_match:
        print(123)
    else:
        print(567)


match1(DNA)
#pic_open_programme picturename


if check1(findpic=test):
    print(2)