import sys

import aircv as ac
import cv2
from PIL import ImageGrab
import time
import numpy as np
import xlrd
import pyautogui
import requests
import json
from lib.pic import *
from pywinauto.base_wrapper import BaseWrapper
from pywinauto import application
import os

def log(content):
    currentTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    date = time.strftime("%Y_%m_%d")
    with open("../log/{}.log".format(date),"a+",encoding="utf-8") as f:
        # f.write(currentTime + " " + content + "\n")
        f.write("{} {}\n".format(currentTime,content))

def dragwindow(appwindow, pic=ihomeicon):
    """
    :type appwindow: BaseWrapper
    :param pic:str
    :return:
    """
    saveimage(appwindow)
    result = matchimg(picturename, pic)
    if result:
        appwindow.click_input(coords=floattoint(result.get("result")))
        pyautogui.mouseDown()
        pyautogui.dragRel(-100, -100, duration=0.25)
        pyautogui.mouseUp()


def findwindow(app, titlename="艾佳生活", timeout=180):
    """
    :type app:application.Application
    :param titlename: str
    :param timeout:int
    :rtype: BaseWrapper
    """
    while True:
        if timeout > 0:
            try:
                App = app.connect(title=titlename)
                print("find window :" + titlename)
                App[titlename].set_focus()
                return App[titlename]
            except:
                time.sleep(1)
                timeout = timeout - 1
        else:
            print("timeout to find window")
            sys.exit()


def matchimgall(imgsrc, imgobj, confidence=0.85):  # imgsrc=原始图像，imgobj=待查找的图片
    """
    :param imgsrc:
    :param imgobj:
    :param confidence:
    :rtype:dict
    :return:

    Args:
        imgsrc(string): 图像、素材
        imgobj(string): 需要查找的图片
        confidence: 阈值，当相识度小于该阈值的时候，就忽略掉

    Returns:
        A tuple of found [(point, score), ...]
    """

    confidence = getattr(matchimg, "confidence", confidence)
    imsrc = cv2.imdecode(np.fromfile(imgsrc, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    imobj = cv2.imdecode(np.fromfile(imgobj, dtype=np.uint8), cv2.IMREAD_UNCHANGED)

    match_result = ac.find_all_template(imsrc, imobj, confidence)
    # if match_result is not None:
    #     match_result['shape'] = (imsrc.shape[1], imsrc.shape[0])  # 获取原始图片宽高 0为高，1为宽
    print(imgsrc, imgobj, match_result)
    return match_result


def matchimg(imgsrc, imgobj, confidence=0.85):  # imgsrc=原始图像，imgobj=待查找的图片
    """
     :param imgsrc:
     :param imgobj:
     :param confidence:
     :rtype:dict
     :return:

     Args:
         imgsrc(string): 图像、素材
         imgobj(string): 需要查找的图片
         confidence: 阈值，当相识度小于该阈值的时候，就忽略掉

     Returns:
         A tuple of found [(point, score), ...]
     """
    confidence = getattr(matchimg, "confidence", confidence)
    imsrc = cv2.imdecode(np.fromfile(imgsrc, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    imobj = cv2.imdecode(np.fromfile(imgobj, dtype=np.uint8), cv2.IMREAD_UNCHANGED)

    match_result = ac.find_template(imsrc, imobj, confidence)
    # if match_result is not None:
    #     match_result['shape'] = (imsrc.shape[1], imsrc.shape[0])  # 获取原始图片宽高 0为高，1为宽
    print(imgsrc, imgobj, match_result)
    # log("{} {} {}".format(imgsrc, imgobj, match_result))
    return match_result


def floattoint(changeobject, xoffset=0, yoffset=0):
    """
    :type changeobject: tuple
    :param xoffset:  根据当前坐标x平移xoffset
    :param yoffset: 根据当前左边y平移yoffset
    :return:把浮点型左边转换成整形坐标
    :rtype:tuple
    """
    # intx 横向坐标偏移量
    # inty 纵向坐标偏移量
    if isinstance(changeobject, tuple):
        a = int(changeobject[0]) + xoffset
        b = int(changeobject[1]) + yoffset
        resultobject = (a, b)
        return resultobject


def saveimage(appwindow, fullscreenpic=picturename):
    """
    :type appwindow: BaseWrapper
    :param fullscreenpic: 保存的图片路径和名字
    :return:
    """
    print(appwindow.rectangle())  # 获取应用图形矩阵

    left = appwindow.rectangle().left
    right = appwindow.rectangle().right
    top = appwindow.rectangle().top
    bottom = appwindow.rectangle().bottom

    bbox = (left, top, right, bottom)  # 图形矩阵
    img = ImageGrab.grab(bbox)
    img.save(fullscreenpic)
    # img.show()

def savePic(appwindow,info):
    """
    :type appwindow: BaseWrapper
    :return:
    """

    print(appwindow.rectangle())  # 获取应用图形矩阵

    left = appwindow.rectangle().left
    right = appwindow.rectangle().right
    top = appwindow.rectangle().top
    bottom = appwindow.rectangle().bottom

    bbox = (left, top, right, bottom)  # 图形矩阵
    img = ImageGrab.grab(bbox)

    date = time.strftime("%Y_%m_%d")
    if not os.path.exists(r"../log/logPic/{}".format(date)):
        os.makedirs(r"../log/logPic/{}".format(date))
    currentTime = time.strftime("%Y-%m-%d %H-%M-%S", time.localtime())
    img.save(r"../log/logPic/{}/{}-{}.jpg".format(date,info,currentTime))


class ParseExcel(object):
    def __init__(self, file_path):
        self.data = xlrd.open_workbook(file_path)

    def getsheet(self, sheetname):
        return self.data.sheet_by_name(sheetname)

    @classmethod
    def getdatafromexcel(cls, file_path, sheetname, xrow, ycol):
        """文件路径比较重要，要以这种方式去写文件路径不用"""
        # file_path = r'd:/功率因数.xlsx'
        # 读取的文件路径
        # file_path = file_path.decode('utf-8')
        # 文件中的中文转码
        data = xlrd.open_workbook(file_path)
        # 获取数据
        table = data.sheet_by_name(sheetname)
        # 获取sheet
        # nrows = table.nrows
        # # 获取总行数
        # ncols = table.ncols
        # 获取总列数
        # table.row_values(i)
        # # 获取一行的数值
        # table.col_values(i)
        # 获取一列的数值

        # 获取一个单元格的数值
        cell_value = table.cell(xrow, ycol).value

        return cell_value


class WebRequests:
    @staticmethod
    def get(url, para, headers):
        try:
            r = requests.get(url, params=para, headers=headers)
            print("获取返回的状态码", r.status_code)
            json_r = r.json()
            print("json类型转化成python数据类型", json_r)
            return json_r
        except BaseException as e:
            print("请求失败！", str(e))

    @staticmethod
    def post(url, para, headers):
        try:
            r = requests.post(url, data=para, headers=headers)
            print("获取返回的状态码", r.status_code)
            json_r = r.json()
            print("json类型转化成python数据类型", json_r)
            return json_r
        except BaseException as e:
            print("请求失败！", str(e))

    @staticmethod
    def post_json(url, para, headers):
        try:
            data = para
            data = json.dumps(data)  # python数据类型转化为json数据类型
            r = requests.post(url, data=data, headers=headers)
            print("获取返回的状态码", r.status_code)
            # print("返回值", r.content.decode("utf-8"))
            json_r = r.json()
            print("json类型转化成python数据类型", json_r)
            return json_r
        except BaseException as e:
            print("请求失败！", str(e))


if __name__ == "__main__":
    pass
    log(123)