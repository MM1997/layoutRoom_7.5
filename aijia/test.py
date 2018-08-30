from pywinauto import application
import time
from lib.public import matchimg, findwindow
from aijia import appaj
from lib.pic import *
import pyautogui
from pywinauto.base_wrapper import BaseWrapper
import os

def dna(appwindow, menupic, selectpic, programmetext, dnaid=9245):
    appaj.clickpic(appwindow, findpic=productsolution)
    time.sleep(1)
    # 测试select coords
    # allcompany_coords = appaj.select(appwindow, menupic=menupic_allcompany, selectpic=selectpic_aijia)
    # appaj.select(appwindow, coords=allcompany_coords, selectpic=menupic_allcompany)

    appaj.downloadandopenprogramme(appwindow, menupic, selectpic, programmetext)

    appaj.clickpic(appwindow, findpic=DNA)
    appaj.clickpic(appwindow, findpic=DNA下方案名称)
    appaj.clickpic(appwindow, findpic=DNA下拉选择方案ID)
    # appaj.search(appwindow, searchbuttonpic=DNA下搜索, text="8617", intx=-100)
    # time.sleep(0.5)
    # appaj.clickpic(appwindow, findpic=DNA打开方案)
    # time.sleep(0.5)
    # appaj.clickpic(appwindow, findpic=DNA全屋自动布局)
    # time.sleep(3)
    # appaj.clickpic(appwindow, findpic=DNA返回DNA首页)
    appaj.search(appwindow, searchbutton=DNA下搜索, text=str(dnaid), xoffset=-100)
    time.sleep(0.5)
    appaj.clickpic(appwindow, findpic=DNA打开方案)
    time.sleep(0.5)
    appaj.clickpic(appwindow, findpic=DNA全屋自动布局)
    time.sleep(3)
    appaj.clickpic(appwindow, findpic=DNA返回DNA首页)
    appaj.exitprogramme(appwindow)


def connect():
    app = application.Application()  # type:application.Application
    # 设置全局图片相似度
    setattr(matchimg, "confidence", 0.8)
    appwindow: BaseWrapper = findwindow(app, titlename=title_name1,timeout=3)
    time.sleep(1)
    appwindow.Maximize()
    return appwindow


def start():
    app = application.Application()  # type:application.Application
    # 设置全局图片相似度
    setattr(matchimg, "confidence", 0.85)
    exe = r"D:\wj\AI_0605_1722\WindowsNoEditor\ajdr.exe"
    # exe = r"F:\WindowsNoEditor\ajdr.exe"

    app.start(exe)

    # 线上版本需要开启
    # 找到title_name = r"艾佳生活"
    # appwindow: BaseWrapper = findwindow(app, titlename=title_name,timeout=3)
    # appaj.start(appwindow)

    # 登录时的title_name1 = "ihome"
    appwindow: BaseWrapper = findwindow(app, titlename=title_name1)
    appaj.login(appwindow)
    time.sleep(1)
    appwindow.Maximize()
    return appwindow




if __name__ == "__main__":
    try:
        myappwindow = connect()
    except:
        myappwindow = start()
    print(12345)
    print(type(myappwindow))
    time.sleep(2)

    for id in ["10980"]: #,"10967"
        # #首页点击产品方案
        # appaj.enter_Solution(myappwindow)
        # #根据方案id搜索进入方案里
        # appaj.download_programme(myappwindow, programme_name, programme_id, "10980")
        # #点击DNA
        # appaj.enter_dna(myappwindow)
        # #下拉框选择方案id
        # appaj.enter_solutionId(myappwindow)
        # # 获取搜索按钮的坐标
        # the1coords = appaj.getpiccoords(myappwindow, findpic=search, xoffset=50, yoffset=220)
        for dna_id in ["6008"]:#,"9010"
            # # 输入框输入及点击搜索
            # appaj.search(myappwindow, searchbutton=search, text=dna_id, xoffset=-100)
            # time.sleep(2)
            # # DNA-->打开方案
            # appaj.openSoltion_dna(myappwindow)
            #
            # # 鼠标向下移动
            # pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            #
            # # zhanKai_coords = appaj.getpiccoords(myappwindow, findpic=批量下载和数据反馈, xoffset=-420, yoffset=-10)
            # # print(zhanKai_coords)
            # # pyautogui.moveTo(zhanKai_coords[0], zhanKai_coords[1], duration=0.25)
            # # # (440, 237)
            # # myappwindow.click_input(coords=zhanKai_coords)
            # #
            # # pyautogui.moveTo(zhanKai_coords[0], zhanKai_coords[1]+40, duration=0.25)
            # #
            # # zhanKai_coords = appaj.getpiccoords(myappwindow, findpic=一个展开和一个收缩按钮, xoffset=-10, yoffset=10)
            # # print(zhanKai_coords)
            # # pyautogui.moveTo(zhanKai_coords[0], zhanKai_coords[1], duration=0.25)
            #
            # #在打开的dnaId界面下收缩房间内的各个域
            # appaj.close_pic(myappwindow)
            #
            #
            # #删除屋内所有物品
            # appaj.delete_all_of_room(myappwindow)
            # appaj.loop_mouse_click(myappwindow)


            try:
                cmd = 'taskkill /F /IM ajdr.exe'
                os.system(cmd)
            except:
                print("已杀掉DR的进程！")
            finally:
                myappwindow = start()



            # 点击 返回DNA首页
        #     appaj.return_dna(myappwindow)
        # appaj.exitprogramme(myappwindow)

