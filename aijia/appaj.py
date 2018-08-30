from typing import List
from lib.public import saveimage, matchimg, floattoint, WebRequests,savePic,log
import time
import pyautogui
from lib.pic import *
from pywinauto.base_wrapper import BaseWrapper
from lib.pic import *
import win32api
import win32con
import win32gui
from ctypes import *
import traceback
#
# def log(content):
#     currentTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
#     with open("../log/log.txt","a+",encoding="utf-8") as f:
#         f.write(currentTime + " " + content + "\n")


# 登录界面点击启动按钮
def start(appwindow, timeout=60, checkfreq=5):
    """
    :type appwindow: BaseWrapper
    :param timeout: 超时时间
    :param checkfreq:检查频率，每隔多久检查一次
    :return:
    """
    while timeout > 0:
        saveimage(appwindow, picturename)
        start_match = matchimg(picturename, startpic)
        if start_match:
            print(start_match)
            start_coordinate = floattoint(start_match.get('result'))
            print(start_coordinate)
            if start_coordinate:
                appwindow.click_input(coords=start_coordinate)
                return
        else:
            time.sleep(checkfreq)
            timeout = timeout - checkfreq


# 输入账号和密码登录
def login(appwindow, timeout=60, checkfreq=5):
    """
    :type appwindow: BaseWrapper
    :param timeout: 超时时间
    :param checkfreq:检查频率，每隔多久检查一次
    :return:
    """
    while timeout > 0:
        saveimage(appwindow, picturename)
        account_match = matchimg(picturename, accountpic)
        if account_match:
            appwindow.click()
            account_coordinate = floattoint(account_match.get('result'), xoffset=50)
            appwindow.click_input(coords=account_coordinate)
            appwindow.type_keys("^a{DELETE}" + account)
            time.sleep(0.5)
            passwd_match = matchimg(picturename, passwdpic)
            passwd_coordinate = floattoint(passwd_match.get('result'), xoffset=50)
            appwindow.click_input(coords=passwd_coordinate)
            appwindow.type_keys("^a{DELETE}" + passwd + "{ENTER}")
            time.sleep(0.5)
            login_match = matchimg(picturename, loginpic)
            login_coordinate = floattoint(login_match.get('result'))
            print(login_coordinate)
            appwindow.click_input(coords=login_coordinate)
            return
        else:
            time.sleep(checkfreq)
            timeout = timeout - checkfreq


def selectbycoords(appwindow, coords=(None, None), selectpic=selectpic_aijia, timeout=60,
                   checkfreq=1):
    """
    :type appwindow:BaseWrapper
    :param coords:下拉框当前坐标
    :param selectpic:下拉要选择的图片信息
    :param timeout:操作超时时间
    :param checkfreq:检查频率
    :return:选择框的坐标
    :rtype:tuple
    """
    while timeout > 0:
        appwindow.click_input(coords=coords)
        time.sleep(0.5)
        saveimage(appwindow, picturename)
        select_match = matchimg(picturename, selectpic)
        if select_match:
            select_coordinate = floattoint(select_match.get('result'))
            appwindow.click_input(coords=select_coordinate)
            return coords
        else:
            time.sleep(checkfreq)
            timeout = timeout - checkfreq


# 对菜单和下拉框点击
def select(appwindow, menupic=menupic_allcompany, selectpic=selectpic_aijia, timeout=60,
           checkfreq=1):
    """
    :type appwindow: BaseWrapper
    :param menupic:下拉选择框当前图片信息
    :param selectpic:下拉要选择的图片信息
    :param timeout:操作超时时间
    :param checkfreq:检查频率
    :return:选择框的坐标
    :rtype:tuple
    """
    if check(appwindow, findpic=selectpic,timeout=3):
        log('下拉框----方案id已经存在了')
    else:
        while timeout > 0:
            log('方案id存在了！！！！！！！！！！！')
            saveimage(appwindow, picturename)
            menu_match = matchimg(picturename, menupic)
            if menu_match:
                menu_match = floattoint(menu_match.get('result'))
                appwindow.click_input(coords=menu_match)
                appwindow.type_keys("{HOME}")
                time.sleep(0.5)
                saveimage(appwindow, picturename)
                select_match = matchimg(picturename, selectpic)
                if select_match:
                    select_coordinate = floattoint(select_match.get('result'))
                    appwindow.click_input(coords=select_coordinate)
                    return menu_match
                else:
                    time.sleep(checkfreq)
                    timeout = timeout - checkfreq
            else:
                time.sleep(checkfreq)
                timeout = timeout - checkfreq


# 输入框输入搜索
def search(appwindow, searchbutton=searchbuttonpic, text="8301", xoffset=0, yoffset=0, timeout=60, checkfreq=1):
    """
    :type appwindow: BaseWrapper
    :param searchbutton:搜索按钮图片信息
    :param text:搜索id信息
    :type text:str
    :param xoffset:  根据当前坐标x平移xoffset
    :param yoffset: 根据当前左边y平移yoffset
    :param timeout: 超时时间
    :param checkfreq:检查频率
    :return:
    """
    while timeout > 0:
        saveimage(appwindow, picturename)
        search_match = matchimg(picturename, searchbutton)
        if search_match:
            input_match = floattoint(search_match.get('result'), xoffset=xoffset, yoffset=yoffset)
            appwindow.click_input(coords=input_match)
            time.sleep(0.5)
            appwindow.type_keys("^a {DELETE} " + text)
            # appwindow.type_keys("^a")
            # appwindow.type_keys("^a")
            # time.sleep(0.2)
            # appwindow.type_keys("{DELETE}")
            # time.sleep(0.5)
            # appwindow.type_keys(text)
            
            search_match = floattoint(search_match.get('result'))
            appwindow.click_input(coords=search_match)
            log("输入框输入dnaId及点击搜索完成")
            return
        else:
            time.sleep(checkfreq)
            timeout = timeout - checkfreq


def getpiccoords(appwindow, findpic=DREAMROOM, xoffset=0, yoffset=0, timeout=60, checkfreq=1):
    """
    :type appwindow:BaseWrapper
    :param findpic:需要查找的图片
    :param xoffset:  根据当前坐标x平移xoffset
    :param yoffset: 根据当前左边y平移yoffset
    :param timeout: 超时时间
    :param checkfreq:检查频率
    :return:图片对应的坐标信息
    :rtype:tuple
    """
    while timeout > 0:
        saveimage(appwindow, picturename)
        search_match = matchimg(picturename, findpic)
        if search_match:
            input_match = floattoint(search_match.get('result'), xoffset=xoffset, yoffset=yoffset)
            return input_match
        else:
            time.sleep(checkfreq)
            timeout = timeout - checkfreq






def clickpic(appwindow, findpic=DREAMROOM, xoffset=0, yoffset=0, timeout=60, checkfreq=1):
    """
    :type appwindow:BaseWrapper
    :param findpic: 需点击的图片
    :param xoffset:  根据当前坐标x平移xoffset
    :param yoffset: 根据当前左边y平移yoffset
    :param timeout: 超时时间
    :param checkfreq:检查频率
    :return:
    """
    input_match = getpiccoords(appwindow, findpic=findpic, xoffset=xoffset, yoffset=yoffset, timeout=timeout,
                               checkfreq=checkfreq)
    time.sleep(0.5)
    if input_match:
        appwindow.click_input(coords=input_match)


def click_input(appwindow, button="left", coords=(None, None), button_down=False, button_up=False, pressed="",
                absolute=True, key_down=True, key_up=True):
    """
    :type appwindow:BaseWrapper
    :param button:
    :type coords:tuple
    :param button_down:
    :param button_up:
    :param pressed: str
    :param absolute:
    :param key_down:
    :param key_up:
    :return:
    """
    appwindow.click_input(button=button, coords=coords, button_down=button_down, button_up=button_up, pressed=pressed,
                          absolute=absolute, key_down=key_down, key_up=key_up)


def clickbycoords(appwindow, coords):
    """
    :type appwindow: BaseWrapper
    :type coords: tuple
    :return:
    """
    appwindow.click_input(coords=coords)


def pressbycoords(appwindow, coords):
    """
    :type appwindow: BaseWrapper
    :type coords: tuple
    :return:
    """
    appwindow.press_mouse_input(coords=coords)


def releasebycoords(appwindow, coords):
    """
    :type appwindow: BaseWrapper
    :type coords: tuple
    :return:
    """
    appwindow.release_mouse_input(coords=coords)


def movebycoords(appwindow, coords):
    """
    :type appwindow: BaseWrapper
    :type coords: tuple
    :return:
    """
    appwindow.move_mouse_input(coords=coords)


def dragbycoords(appwindow, srccoords, dstcoords):
    """
    :type appwindow: BaseWrapper
    :type srccoords: tuple
    :type dstcoords: tuple
    :return:
    """
    appwindow.drag_mouse_input(src=srccoords, dst=dstcoords)


def check(appwindow, findpic=DREAMROOM, timeout=60, checkfreq=1,confidence=0.85):
    """
    :type appwindow:BaseWrapper
    :param findpic:需要查找的图片
    :param timeout: 超时时间
    :param checkfreq:检查频率
    :return:检查当前图片是否存在
    :rtype:bool
    """
    while timeout > 0:
        saveimage(appwindow, picturename)
        # search_match = matchimg(picturename, findpic)
        search_match = matchimg(picturename, findpic, confidence=confidence)
        print(search_match)
        if search_match:
            return True
        else:
            time.sleep(checkfreq)
            timeout = timeout - checkfreq
    else:
        return False




# 下载方案后打开方案
def downloadandopenprogramme(appwindow, menupic, selectpic, programmetext):
    """
    :type appwindow: BaseWrapper
    :param menupic:下拉选择框图片
    :param selectpic:选择搜索条件图片
    :param programmetext:方案id或方案名称
    :return:
    """
    time.sleep(0.5)
    select(appwindow, menupic=menupic, selectpic=selectpic, timeout=3)
    if isinstance(programmetext, int):
        search(appwindow, text=str(programmetext), xoffset=-100)
    else:
        search(appwindow, text=programmetext, xoffset=-100)

    the1coords = getpiccoords(appwindow, findpic=DREAMROOM, xoffset=50, yoffset=220)
    the2coords = getpiccoords(appwindow, findpic=DREAMROOM)
    time.sleep(0.5)
    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    time.sleep(0.5)
    # 容错处理，找不到对应id信息
    if not check(appwindow, findpic=点击下载, timeout=1) and not check(appwindow, findpic=打开方案, timeout=1):
        return
    if check(appwindow, findpic=点击下载, timeout=1):
        clickbycoords(appwindow, the1coords)
        timeout = 1800
        while timeout > 0:
            if check(appwindow, findpic=打开方案, timeout=1):
                clickbycoords(appwindow, the1coords)
                break
            # 容错处理，当点击下载，就直接进入
            if check(appwindow, findpic=方案单品, timeout=1):
                # 容错处理，当退出按钮点击失败
                exittimeout = 180
                while exittimeout > 0:
                    if check(appwindow, findpic=方案退出确定, timeout=1):
                        clickpic(appwindow, findpic=方案退出确定, timeout=1)
                        time.sleep(0.5)
                        if check(appwindow, findpic=DREAMROOM, timeout=1):
                            print("已退出到主界面")
                            # 重新打开方案
                            select(appwindow, menupic=menupic, selectpic=selectpic, timeout=3)
                            if isinstance(programmetext, int):
                                search(appwindow, text=str(programmetext), xoffset=-100)
                            else:
                                search(appwindow, text=programmetext, xoffset=-100)
                            clickbycoords(appwindow, the1coords)
                            break
                    else:
                        clickpic(appwindow, findpic=方案退出, timeout=1)
                        time.sleep(1)
                        exittimeout = exittimeout - 1

            pyautogui.moveTo(the2coords[0], the2coords[1], duration=0.25)
            time.sleep(0.5)
            pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            time.sleep(0.5)
            timeout = timeout - 1
    else:
        if check(appwindow, findpic=打开方案, timeout=1):
            clickbycoords(appwindow, the1coords)


# 打开户型id进入绘制界面
def opendrawid(appwindow, housingid=583):
    """
    :type appwindow:BaseWrapper
    :param housingid: 户型id
    :return:
    """
    clickpic(appwindow, findpic=项目户型)
    time.sleep(0.5)
    select(appwindow, menupic=户型名称, selectpic=户型ID, timeout=3)
    search(appwindow, text=str(housingid), xoffset=-100)
    the1coords = getpiccoords(appwindow, findpic=DREAMROOM, xoffset=50, yoffset=220)
    the2coords = getpiccoords(appwindow, findpic=DREAMROOM)
    time.sleep(0.5)
    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    time.sleep(0.5)

    if check(appwindow, findpic=下载户型, timeout=1):
        clickbycoords(appwindow, the1coords)
        timeout = 180
        while timeout > 0:
            if check(appwindow, findpic=修改户型, timeout=1):
                clickbycoords(appwindow, the1coords)
                timeout = 10
                while timeout > 0:
                    if check(appwindow, findpic=工具修改, timeout=1):
                        print("进入绘制界面")
                        break
                    else:
                        timeout = timeout - 1
                        time.sleep(1)
                # 退出外部循环
                break
            pyautogui.moveTo(the2coords[0], the2coords[1], duration=0.25)
            time.sleep(0.5)
            pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            time.sleep(0.5)
            timeout = timeout - 1
    elif check(appwindow, findpic=修改户型, timeout=1):
        clickbycoords(appwindow, the1coords)
        timeout = 10
        while timeout > 0:
            if check(appwindow, findpic=工具修改, timeout=1):
                print("进入绘制界面")
                break
            else:
                timeout = timeout - 1
                time.sleep(1)
    elif check(appwindow, findpic=绘制户型, timeout=1):
        clickbycoords(appwindow, the1coords)
        timeout = 10
        while timeout > 0:
            if check(appwindow, findpic=工具修改, timeout=1):
                print("进入绘制界面")
                break
            else:
                timeout = timeout - 1
                time.sleep(1)


def submitbysearchid(appwindow, menupic, selectpic, programmetext=11066):
    downloadandopenprogramme(appwindow, menupic, selectpic, programmetext=programmetext)
    saveprogramme(appwindow)
    exitprogramme(appwindow)


def saveprogramme(appwindow):
    if check(appwindow, findpic=方案保存, timeout=600):
        time.sleep(3)
        clickpic(appwindow, findpic=方案保存)

    sbtimeout = 300
    while sbtimeout > 0:
        if check(appwindow, findpic=正在上传缩略图, timeout=1):
            time.sleep(1)
            sbtimeout = sbtimeout - 1
        else:
            print("方案已保存完成")
            break
    time.sleep(0.5)


def exitprogramme(appwindow):
    log(">>>>>>>>步骤：方案退出.....")
    exittimeout = 2
    # 容错处理，当退出按钮点击失败
    try:
        while exittimeout > 0:
            if check(appwindow, findpic=方案退出确定, timeout=1):
                log("存在方案退出确定按钮.....")
                clickpic(appwindow, findpic=方案退出确定, timeout=1)
                the1coords = getpiccoords(appwindow, findpic=方案退出确定, xoffset=50, yoffset=220)
                pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
                time.sleep(1)
                log("在匹配dreamroom。。。。。。")
                if check(appwindow, findpic=DREAMROOM, timeout=1):
                    log("匹配了dreamroom，方案退出完成")
                    break
            else:
                if not check(appwindow, findpic=方案退出确定, timeout=1):
                    clickpic(appwindow, findpic=方案退出, timeout=1)
                    the1coords = getpiccoords(appwindow, findpic=方案退出确定, xoffset=50, yoffset=220)
                    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
                    clickpic(appwindow, findpic=方案退出确定, timeout=1)

                    if check(appwindow, findpic=DREAMROOM, timeout=1):
                        log("匹配了dreamroom，方案退出完成")
                        break
                    time.sleep(1)
                    exittimeout = exittimeout - 1
                else:
                    log("匹配了dreamroom，方案退出完成")
                    break
    except:
        log("方案退出异常，报错信息:{}".format(traceback.format_exc()))

def exit(appwindow):
    # the1coords = getpiccoords(appwindow, findpic=方案退出确定, xoffset=50, yoffset=220, timeout=2)
    if check(appwindow, findpic=方案退出, timeout=3):
        print("方案退出确定,,,,,,,,,,,")


def getids(url="http://192.168.1.32:10026/dr-web/solution/queryByCondition", pagesize=10):
    para = {
        "pageNo": 1,
        "pageSize": pagesize,
        "type": 1
    }
    headers = {"Content-Type": "application/json; charset=utf-8"}
    httpreq = WebRequests()
    result = httpreq.post_json(url, para, headers)
    # print(result)
    totalPage = result['data']['totalPage']
    para["pageSize"] = pagesize
    listids: List[int] = []
    for i in range(totalPage):
        para["pageNo"] = i + 1
        result = httpreq.post_json(url, para, headers)
        for j in range(pagesize):
            try:
                print(result['data']["list"][j]['id'])
                # if result['data']["list"][j]['id']<=8984:
                listids.append(result['data']["list"][j]['id'])
            except:
                print("complete")
    return listids


def drawquickwall(appwindow, wallpoints=((500, 500), (1000, 500), (1000, 700), (500, 700), (500, 500))):
    clickpic(appwindow, 工具快速墙体绘制)
    movebycoords(appwindow, coords=wallpoints[0])
    time.sleep(1)
    for point in wallpoints:
        movebycoords(appwindow, coords=point)
        click_input(appwindow, coords=point, button_down=True, button_up=False)
        time.sleep(1)
        click_input(appwindow, coords=point, button_down=False, button_up=True)
        time.sleep(1)
    clickpic(appwindow, findpic=工具修改)
    movebycoords(appwindow, coords=wallpoints[0])


def drawwall(appwindow, wallpoints=((500, 500), (1000, 500))):
    if not check(appwindow, 工具墙体绘制选中, timeout=1):
        clickpic(appwindow, 工具墙体绘制)
    movebycoords(appwindow, coords=wallpoints[0])
    dragbycoords(appwindow, wallpoints[0], wallpoints[1])


def drawdelete(appwindow, areapoints=((500, 500), (1000, 500))):
    clickpic(appwindow, findpic=工具删除)
    movebycoords(appwindow, coords=areapoints[0])
    dragbycoords(appwindow, areapoints[0], areapoints[1])


def drawarea(appwindow, areapoints=((510, 210), (990, 210), (990, 390), (510, 390), (510, 210))):
    clickpic(appwindow, 工具绘制区域)
    movebycoords(appwindow, coords=areapoints[0])
    time.sleep(1)
    for point in areapoints:
        movebycoords(appwindow, coords=point)
        pressbycoords(appwindow, coords=point)
        time.sleep(1)
        releasebycoords(appwindow, coords=point)
        time.sleep(1)
    clickpic(appwindow, findpic=工具修改)
    movebycoords(appwindow, coords=areapoints[0])


def dragobjecttoarea(appwindow, findpic, doorarea=(500, 500)):
    doorcoords = getpiccoords(appwindow, findpic=findpic)
    movebycoords(appwindow, coords=doorcoords)
    click_input(appwindow, coords=doorcoords, button_down=True, button_up=False)
    time.sleep(1)
    dragbycoords(appwindow, doorcoords, doorarea)
    time.sleep(1)


def pressarea(appwindow, area=(500, 500)):
    movebycoords(appwindow, coords=area)
    click_input(appwindow, coords=area, button_down=True, button_up=False)
    time.sleep(1)
    click_input(appwindow, coords=area, button_down=False, button_up=True)



def enter_dna_programme(appwindow, find_pic):
    '''
    进入DNA方案下
    '''
    clickpic(appwindow, findpic=find_pic, timeout=1)



def check_solutionId_exist(appwindow,menupic, selectpic,programmetext):
    log(">>>>>>>>步骤：开始检查solutionId是否存在....")
    # 下拉框选择方案ID
    select(appwindow, menupic=menupic, selectpic=selectpic, timeout=30)
    time.sleep(1)
    # 输入框输入及点击搜索
    search(appwindow, searchbutton=searchbuttonpic, text=programmetext, xoffset=-100)
    the1coords = getpiccoords(appwindow, findpic=DREAMROOM, xoffset=50, yoffset=220)
    time.sleep(0.5)
    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    time.sleep(1.5)
    # 容错处理，找不到对应id信息

    for i in range(4):
        if not check(appwindow, findpic=点击下载, timeout=1) and not check(appwindow, findpic=打开方案, timeout=1):
            search(appwindow, searchbutton=searchbuttonpic, text=programmetext, xoffset=-100)
            time.sleep(0.5)
            pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            if i == 3:
                log("solutionId不存在")
                savePic(appwindow,"solutionId不存在")
                return False
                break
        else:
            log("solutionId存在")
            return True

def check_dnaSolutionId_exist(appwindow,findpic,programmetext):
    # search(appwindow, searchbutton=search1, text=programmetext, xoffset=-100)
    log("开始检查dnaId是否存在.....")
    time.sleep(2)
    the1coords = getpiccoords(appwindow, findpic=search1, xoffset=50, yoffset=220)
    time.sleep(0.5)
    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    for i in range(10):
        if not check(appwindow, findpic=findpic, timeout=1) and check(appwindow, findpic=首页尾页_DNA方案, timeout=1):
            if not check(appwindow, findpic=pic_enter_dnaid, timeout=1):
                return True
            elif check(appwindow, findpic=pic_enter_dnaid, timeout=1):
                # clickpic(appwindow, findpic=DNA返回DNA首页, xoffset=0, yoffset=0, timeout=1, checkfreq=1)
                log("检查dnaId是否存在时，已经在打开dna方案的界面，请检查....")
                savePic(appwindow,"检查dnaId是否存在时，已经在打开dna方案的界面")
                return_dna(appwindow)
                return False

        elif check(appwindow, findpic=findpic, timeout=1): #searchbuttonpic
            search(appwindow, searchbutton=search1, text=programmetext, xoffset=-100)
            time.sleep(1)
            pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            if i == 2:
                log("dnaId不存在.....")
                savePic(appwindow,"dnaId不存在")
                #返回DNA首页
                return_dna(appwindow)
                return False
        else:
            search(appwindow, searchbutton=search1, text=programmetext, xoffset=-100)
            time.sleep(1)
            pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            if i == 1:
                log("以上两种情况没出现，dnaId不存在.....")
                savePic(appwindow, "以上两种情况没出现，dnaId不存在")
                # 返回DNA首页
                return_dna(appwindow)
            elif i == 3:
                return False






# 下载方案
def download_programme(appwindow, menupic, selectpic, programmetext,xoffset=-100):
    """
    :type appwindow: BaseWrapper
    :param menupic:下拉选择框图片
    :param selectpic:选择搜索条件图片
    :param programmetext:方案id或方案名称
    :return:
    """
    log(">>>>>>>>步骤：下载或者打开产品方案....")
    the1coords = getpiccoords(appwindow, findpic=DREAMROOM, xoffset=50, yoffset=220)
    the2coords = getpiccoords(appwindow, findpic=DREAMROOM)

    if check(appwindow, findpic=点击下载, timeout=1):
        log("有点击下载按钮，开始点击下载....")
        clickbycoords(appwindow, the1coords)
        timeout = 180
        while timeout > 0:
            if check(appwindow, findpic=打开方案, timeout=1):
                log("存在打开方案按钮，开始打开方案....")
                clickbycoords(appwindow, the1coords)
                # 打开方案后检查是否跳转到指定页面（ihome页面）
                num = 300
                while num > 0:
                    if check(appwindow, findpic=DREAMROOM, timeout=1):
                        log("打开方案未完成，还在打开中.....")
                        num = num - 1
                        time.sleep(1)
                    elif check(appwindow, findpic=pic_openSolutionAfter, timeout=1):
                        log("打开方案完成.....")
                        break

            # 容错处理，当点击下载，就直接进入
            if check(appwindow, findpic=方案单品, timeout=1):
                break

            pyautogui.moveTo(the2coords[0], the2coords[1], duration=0.25)
            time.sleep(0.5)
            pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            time.sleep(0.5)
            timeout = timeout - 1
    else:
        log("首页已存在下载好的方案，开始打开方案....")
        clickbycoords(appwindow, the1coords)
        #打开方案后检查是否跳转到指定页面（ihome页面）
        num = 60
        while num > 0:
            if check(appwindow, findpic=DREAMROOM, timeout=1):
                log("打开方案未完成，还在打开中.....")
                num = num -1
                time.sleep(1)
            elif check(appwindow, findpic=pic_openSolutionAfter, timeout=1):
                log("打开方案完成.....")
                break


# 打开方案
def open_programme(appwindow, menupic, selectpic, programmetext,xoffset=-100):
    """
    :type appwindow: BaseWrapper
    :param menupic:下拉选择框图片
    :param selectpic:选择搜索条件图片
    :param programmetext:方案id或方案名称
    :return:
    """
    time.sleep(2)
    # 下拉框选择方案ID
    select(appwindow, menupic=menupic, selectpic=selectpic)
    time.sleep(2)
    # 输入框输入及点击搜索
    search(appwindow, searchbutton=searchbuttonpic, text=programmetext, xoffset=-100)

    the1coords = getpiccoords(appwindow, findpic=DREAMROOM, xoffset=50, yoffset=220)
    the2coords = getpiccoords(appwindow, findpic=DREAMROOM)
    time.sleep(0.5)
    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    time.sleep(1.5)
    # 容错处理，找不到对应id信息
    if not check(appwindow, findpic=点击下载, timeout=1) and not check(appwindow, findpic=打开方案, timeout=1):
        return
    if check(appwindow, findpic=点击下载, timeout=1):
        clickbycoords(appwindow, the1coords)
        timeout = 1800
        while timeout > 0:
            if check(appwindow, findpic=打开方案, timeout=1):
                # clickbycoords(appwindow, the1coords)
                break
            # 容错处理，当点击下载，就直接进入
            if check(appwindow, findpic=方案单品, timeout=1):
                # 容错处理，当退出按钮点击失败
                exittimeout = 180
                while exittimeout > 0:
                    if check(appwindow, findpic=方案退出确定, timeout=1):
                        clickpic(appwindow, findpic=方案退出确定, timeout=1)
                        time.sleep(0.5)
                        if check(appwindow, findpic=DREAMROOM, timeout=1):
                            print("已退出到主界面")
                            # 重新打开方案
                            select(appwindow, menupic=menupic, selectpic=selectpic, timeout=3)
                            if isinstance(programmetext, int):
                                search(appwindow, text=str(programmetext), xoffset=-100)
                            else:
                                search(appwindow, text=programmetext, xoffset=-100)
                            clickbycoords(appwindow, the1coords)
                            break
                    else:
                        clickpic(appwindow, findpic=方案退出, timeout=1)
                        time.sleep(1)
                        exittimeout = exittimeout - 1

            pyautogui.moveTo(the2coords[0], the2coords[1], duration=0.25)
            time.sleep(0.5)
            pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
            time.sleep(0.5)
            timeout = timeout - 1
    else:
        pass
        # print("打开方案")
        # clickbycoords(appwindow, the1coords)


def enter_dnaSolution(appwindow):
    '''
        进入DNA方案下
    '''
    clickpic(appwindow, findpic=pic_dnaSolution, timeout=1)

def enter_Solution(appwindow):
    '''
        进入产品方案下
    '''
    log(">>>>>>>>步骤：点击产品方案")
    clickpic(appwindow, findpic=productsolution, timeout=1)

def enter_dna(appwindow):
    '''
        进入DNA下
    '''
    log(">>>>>>>>步骤：点击DNA")
    clickpic(appwindow, findpic=DNA, timeout=1)
    log("点击DNA结束")
    num = 10
    while num > 0:
        if check(appwindow, findpic=DNA返回DNA首页, timeout=60, checkfreq=1):
            break
        else:
            clickpic(appwindow, findpic=DNA, timeout=1)
            time.sleep(1)
            num = num -1

def  enter_dna1(appwindow):
    log(">>>>>>>>步骤：进入dna方案")
    time.sleep(1)
    the1coords = getpiccoords(appwindow, findpic=VR隐藏商品, xoffset=0, yoffset=0)
    appwindow.click_input(coords=(the1coords[0], the1coords[1]))
    # appwindow.clickpic(appwindow, findpic=VR隐藏商品, timeout=1)



def enter_solutionId(appwindow):
    '''
        进入方案id下
    '''
    log(">>>>>>>>步骤：下拉框选择方案Id")
    num = 30
    while num > 0:
        if check(appwindow, findpic=DNA下方案名称, timeout=1, checkfreq=1):
            log("页面显示方案名称，开始选择方案id")
            clickpic(appwindow, findpic=DNA下方案名称)
            clickpic(appwindow, findpic=DNA下拉选择方案ID)
            break
        elif check(appwindow, findpic=DNA下拉选择方案ID, timeout=1, checkfreq=1):
            log("页面直接显示方案id")
            break
        else:
            log("此页面没有方案名称和方案id")
            savePic(appwindow,"此页面没有方案名称和方案id")
            time.sleep(1)
            num = num - 1


def openSoltion_dna(appwindow):
    '''
        打开方案
    '''
    log(">>>>1")
    log(">>>>>>>>步骤：打开DNA方案...")



    if check(appwindow, findpic=DNA打开方案, timeout=2):
        log(">>>>2")
        log("存在DNA方案...")


        clickpic(appwindow, findpic=DNA打开方案, xoffset=0, yoffset=0, timeout=10, checkfreq=1)
        # time.sleep(3)
        num = 30
        while num > 0:
            if check(appwindow, findpic=pic_enter_dnaid, timeout=1):
                log(">>>>3")
                log("打开DNA方案完成！")
                break

            else:
                log(">>>>3")
                log("打开DNA方案未完成，重新打开DNA方案...")
                if num == 28:
                    savePic(appwindow,"打开DNA方案未完成，重新打开DNA方案")
                clickpic(appwindow, findpic=DNA打开方案, xoffset=0, yoffset=0, timeout=1, checkfreq=1)
                time.sleep(1)
                num = num -1
    else:
        log(">>>>3")
        log("不存在DNA方案...")
        savePic(appwindow,"不存在DNA方案")

def return_dna(appwindow):
    '''
        DNA  返回DNA首页
    '''

    log(">>>>>>>>步骤：返回DNA首页.....")
    num = 20
    while num > 0:
        if check(appwindow, findpic=pic_enter_dnaid, timeout=1):
            clickpic(appwindow, findpic=DNA返回DNA首页, xoffset=0, yoffset=0, timeout=60, checkfreq=1)
            time.sleep(1)
            #获取返回DNA首页的坐标
            the1coords = getpiccoords(appwindow, findpic=DNA返回DNA首页, xoffset=0, yoffset=50)
            if check(appwindow, findpic=pic_enter_dnaid, timeout=1):
                # 鼠标向下移动
                pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
                num = num - 1
                if num == 1:
                    log("返回DNA首页失败.....")
                    savePic(appwindow,"返回DNA首页失败")
                    break
        else:
            time.sleep(1)
            num = num - 1
            log("返回DNA首页完成.....")
            break



def getCurrentTime():
    currentTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    return currentTime



def close_pic(appwindow):
    setattr(matchimg, "confidence", 0.94)

    zhanKai_coords = getpiccoords(appwindow, findpic=批量下载和数据反馈, xoffset=-420, yoffset=-10)
    pyautogui.moveTo(zhanKai_coords[0], zhanKai_coords[1], duration=0.25)
    appwindow.click_input(coords=zhanKai_coords)

    num = 15
    while num > 0:

        zhanKai_coords1 = getpiccoords(appwindow, findpic=一个展开和一个收缩按钮, xoffset=-10, yoffset=10, timeout=2)
        print(zhanKai_coords1)
        zhanKai_coords = (zhanKai_coords[0], zhanKai_coords[1] + 40)
        if not check(appwindow, findpic=一个展开和一个收缩按钮, timeout=2):
            break
        else:
            appwindow.click_input(coords=zhanKai_coords1)



def delete_linmotu(appwindow):
    num = 15
    while num > 0:
        if not check(appwindow,findpic=临摹图, timeout=2):
            break
        else:
            num = num - 1
            #获取临摹图的坐标
            linmotu_coords = getpiccoords(appwindow, findpic=临摹图, xoffset=0, yoffset=130)
            # pyautogui.moveTo(linmotu_coords[0], linmotu_coords[1], duration=0.25)
            # pyautogui.moveTo(linmotu_coords[0], linmotu_coords[1], duration=0.25)
            appwindow.click_input(coords=linmotu_coords)
            if not check(appwindow,findpic=临摹图, timeout=2):
                break
            else:
                num = num - 1



#删除屋内所有物品
def delete_all_of_room(appwindow):
    """
        删除屋内所有物品
    :return:
    """
    log(">>>>>>>>步骤：删除屋内所有物品")
    #获取当前界面的大小
    size = pyautogui.size()
    coord = (size[0]-500,size[1]//2)
    appwindow.click_input(coords=coord)
    try:
        #模拟键盘m、h、ctrl-a、delete键
        time.sleep(2)
        #全选
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('a')
        pyautogui.keyUp('a')
        pyautogui.keyUp('ctrl')
        time.sleep(2)
        # 删除
        pyautogui.press('delete')
        time.sleep(2)
    except Exception as e:
        log("删除屋内所有物品报异常：{}".format(traceback.format_exc()))
        try:
            savePic(appwindow, "删除屋内所有物品报异常")
        except:
            log("删除屋内物品图片保存失败...{}".format(traceback.format_exc()))

    finally:
        log("删除屋内所有物品完成")

#清空屋内所有物品
def delete_all_of_room1(appwindow):
    log(">>>>>>>>步骤：清空房屋...")
    if check(appwindow, findpic=清空, timeout=2):
        log("点击清空...")
        clickpic(appwindow, findpic=清空, xoffset=0, yoffset=0, timeout=2, checkfreq=1)
        time.sleep(1)

    if check(appwindow, findpic=清空确定, timeout=1):
        log("确定清空！")
        the1coords = getpiccoords(appwindow, findpic=清空确定, xoffset=50, yoffset=220)
        pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
        clickpic(appwindow, findpic=清空确定, timeout=1)
        time.sleep(1)




    # log(">>>>>>>>步骤：清除信息.....")
    # exittimeout = 1
    # # 容错处理，当退出按钮点击失败
    # try:
    #     while exittimeout > 0:
    #         if check(appwindow, findpic=清空确定, timeout=1):
    #             log("存在清空确定按钮.....")
    #             clickpic(appwindow, findpic=清空确定, timeout=1)
    #             the1coords = getpiccoords(appwindow, findpic=清空确定, xoffset=50, yoffset=220)
    #             pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    #             time.sleep(1)
    #             # log("在匹配dreamroom。。。。。。")
    #             # if check(appwindow, findpic=DREAMROOM, timeout=1):
    #             #     log("匹配了dreamroom，清空完成")
    #             #     break
    #         else:
    #             if not check(appwindow, findpic=清空确定, timeout=1):
    #                 clickpic(appwindow, findpic=清空, timeout=1)
    #                 the1coords = getpiccoords(appwindow, findpic=清空确定, xoffset=50, yoffset=220)
    #                 pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    #                 clickpic(appwindow, findpic=清空确定, timeout=1)
    #                 time.sleep(1)
    #                 exittimeout = exittimeout - 1
    #             else:
    #                 log("匹配了dreamroom，清除完成")
    #                 break
    # except:
    #     log("清除异常，报错信息:{}".format(traceback.format_exc()))


def loop_mouse_click(appwindow):
    delete_linmotu(appwindow)
    # 获取当前界面的大小
    size = pyautogui.size()
    #x轴
    x = (int(size[0]*0.5), int(size[0]*0.8))
    #Y轴
    y = (int(size[1] * 0.5), int(size[1] * 0.8))
    appwindow.click_input(coords=(size[0]//2,size[1]//2))
    appwindow.click_input(coords=(size[0] // 2, size[1] // 2))

    # m = PyMouse()
    for y1 in range(y[0],y[1],80):
        for x1 in range(x[0],x[1],30):
            # time.sleep(1.5)
            # pyautogui.moveTo(x1, y1, duration=0.25)
            pyautogui.click(x=x1, y=y1,interval=1.5)
            # appwindow.click_input(coords=(x1, y1))
            # m.click(x1, y1,n=3)
            # pyautogui.moveTo(x1, y1, duration=0.25)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)  # 点击鼠标
            time.sleep(0.1)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)  # 抬起鼠标

    pyautogui.keyDown('h')
    pyautogui.keyUp('h')



def click_layoutWholeRoom(appwindow):
    log(">>>>>>>>步骤：全屋自动布局.....")
    try:
        for i in range(20):
            if check(appwindow, findpic=全屋自动布局, timeout=1):
                log("存在全屋自动布局按钮.....")
                #获取全屋自动布局按钮的坐标
                coords = getpiccoords(appwindow, findpic=全屋自动布局, xoffset=0, yoffset=-6)
                #鼠标移动到全屋自动布局按钮上
                pyautogui.moveTo(coords[0], coords[1], duration=0.25)
                # appwindow.click_input(coords=(coords[0], coords[1]))
                #点击按钮
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)  # 点击鼠标
                time.sleep(0.3)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)  # 抬起鼠标
                time.sleep(10)
                break
            else:
                log("没有全屋自动布局按钮，请检查.....")
                savePic(appwindow,"没有全屋自动布局按钮")
                #打开DNA方案
                # openSoltion_dna(appwindow)
                time.sleep(1)

        #检查全屋布局是否加载完成
        # for i in range(40):
        #     if check(appwindow, findpic=加载全屋布局, timeout=1, confidence=0.5):
        #         log("正在全屋布局中.....")
        #         for j in range(30):
        #             if not check(appwindow, findpic=加载全屋布局, timeout=1, confidence=0.5):
        #                 log("全屋自动布局完成")
        #                 savePic(appwindow,"全屋自动布局完成")
        #                 break
        #             else:
        #                 log("正在全屋布局中，等待中.....")
        #                 time.sleep(1)
        #         break
        #     else:
        #         log("全屋布局在加载中。。。")
        #         time.sleep(1)
        #         if i == 30:
        #             log("全屋布局加载时间过长，已取消加载。。。")
        #             savePic(appwindow,"全屋布局加载时间过长，已取消加载")
        #             pyautogui.keyDown('esc')
        #             time.sleep(0.5)
        #             pyautogui.keyUp('esc')
        #             time.sleep(1)
        #             break
    except Exception as e:
        log("全屋自动布局发生异常,异常信息：{}".format(traceback.format_exc()))
        time.sleep(3)



def closeDR(appwindow):
    #鼠标定位到关闭按钮
    the1coords = getpiccoords(appwindow, findpic=关闭按钮, xoffset=-10, yoffset=-8)
    # pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
    #点击关闭按钮
    appwindow.click_input(coords=(the1coords[0],the1coords[1]+25))

if __name__ == "__main__":
    pass
