from pywinauto import application
import time
from lib.public import matchimg, findwindow,savePic
from aijia import appaj
from lib.pic import *
import pyautogui
from pywinauto.base_wrapper import BaseWrapper
from  aijia.readExcle import ReadExcle
import os
import psutil
import signal
from lib.public import log
import copy
import traceback

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
   # exe = r"D:\wj\AI_0605_1722\WindowsNoEditor\ajdr.exe"
    # exe = r"D:\DR4AI20180515\WindowsNoEditor\ajdr.exe"
    # exe = r"F:/0725/WindowsNoEditor/ajdr.exe"
    exe = r"F:/ihomefnt/DreamRoom/Database/WindowsNoEditor/ajdr.exe"
    app.start(exe)

    # 线上版本需要开启
    # 找到title_name = r"艾佳生活"
    # appwindow: BaseWrapper = findwindow(app, titlename=title_name,timeout=3)
    # appaj.start(appwindow)

    # 登录时的title_name1 = "ihome"

    appwindow: BaseWrapper = findwindow(app, titlename=title_name1)
    appaj.login(appwindow)
    time.sleep(1)
    appwindow.maximize()
    return appwindow


# def solutionIdUsed(solutionIds,solutionId):
#     len_solutionId1 = len(solutionId1)
#     solutionIds = solutionIds[len_solutionId1-1:]
#     print("已经使用的solutionId：%s" % solutionId1)
#     print("剩下的solutionIds：%s" % solutionIds)
#     log("已经使用的solutionId：%s" % solutionId1)
#     log("剩下的solutionIds：%s" % solutionIds)
#     return solutionIds

def solutionIdUsed(solutionIds,solutionId):
    index1 = solutionIds.index(solutionId)
    solutionIds = solutionIds[index1:]
    print("solutionIds剩下：%s" %solutionIds)
    log("solutionIds剩下：%s" %solutionIds)
    return solutionIds

def dnaSolutionIdUsed(dnaSolutionIds,dna_id):
    index1 = dnaSolutionIds.index(dna_id)
    dnaSolutionIds = dnaSolutionIds[index1:]
    print("dnaSolutionId下剩下的dna_id：%s" %dnaSolutionIds)
    log("SolutionId下剩下的dna_id：%s" %dnaSolutionIds)
    return dnaSolutionIds

def killDR():
    time.sleep(5)
    pids = psutil.pids()
    for pid in pids:
        p = psutil.Process(pid)
        if p.name() == "ajdr.exe":
            try:
                cmd = 'taskkill /F /IM ajdr.exe'
                os.system(cmd)
            except:
                print("没有如此进程！！！")
    time.sleep(5)


def layoutWholeRoom(solutionIds,dnaSolutionIds):
    solutionId1 = []
    global counts
    global count
    global err_count
    global solutionId
    global dna_id
    try:
        # 重新启动
        try:
            myappwindow = connect()
        except:
            myappwindow = start()

        for solutionId in solutionIds:
            try:
                log("\n\n========================= solutionId:%s =========================" %solutionId)
                #记录solutionId使用次数
                err_count.setdefault(solutionId,0)
                count.setdefault(solutionId,0)
                if count[solutionId] == 0:
                    dnaSolutionIds = dna_ids
                count[solutionId] = count[solutionId] + 1

                # 判断页面是否在首页，否则返回首页
                # appaj.exitprogramme(myappwindow)
                # 首页点击产品方案
                appaj.enter_Solution(myappwindow)
                # 判断产品方案ID是否存在，不存在取下一个方案id
                if not appaj.check_solutionId_exist(myappwindow, pic_solutionName, pic_solutionId, solutionId):
                    log("solutionId：{} 不存在！".format(solutionId))
                    continue
                # 根据方案id搜索进入方案里
                appaj.download_programme(myappwindow, pic_solutionName, pic_solutionId, solutionId)

                appaj.enter_dna1(myappwindow)
                # 点击DNA
                appaj.enter_dna(myappwindow)
                # 下拉框选择方案id
                appaj.enter_solutionId(myappwindow)
                # 获取搜索按钮的坐标
                the1coords = appaj.getpiccoords(myappwindow, findpic=search, xoffset=50, yoffset=220)
            except Exception as e:
                log("在进入产品方案下有异常,solutionId：%s，报错信息：%s" %(solutionId,traceback.format_exc()))
                err_count[solutionId] = err_count[solutionId] + 1
                if err_count[solutionId] > 1:
                    solutionIds = solutionIdUsed(solutionIds, solutionId)  # 如果次数达到2此，solutionId下一个开始
                    solutionIds.pop(0)
                    log("如果次数达到2此，从下一个solutionId开始,solutionIds:{}".format(solutionIds))
                else:
                    solutionIds = solutionIdUsed(solutionIds, solutionId)
                try:
                    killDR()
                    log("已杀掉DR的进程！")
                except Exception as e:
                    log("杀掉DR的进程异常:{}".format(traceback.format_exc()))
                finally:
                    layoutWholeRoom(solutionIds, dnaSolutionIds)

            for dna_id in dnaSolutionIds:
                try:
                    log("\n------------------------------ solutionId:%s, dna_id:%s ------------------------------" %(solutionId,dna_id))
                    err_count.setdefault(dna_id,0)


                    # 输入框输入及点击搜索
                    appaj.search(myappwindow, searchbutton=search, text=dna_id, xoffset=-100)
                    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)

                    # 判断dnaId是否存在

                    # if not appaj.check_dnaSolutionId_exist(myappwindow, 未搜到结果, dna_id):
                    #     log("dna_id：{} 不存在！".format(dna_id))
                    #     continue


                    # DNA-->打开方案
                    appaj.openSoltion_dna(myappwindow)

                    # 鼠标向下移动
                    pyautogui.moveTo(the1coords[0], the1coords[1], duration=0.25)
                    # 删除屋内所有物品

                    appaj.delete_all_of_room1(myappwindow)
                    # 全屋自动布局
                    appaj.click_layoutWholeRoom(myappwindow)
                    # 点击 返回DNA首页
                    appaj.return_dna(myappwindow)

                except Exception as e:
                    log("全屋自动布局过程中有异常,dnsSolutionId：%s，报错信息：%s" % (dna_id, traceback.format_exc()))
                    print("全屋自动布局过程中有异常,dnsSolutionId：%s，报错信息：%s" % (dna_id, traceback.format_exc()))
                    appaj.return_dna(myappwindow)
                    continue
                counts = counts + 1
                log("总共执行了：{} 次".format(counts))
            appaj.exitprogramme(myappwindow) #退回首页

    except Exception as e:
        log("整个过程中有异常,报错信息:{}".format(traceback.format_exc()))
        err_count[dna_id] = err_count[dna_id] + 1
        log("dna_id长度是:{}".format(err_count[dna_id]))

        if err_count[dna_id] > 1:
            dnaSolutionIds = dnaSolutionIdUsed(dnaSolutionIds, dna_id) #如果次数达到2此，从下一个dnaid开始
            dnaSolutionIds.pop(0)
            log("如果次数达到2此，从下一个dnaid开始,dnaSolutionIds:{}".format(dnaSolutionIds))
        else:
            dnaSolutionIds = dnaSolutionIdUsed(dnaSolutionIds, dna_id)

        if  len(dnaSolutionIds) == 0:
            if len(solutionIds) != 0:
                solutionIds = solutionIdUsed(solutionIds, solutionId)  #如果dnaid为空了，就从下一个solution开始循环
                solutionIds.pop(0)
            dnaSolutionIds = dna_ids
            log("solutionIds;{}".format(solutionIds))
            log("dnaSolutionIds剩下;{}".format(dnaSolutionIds))
        else:
            solutionIds = solutionIdUsed(solutionIds,solutionId)

        try:
            killDR()
            print("已杀掉DR的进程！")
            log("已杀掉DR的进程！")
        except Exception as e:
            log("杀掉DR的进程异常:{}".format(traceback.format_exc()))
        finally:
            layoutWholeRoom(solutionIds, dnaSolutionIds)




if __name__ == "__main__":
    #获取solutionId和dnaSolutionId
    path = r'..\test2.xlsx'
    wb = ReadExcle(path)
    id = wb.get_solutionId()
    solutionIds = id["solutionId"]
    dnaSolutionIds = id["dnaSolutionId"]
    #记录solutionIds、dnaSolutionIds的值
    dnaSolutionIds = id["dnaSolutionId"]
    solutionIds= ['17355','17376']
    dnaSolutionIds = ['14194','13536']
    dna_ids = copy.deepcopy(dnaSolutionIds)

    log("\n\nsolutionIds：{}".format(solutionIds))
    log("dnaSolutionIds：{}".format(dnaSolutionIds))
    #记录每个solutionId循环出现的次数
    count = {}
    #报错的记录
    err_count = {}
    #执行总数
    counts = 0


    try:
        # 开始时间
        startTime = time.time()
        currentTime1 = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        log("开始时间：%s" %currentTime1)
        #执行的脚本
        layoutWholeRoom(solutionIds,dnaSolutionIds)
    except Exception as e:
        print("执行脚本报错：%s" %(traceback.format_exc()))

    finally:
        #结束时间
        endTime = time.time()
        currentTime2 = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        #总耗时
        t = endTime-startTime
        print("开始时间：%s" %currentTime1)
        print("结束时间：%s" %currentTime2)
        print("总耗时:%s" %t)

