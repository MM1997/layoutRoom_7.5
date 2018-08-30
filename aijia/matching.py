from pywinauto import application
import time
from lib.public import matchimg, findwindow
from aijia import appaj
from lib.pic import *
import pyautogui
from pywinauto.base_wrapper import BaseWrapper
import jsonData

def check(findpic=DREAMROOM, timeout=60, checkfreq=1,confidence=0.85):
    """
    :type appwindow:BaseWrapper
    :param findpic:需要查找的图片
    :param timeout: 超时时间
    :param checkfreq:检查频率
    :return:检查当前图片是否存在
    :rtype:bool
    """
    while timeout > 0:
        search_match = matchimg(picturename, findpic,confidence=0.85)
        if search_match:
            return True
        else:
            time.sleep(checkfreq)
            timeout = timeout - checkfreq
    else:
        return False
startTime = time.time()
# check(findpic=search)
time.sleep(2)
# 结束时间
endTime = time.time()
# 总耗时
t = endTime - startTime
currentTime2 = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
print("结束时间：%s" % currentTime2)
print("总耗时:%s" % t)
def log(content):
    currentTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    with open("log.txt","a+",encoding="utf-8") as f:
        f.write(currentTime + " " + content + "\n")

a = ["1","1","1","1","1","1","1","1","1"]
b = ["1","1","1","1","1"]
len_b = len(b)
a = a[len_b-1:]
print("ddddd:%s" %a)
solutionId = "222"
e = "222"
log("在进入产品方案下有异常,solutionId：%s，报错信息：%s" %(solutionId,e))
