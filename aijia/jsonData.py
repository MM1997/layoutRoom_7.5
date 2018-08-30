import time
import json
from aijia import appaj

dic = {
    "time": "",
    "solutionId": "",
    "dnaSolutionId": "",
    "area":{
        "livingRoom": {
            "description": "客厅",
            "num":"0",
            "data":{
                "before": {"objectNum": "0", "selectNum": "0"},
                "after": {"objectNum": "0", "selectNum": "0"}
            }
        },
        "restaurant": {
            "description": "餐厅",
            "num":"0",
            "data":{
                "before": {"objectNum": "0", "selectNum": "0"},
                "after": {"objectNum": "0", "selectNum": "0"}
            }
        },
        "masterBedroom": {
            "description": "主卧",
            "num":"0",
            "data":{
                "before": {"objectNum": "0", "selectNum": "0"},
                "after": {"objectNum": "0", "selectNum": "0"}
            }
        },
        "secondBedroom": {
            "description": "次卧",
            "num":"0",
            "data": {
                "before": {"objectNum": "0", "selectNum": "0"},
                "after": {"objectNum": "0", "selectNum": "0"}
            }
        },
        "kitchen": {
            "description": "厨房",
            "num":"0",
            "data": {
                "before": {"objectNum": "0", "selectNum": "0"},
                "after": {"objectNum": "0", "selectNum": "0"}
            }
        },
        "bathroom": {
            "description": "主卫",
            "num":"0",
            "data": {
                "before": {"objectNum": "0", "selectNum": "0"},
                "after": {"objectNum": "0", "selectNum": "0"}
            }
        }
    }
}


def saveTime(hash = dic):

    """
    将当前时间存在json数据里
    :param times: 当前时间值
    :param hash: json数据
    :return:
    """
    hash["time"] = appaj.getCurrentTime()
    # hash = json.dumps(hash,indent=2)
    return hash

def saveSolutionId(solutionId,hash = dic):

    """
    将产品方案id存到json数据里
    :param solutionId: 产品方案id
    :param hash: json数据
    :return:
    """
    hash["solutionId"] = solutionId
    # hash = json.dumps(hash,indent=2)
    return hash

def saveDnaSolutionId(dnaSolutionId,hash = dic):

    """
    将DNA方案id存到json数据里
    :param dnaSolutionId: DNA方案id
    :param hash: json数据
    :return:
    """
    hash["dnaSolutionId"] = dnaSolutionId
    # hash = json.dumps(hash,indent=2)
    return hash

def beforeObjectNum(property,data,hash = dic):
    """
    将智能套用前的物体个数存到对应得属性下
    :param property: 房间内的域，如：客厅、餐厅、主卧等
    :param data: 物体个数
    :param hash: json数据
    :return:
    """
    hash["area"][property]["data"]["before"]["objectNum"] = str(data)
    # hash = json.dumps(hash,indent=2)
    return hash

def beforeSelectNum(property,data,hash = dic):
    """
    将智能套用前的选择的物体个数存到对应得属性下
    :param property: 房间内的域，如：客厅、餐厅、主卧等
    :param data: 物体个数
    :param hash: json数据
    :return:
    """
    hash["area"][property]["data"]["before"]["selectNum"] = str(data)
    # hash = json.dumps(hash,indent=2)
    return hash

def afterObjectNum(property,data,hash = dic):
    """
    将智能套用后的物体个数存到对应得属性下
    :param property:房间内的域，如：客厅、餐厅、主卧等
    :param data:物体个数
    :param hash:json数据
    :return:
    """
    hash["area"][property]["data"]["after"]["objectNum"] = str(data)
    # hash = json.dumps(hash,indent=2)
    return hash

def afterSelectNum(property,data,hash = dic):
    """
    将智能套用后的选择的物体个数存到对应得属性下
    :param property:房间内的域，如：客厅、餐厅、主卧等
    :param data:物体个数
    :param hash:json数据
    :return:
    """
    hash["area"][property]["data"]["after"]["selectNum"] = str(data)
    # hash = json.dumps(hash,indent=2)
    return hash

def getRoomAttribute(property,hash = dic):
    """
    获取房间内的域
    :param property:房间内的域，如：客厅、餐厅、主卧等
    :param hash:json数据
    :return:
    """
    attribute = hash["area"][property]["description"]
    return attribute


#
# print(afterObjectNum("bathroom",2222))
#
#
# print(saveSolutionId("10980"))
# print(saveDnaSolutionId("10980"))
# print(getRoomAttribute("secondBedroom"))
