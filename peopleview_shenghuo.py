import time

import pymongo
import xlsxwriter
from pymongo.errors import DuplicateKeyError

metaData = {"errCode":0,"errMsg":"ok","data":[{"tagType":1383,"tagTypeName":"生活兴趣","tag":"lr_life_interest","total":85454,"tagData":[{"cnt":14603,"rate":17.09,"tgi":-1,"tagValue":"2","tagValueName":"理想家","tagValueSortIndex":4426},{"cnt":13765,"rate":16.11,"tgi":-1,"tagValue":"1","tagValueName":"囤货小当家","tagValueSortIndex":4425},{"cnt":6153,"rate":7.2,"tgi":-1,"tagValue":"6","tagValueName":"旅行家","tagValueSortIndex":4429},{"cnt":4551,"rate":5.33,"tgi":-1,"tagValue":"10","tagValueName":"DIY达人","tagValueSortIndex":4433},{"cnt":4423,"rate":5.18,"tgi":-1,"tagValue":"7","tagValueName":"雅致居家控","tagValueSortIndex":4430},{"cnt":3979,"rate":4.66,"tgi":-1,"tagValue":"24","tagValueName":"品味家","tagValueSortIndex":4443},{"cnt":3672,"rate":4.3,"tgi":-1,"tagValue":"3","tagValueName":"装修家","tagValueSortIndex":4427},{"cnt":3041,"rate":3.56,"tgi":-1,"tagValue":"15","tagValueName":"杯子控","tagValueSortIndex":4437},{"cnt":2817,"rate":3.3,"tgi":-1,"tagValue":"25","tagValueName":"木作匠人","tagValueSortIndex":4444},{"cnt":2537,"rate":2.97,"tgi":-1,"tagValue":"4","tagValueName":"爱车一族","tagValueSortIndex":4428},{"cnt":2240,"rate":2.62,"tgi":-1,"tagValue":"16","tagValueName":"多肉控","tagValueSortIndex":4438},{"cnt":2200,"rate":2.57,"tgi":-1,"tagValue":"12","tagValueName":"绿植控","tagValueSortIndex":4434},{"cnt":2127,"rate":2.49,"tgi":-1,"tagValue":"21","tagValueName":"手工匠人","tagValueSortIndex":4441},{"cnt":2044,"rate":2.39,"tgi":-1,"tagValue":"19","tagValueName":"毛驴党","tagValueSortIndex":4440},{"cnt":1982,"rate":2.32,"tgi":-1,"tagValue":"23","tagValueName":"美式家","tagValueSortIndex":4442},{"cnt":1890,"rate":2.21,"tgi":-1,"tagValue":"26","tagValueName":"盘子控","tagValueSortIndex":4445},{"cnt":1006,"rate":1.18,"tgi":-1,"tagValue":"8","tagValueName":"智慧家","tagValueSortIndex":4431},{"cnt":887,"rate":1.04,"tgi":-1,"tagValue":"14","tagValueName":"水族爱好者","tagValueSortIndex":4436},{"cnt":849,"rate":0.99,"tgi":-1,"tagValue":"28","tagValueName":"古典匠人","tagValueSortIndex":4446},{"cnt":644,"rate":0.75,"tgi":-1,"tagValue":"9","tagValueName":"汪星人","tagValueSortIndex":4432},{"cnt":337,"rate":0.39,"tgi":-1,"tagValue":"17","tagValueName":"喵星人","tagValueSortIndex":4439},{"cnt":326,"rate":0.38,"tgi":-1,"tagValue":"13","tagValueName":"机车骑士","tagValueSortIndex":4435},{"cnt":275446,"rate":100,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":4424}],"sr":0,"chartType":"bar","tips":"通过消费者在淘宝天猫上有强互动行为的商品预测得出","valueSortType":2},{"tagType":1384,"tagTypeName":"母婴相关","tag":"lr_maternal_child","total":85454,"tagData":[{"cnt":34834,"rate":40.77,"tgi":-1,"tagValue":"17","tagValueName":"小小运动员","tagValueSortIndex":4412},{"cnt":24759,"rate":28.98,"tgi":-1,"tagValue":"12","tagValueName":"小正太","tagValueSortIndex":4409},{"cnt":18188,"rate":21.29,"tgi":-1,"tagValue":"15","tagValueName":"独立小萌宝","tagValueSortIndex":4411},{"cnt":9724,"rate":11.38,"tgi":-1,"tagValue":"9","tagValueName":"新生萌宝","tagValueSortIndex":4406},{"cnt":6058,"rate":7.09,"tgi":-1,"tagValue":"1","tagValueName":"家有萌娃","tagValueSortIndex":4402},{"cnt":5323,"rate":6.23,"tgi":-1,"tagValue":"29","tagValueName":"早教专家","tagValueSortIndex":4418},{"cnt":5171,"rate":6.05,"tgi":-1,"tagValue":"11","tagValueName":"小小二次元","tagValueSortIndex":4408},{"cnt":4494,"rate":5.26,"tgi":-1,"tagValue":"19","tagValueName":"小小画家","tagValueSortIndex":4414},{"cnt":4335,"rate":5.07,"tgi":-1,"tagValue":"33","tagValueName":"宝宝护理师","tagValueSortIndex":4419},{"cnt":3807,"rate":4.46,"tgi":-1,"tagValue":"18","tagValueName":"职场辣妈","tagValueSortIndex":4413},{"cnt":2772,"rate":3.24,"tgi":-1,"tagValue":"10","tagValueName":"宝宝营养师","tagValueSortIndex":4407},{"cnt":2605,"rate":3.05,"tgi":-1,"tagValue":"7","tagValueName":"小小军事迷","tagValueSortIndex":4405},{"cnt":2590,"rate":3.03,"tgi":-1,"tagValue":"24","tagValueName":"芭比收藏家","tagValueSortIndex":4416},{"cnt":2316,"rate":2.71,"tgi":-1,"tagValue":"3","tagValueName":"天才科学家","tagValueSortIndex":4403},{"cnt":1878,"rate":2.2,"tgi":-1,"tagValue":"35","tagValueName":"小小舞者","tagValueSortIndex":4421},{"cnt":1629,"rate":1.91,"tgi":-1,"tagValue":"14","tagValueName":"小小工程师","tagValueSortIndex":4410},{"cnt":1335,"rate":1.56,"tgi":-1,"tagValue":"4","tagValueName":"海派妈咪","tagValueSortIndex":4404},{"cnt":1112,"rate":1.3,"tgi":-1,"tagValue":"21","tagValueName":"小小赛车手","tagValueSortIndex":4415},{"cnt":530,"rate":0.62,"tgi":-1,"tagValue":"34","tagValueName":"小小厨师","tagValueSortIndex":4420},{"cnt":131,"rate":0.15,"tgi":-1,"tagValue":"36","tagValueName":"小小钢琴家","tagValueSortIndex":4422},{"cnt":104,"rate":0.12,"tgi":-1,"tagValue":"28","tagValueName":"备孕妈咪","tagValueSortIndex":4417},{"cnt":62,"rate":0.07,"tgi":-1,"tagValue":"37","tagValueName":"小小摄影师","tagValueSortIndex":4423},{"cnt":238206,"rate":100,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":4401}],"sr":0,"chartType":"bar","tips":"通过消费者在淘宝天猫上有强互动行为的商品预测得出","valueSortType":2}]}
client = pymongo.MongoClient(host='192.168.0.47', port=27017)
db= client.DataBank

collection = db.dataMerge
creatTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
data_dict = {}
data_dict["_id"]="数据融合" + str(creatTime)

def insert_item(collection, item):
    try:
        collection.insert_one(dict(item))
    except DuplicateKeyError:
        print('日期重复')
        pass
    except Exception as e:
        print('error!')
        print(e)



def data_merge():

    data = metaData["data"]

    #奶粉偏好包装
    pred_gender_data = data[0]["tagData"]
    data_dict["生活兴趣"]={}
    for i in pred_gender_data:
        data_dict["生活兴趣"][i["tagValueName"]] = str(i["rate"])+"%"

    #纸尿裤适用性别偏好
    pred_age_level_data =data[1]["tagData"]
    data_dict["母婴相关"] = {}
    for i in pred_age_level_data:
        data_dict["母婴相关"][i["tagValueName"]] = str(i["rate"]) + "%"


data_merge()

# 插入数据库
insert_item(db.dataMerge, data_dict)
