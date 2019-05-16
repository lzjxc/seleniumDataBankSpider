import time

import pymongo
import xlsxwriter
from pymongo.errors import DuplicateKeyError

metaData = {"errCode":0,"errMsg":"ok","data":[{"tagType":1090,"tagTypeName":"奶粉包装偏好","tag":"package_type_prefer_milk_powder","total":683625,"tagData":[{"cnt":35900,"rate":5.25,"tgi":-1,"tagValue":"2","tagValueName":"盒装","tagValueSortIndex":4148},{"cnt":160858,"rate":23.54,"tgi":-1,"tagValue":"3","tagValueName":"罐装","tagValueSortIndex":4149},{"cnt":1937,"rate":0.28,"tgi":-1,"tagValue":"4","tagValueName":"袋装","tagValueSortIndex":4150},{"cnt":232,"rate":0.03,"tgi":-1,"tagValue":"5","tagValueName":"试用装","tagValueSortIndex":4151},{"cnt":1312,"rate":0.19,"tgi":-1,"tagValue":"6","tagValueName":"箱装","tagValueSortIndex":4152},{"cnt":28930,"rate":4.23,"tgi":-1,"tagValue":"1","tagValueName":"其他","tagValueSortIndex":4153},{"cnt":454617,"rate":66.53,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":4154}],"sr":0,"chartType":"vertical_bar","tips":"","valueSortType":0},{"tagType":1091,"tagTypeName":"纸尿裤适用性别偏好","tag":"gender_prefer_diaper","total":683625,"tagData":[{"cnt":5414,"rate":0.79,"tgi":-1,"tagValue":"4","tagValueName":"男","tagValueSortIndex":4770},{"cnt":5939,"rate":0.86,"tgi":-1,"tagValue":"3","tagValueName":"女","tagValueSortIndex":4771},{"cnt":235494,"rate":34.26,"tgi":-1,"tagValue":"2","tagValueName":"男女通用","tagValueSortIndex":4772},{"cnt":9342,"rate":1.36,"tgi":-1,"tagValue":"1","tagValueName":"其他","tagValueSortIndex":4773},{"cnt":427696,"rate":62.22,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":4774}],"sr":0,"chartType":"vertical_bar","tips":"","valueSortType":0},{"tagType":1361,"tagTypeName":"辅食产地偏好","tag":"pref_food_supplement_place","total":683625,"tagData":[{"cnt":10743,"rate":1.57,"tgi":-1,"tagValue":"1","tagValueName":"国产","tagValueSortIndex":3749},{"cnt":88158,"rate":12.9,"tgi":-1,"tagValue":"2","tagValueName":"进口","tagValueSortIndex":3750},{"cnt":584652,"rate":85.53,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":3751}],"sr":0,"chartType":"vertical_bar","tips":"","valueSortType":0},{"tagType":1392,"tagTypeName":"辅食年消费频次","tag":"pay_ord_cnt_1y_complementary_food","total":683625,"tagData":[{"cnt":65110,"rate":9.52,"tgi":-1,"tagValue":"1","tagValueName":"1-2次","tagValueSortIndex":4004},{"cnt":25618,"rate":3.75,"tgi":-1,"tagValue":"2","tagValueName":"3-4次","tagValueSortIndex":4005},{"cnt":27155,"rate":3.97,"tgi":-1,"tagValue":"3","tagValueName":"5-9次","tagValueSortIndex":4006},{"cnt":22187,"rate":3.25,"tgi":-1,"tagValue":"4","tagValueName":"10次及以上","tagValueSortIndex":4007},{"cnt":543505,"rate":79.51,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":4008}],"sr":0,"chartType":"vertical_bar","tips":"","valueSortType":0},{"tagType":1009995,"tagTypeName":"童鞋最近6个月购买金额偏好","tag":"child_shoes_6m_amt","total":683625,"tagData":[{"cnt":273579,"rate":40.02,"tgi":-1,"tagValue":"1","tagValueName":"小于100","tagValueSortIndex":5591},{"cnt":117918,"rate":17.25,"tgi":-1,"tagValue":"2","tagValueName":"100-200","tagValueSortIndex":5592},{"cnt":125195,"rate":18.31,"tgi":-1,"tagValue":"3","tagValueName":"200-400","tagValueSortIndex":5593},{"cnt":50009,"rate":7.31,"tgi":-1,"tagValue":"4","tagValueName":"400-600","tagValueSortIndex":5594},{"cnt":21655,"rate":3.17,"tgi":-1,"tagValue":"5","tagValueName":"600-800","tagValueSortIndex":5595},{"cnt":10202,"rate":1.49,"tgi":-1,"tagValue":"6","tagValueName":"800-1000","tagValueSortIndex":5596},{"cnt":9143,"rate":1.34,"tgi":-1,"tagValue":"7","tagValueName":"1000-1500","tagValueSortIndex":5597},{"cnt":4415,"rate":0.65,"tgi":-1,"tagValue":"8","tagValueName":"1500以上","tagValueSortIndex":5598},{"cnt":71572,"rate":10.46,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":5599}],"sr":0,"chartType":"bar","tips":"","valueSortType":0},{"tagType":1010018,"tagTypeName":"童装最近6个月购买金额偏好","tag":"child_wear_6m_amt","total":683625,"tagData":[{"cnt":116474,"rate":17.04,"tgi":-1,"tagValue":"1","tagValueName":"小于100","tagValueSortIndex":5582},{"cnt":73204,"rate":10.71,"tgi":-1,"tagValue":"2","tagValueName":"100-200","tagValueSortIndex":5583},{"cnt":112591,"rate":16.47,"tgi":-1,"tagValue":"3","tagValueName":"200-400","tagValueSortIndex":5584},{"cnt":79522,"rate":11.63,"tgi":-1,"tagValue":"4","tagValueName":"400-600","tagValueSortIndex":5585},{"cnt":57208,"rate":8.37,"tgi":-1,"tagValue":"5","tagValueName":"600-800","tagValueSortIndex":5586},{"cnt":40855,"rate":5.98,"tgi":-1,"tagValue":"6","tagValueName":"800-1000","tagValueSortIndex":5587},{"cnt":60038,"rate":8.78,"tgi":-1,"tagValue":"7","tagValueName":"1000-1500","tagValueSortIndex":5588},{"cnt":70542,"rate":10.32,"tgi":-1,"tagValue":"8","tagValueName":"1500以上","tagValueSortIndex":5589},{"cnt":73074,"rate":10.7,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":5590}],"sr":0,"chartType":"bar","tips":"","valueSortType":0},{"tagType":1010028,"tagTypeName":"童装最近6个月消费频次偏好","tag":"child_wear_6m_cnt","total":683625,"tagData":[{"cnt":144489,"rate":21.16,"tgi":-1,"tagValue":"2","tagValueName":"[2-5)","tagValueSortIndex":5576},{"cnt":135476,"rate":19.84,"tgi":-1,"tagValue":"3","tagValueName":"[5-10)","tagValueSortIndex":5577},{"cnt":90386,"rate":13.24,"tgi":-1,"tagValue":"4","tagValueName":"[10-15)","tagValueSortIndex":5578},{"cnt":60153,"rate":8.81,"tgi":-1,"tagValue":"5","tagValueName":"[15-20)","tagValueSortIndex":5579},{"cnt":137687,"rate":20.16,"tgi":-1,"tagValue":"6","tagValueName":"20以上","tagValueSortIndex":5580},{"cnt":114678,"rate":16.79,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":5581}],"sr":0,"chartType":"bar","tips":"","valueSortType":0},{"tagType":1005,"tagTypeName":"宝宝年龄","tag":"pred_baby_age","total":683625,"tagData":[{"cnt":1573,"rate":0.23,"tgi":-1,"tagValue":"1","tagValueName":"0-3个月","tagValueSortIndex":3887},{"cnt":2032,"rate":0.3,"tgi":-1,"tagValue":"2","tagValueName":"3-6个月","tagValueSortIndex":3888},{"cnt":13645,"rate":1.99,"tgi":-1,"tagValue":"3","tagValueName":"6-12个月","tagValueSortIndex":3889},{"cnt":59646,"rate":8.72,"tgi":-1,"tagValue":"4","tagValueName":"1-2岁","tagValueSortIndex":3890},{"cnt":82024,"rate":11.99,"tgi":-1,"tagValue":"5","tagValueName":"2-3岁","tagValueSortIndex":3891},{"cnt":179896,"rate":26.3,"tgi":-1,"tagValue":"6","tagValueName":"3-6岁","tagValueSortIndex":3892},{"cnt":89068,"rate":13.02,"tgi":-1,"tagValue":"8","tagValueName":"6-9岁","tagValueSortIndex":3893},{"cnt":25788,"rate":3.77,"tgi":-1,"tagValue":"9","tagValueName":"9-12岁","tagValueSortIndex":3894},{"cnt":7434,"rate":1.09,"tgi":-1,"tagValue":"10","tagValueName":"12岁以上","tagValueSortIndex":3895},{"cnt":222919,"rate":32.59,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":3896}],"sr":0,"chartType":"vertical_bar","tips":"单值标签，即，一个人只有一个宝宝，并通过算法预测最有可能的宝宝年龄","valueSortType":0},{"tagType":1017,"tagTypeName":"奶粉年消费金额","tag":"milk_year_amt_region","total":683625,"tagData":[{"cnt":621829,"rate":90.86,"tgi":-1,"tagValue":"-9999","tagValueName":"0元","tagValueSortIndex":4248},{"cnt":9384,"rate":1.37,"tgi":-1,"tagValue":"1","tagValueName":"1-199元","tagValueSortIndex":4249},{"cnt":12048,"rate":1.76,"tgi":-1,"tagValue":"2","tagValueName":"200-499元","tagValueSortIndex":4250},{"cnt":10842,"rate":1.58,"tgi":-1,"tagValue":"3","tagValueName":"500-999元","tagValueSortIndex":4251},{"cnt":11811,"rate":1.73,"tgi":-1,"tagValue":"4","tagValueName":"1000-1999元","tagValueSortIndex":4252},{"cnt":6456,"rate":0.94,"tgi":-1,"tagValue":"5","tagValueName":"2000-2999元","tagValueSortIndex":4253},{"cnt":6359,"rate":0.93,"tgi":-1,"tagValue":"6","tagValueName":"3000-4999元","tagValueSortIndex":4254},{"cnt":4194,"rate":0.61,"tgi":-1,"tagValue":"7","tagValueName":"5000-9999元","tagValueSortIndex":4255},{"cnt":715,"rate":0.1,"tgi":-1,"tagValue":"8","tagValueName":"10000-14999元","tagValueSortIndex":4256},{"cnt":727,"rate":0.12,"tgi":-1,"tagValue":"9","tagValueName":"15000元及以上","tagValueSortIndex":4257}],"sr":0,"chartType":"bar","tips":"婴幼儿牛奶粉，最近365天的消费金额","valueSortType":0}]}
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
    data_dict["奶粉偏好包装"]={}
    for i in pred_gender_data:
        data_dict["奶粉偏好包装"][i["tagValueName"]] = str(i["rate"])+"%"

    #纸尿裤适用性别偏好
    pred_age_level_data =data[1]["tagData"]
    data_dict["纸尿裤适用性别偏好"] = {}
    for i in pred_age_level_data:
        data_dict["纸尿裤适用性别偏好"][i["tagValueName"]] = str(i["rate"]) + "%"

    #辅食产地偏好
    interest_prefer_data = data[2]["tagData"]
    data_dict["辅食产地偏好"] = {}
    for i in interest_prefer_data:
        data_dict["辅食产地偏好"][i["tagValueName"]] = str(i["rate"]) + "%"

    #辅食年消费频次
    common_receive_province_180d_data = data[3]["tagData"]
    data_dict["辅食年消费频次"] = {}
    for i in common_receive_province_180d_data:
        data_dict["辅食年消费频次"][i["tagValueName"]] = str(i["rate"]) + "%"

    #童鞋最近6个月购买金额偏好1
    derive_pay_ord_amt_6m_015_range_data = data[4]["tagData"]
    data_dict["童鞋最近6个月购买金额偏好1"]={}
    for i in derive_pay_ord_amt_6m_015_range_data:
        data_dict["童鞋最近6个月购买金额偏好1"][i["tagValueName"]] = str(i["rate"])+"%"

    #童鞋最近6个月购买金额偏好2
    pred_life_stage_data = data[5]["tagData"]
    data_dict["童鞋最近6个月购买金额偏好2"]={}
    for i in pred_life_stage_data:
        data_dict["童鞋最近6个月购买金额偏好2"][i["tagValueName"]] = str(i["rate"])+"%"

    # 童装最近6个月消费频次偏好
    pred_life_stage_data = data[6]["tagData"]
    data_dict["童装最近6个月消费频次偏好"] = {}
    for i in pred_life_stage_data:
        data_dict["童装最近6个月消费频次偏好"][i["tagValueName"]] = str(i["rate"]) + "%"

    # 宝宝年龄
    pred_life_stage_data = data[7]["tagData"]
    data_dict["宝宝年龄"] = {}
    for i in pred_life_stage_data:
        data_dict["宝宝年龄"][i["tagValueName"]] = str(i["rate"]) + "%"

    # 奶粉年消费金额
    pred_life_stage_data = data[8]["tagData"]
    data_dict["奶粉年消费金额"] = {}
    for i in pred_life_stage_data:
        data_dict["奶粉年消费金额"][i["tagValueName"]] = str(i["rate"]) + "%"


data_merge()

# 插入数据库
insert_item(db.dataMerge, data_dict)
