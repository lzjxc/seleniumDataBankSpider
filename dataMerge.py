import time

import pymongo
import xlsxwriter
from pymongo.errors import DuplicateKeyError

metaData= \
    {"errCode": 0, "errMsg": "ok", "data": [
        {"tagType": 1001, "tagTypeName": "预测性别", "tag": "pred_gender", "total": 2229145, "tagData": [
            {"cnt": 824699, "rate": 36.97, "tgi": -1, "tagValue": "1", "tagValueName": "女",
             "tagValueSortIndex": 3867},
            {"cnt": 1263079, "rate": 56.63, "tgi": -1, "tagValue": "2", "tagValueName": "男",
             "tagValueSortIndex": 3868},
            {"cnt": 142765, "rate": 6.4, "tgi": -1, "tagValue": "-9999", "tagValueName": "未知",
             "tagValueSortIndex": 3869}], "sr": 0, "chartType": "pie", "tips": "", "valueSortType": 0},
        {"tagType": 1000, "tagTypeName": "预测年龄", "tag": "pred_age_level", "total": 2229145, "tagData": [
            {"cnt": 338573, "rate": 15.58, "tgi": -1, "tagValue": "10", "tagValueName": "18-24岁",
             "tagValueSortIndex": 3897},
            {"cnt": 236993, "rate": 10.91, "tgi": -1, "tagValue": "11", "tagValueName": "25-29岁",
             "tagValueSortIndex": 3898},
            {"cnt": 329628, "rate": 15.17, "tgi": -1, "tagValue": "12", "tagValueName": "30-34岁",
             "tagValueSortIndex": 3899},
            {"cnt": 252021, "rate": 11.6, "tgi": -1, "tagValue": "13", "tagValueName": "35-39岁",
             "tagValueSortIndex": 3900},
            {"cnt": 275837, "rate": 12.69, "tgi": -1, "tagValue": "14", "tagValueName": "40-44岁",
             "tagValueSortIndex": 3901},
            {"cnt": 259641, "rate": 11.95, "tgi": -1, "tagValue": "15", "tagValueName": "45-49岁",
             "tagValueSortIndex": 3902},
            {"cnt": 131186, "rate": 6.04, "tgi": -1, "tagValue": "16", "tagValueName": "50-54岁",
             "tagValueSortIndex": 3903},
            {"cnt": 78870, "rate": 3.63, "tgi": -1, "tagValue": "17", "tagValueName": "55-59岁",
             "tagValueSortIndex": 3904},
            {"cnt": 129129, "rate": 5.94, "tgi": -1, "tagValue": "18", "tagValueName": ">=60岁",
             "tagValueSortIndex": 3905},
            {"cnt": 140949, "rate": 6.49, "tgi": -1, "tagValue": "-9999", "tagValueName": "未知",
             "tagValueSortIndex": 3906}], "sr": 0, "chartType": "pie", "tips": "", "valueSortType": 0},
        {"tagType": 1008, "tagTypeName": "兴趣偏好", "tag": "interest_prefer", "total": 2229145, "tagData": [
            {"cnt": 588279, "rate": 26.39, "tgi": -1, "tagValue": "29", "tagValueName": "数码达人",
             "tagValueSortIndex": 4701},
            {"cnt": 554378, "rate": 24.87, "tgi": -1, "tagValue": "21", "tagValueName": "烹饪达人",
             "tagValueSortIndex": 4694},
            {"cnt": 538069, "rate": 24.14, "tgi": -1, "tagValue": "4", "tagValueName": "吃货",
             "tagValueSortIndex": 4677},
            {"cnt": 532588, "rate": 23.89, "tgi": -1, "tagValue": "19", "tagValueName": "买鞋控",
             "tagValueSortIndex": 4692},
            {"cnt": 453919, "rate": 20.36, "tgi": -1, "tagValue": "39", "tagValueName": "运动一族",
             "tagValueSortIndex": 4711},
            {"cnt": 422313, "rate": 18.95, "tgi": -1, "tagValue": "25", "tagValueName": "时尚靓妹",
             "tagValueSortIndex": 4697},
            {"cnt": 417227, "rate": 18.72, "tgi": -1, "tagValue": "13", "tagValueName": "家庭主妇",
             "tagValueSortIndex": 4686},
            {"cnt": 405449, "rate": 18.19, "tgi": -1, "tagValue": "8", "tagValueName": "高富帅",
             "tagValueSortIndex": 4681},
            {"cnt": 388176, "rate": 17.41, "tgi": -1, "tagValue": "20", "tagValueName": "美丽教主",
             "tagValueSortIndex": 4693},
            {"cnt": 379421, "rate": 17.02, "tgi": -1, "tagValue": "37", "tagValueName": "有型潮男",
             "tagValueSortIndex": 4709},
            {"cnt": 323773, "rate": 14.52, "tgi": -1, "tagValue": "40", "tagValueName": "职场办公",
             "tagValueSortIndex": 4712},
            {"cnt": 274979, "rate": 12.34, "tgi": -1, "tagValue": "3", "tagValueName": "白富美",
             "tagValueSortIndex": 4676},
            {"cnt": 256016, "rate": 11.48, "tgi": -1, "tagValue": "1", "tagValueName": "爱包人",
             "tagValueSortIndex": 4674},
            {"cnt": 241510, "rate": 10.83, "tgi": -1, "tagValue": "30", "tagValueName": "速食客",
             "tagValueSortIndex": 4702},
            {"cnt": 226150, "rate": 10.15, "tgi": -1, "tagValue": "35", "tagValueName": "养生专家",
             "tagValueSortIndex": 4707},
            {"cnt": 211953, "rate": 9.51, "tgi": -1, "tagValue": "10", "tagValueName": "户外一族",
             "tagValueSortIndex": 4683},
            {"cnt": 202727, "rate": 9.09, "tgi": -1, "tagValue": "27", "tagValueName": "收纳达人",
             "tagValueSortIndex": 4699},
            {"cnt": 171177, "rate": 7.68, "tgi": -1, "tagValue": "34", "tagValueName": "学霸",
             "tagValueSortIndex": 4706},
            {"cnt": 171245, "rate": 7.68, "tgi": -1, "tagValue": "38", "tagValueName": "阅读者",
             "tagValueSortIndex": 4710},
            {"cnt": 168883, "rate": 7.58, "tgi": -1, "tagValue": "2", "tagValueName": "爱听音乐",
             "tagValueSortIndex": 4675},
            {"cnt": 147071, "rate": 6.6, "tgi": -1, "tagValue": "11", "tagValueName": "花卉一族",
             "tagValueSortIndex": 4684},
            {"cnt": 140135, "rate": 6.29, "tgi": -1, "tagValue": "33", "tagValueName": "休闲大咖",
             "tagValueSortIndex": 4705},
            {"cnt": 134731, "rate": 6.04, "tgi": -1, "tagValue": "14", "tagValueName": "健美一族",
             "tagValueSortIndex": 4687},
            {"cnt": 125839, "rate": 5.65, "tgi": -1, "tagValue": "7", "tagValueName": "二手买家",
             "tagValueSortIndex": 4680},
            {"cnt": 101866, "rate": 4.57, "tgi": -1, "tagValue": "41", "tagValueName": "追风骑士",
             "tagValueSortIndex": 4713},
            {"cnt": 88334, "rate": 3.96, "tgi": -1, "tagValue": "36", "tagValueName": "游戏人生",
             "tagValueSortIndex": 4708},
            {"cnt": 71857, "rate": 3.22, "tgi": -1, "tagValue": "6", "tagValueName": "动漫迷",
             "tagValueSortIndex": 4679},
            {"cnt": 55676, "rate": 2.5, "tgi": -1, "tagValue": "15", "tagValueName": "酒品人生",
             "tagValueSortIndex": 4688},
            {"cnt": 48139, "rate": 2.16, "tgi": -1, "tagValue": "26", "tagValueName": "收藏家",
             "tagValueSortIndex": 4698},
            {"cnt": 46830, "rate": 2.1, "tgi": -1, "tagValue": "18", "tagValueName": "旅行者",
             "tagValueSortIndex": 4691},
            {"cnt": 41566, "rate": 1.86, "tgi": -1, "tagValue": "31", "tagValueName": "网络一族",
             "tagValueSortIndex": 4703},
            {"cnt": 37322, "rate": 1.67, "tgi": -1, "tagValue": "24", "tagValueName": "摄影一族",
             "tagValueSortIndex": 4696},
            {"cnt": 34465, "rate": 1.55, "tgi": -1, "tagValue": "12", "tagValueName": "绘画家",
             "tagValueSortIndex": 4685},
            {"cnt": 26952, "rate": 1.21, "tgi": -1, "tagValue": "32", "tagValueName": "舞林人士",
             "tagValueSortIndex": 4704},
            {"cnt": 25888, "rate": 1.16, "tgi": -1, "tagValue": "16", "tagValueName": "乐器迷",
             "tagValueSortIndex": 4689},
            {"cnt": 7087, "rate": 0.32, "tgi": -1, "tagValue": "23", "tagValueName": "商家会",
             "tagValueSortIndex": 4695},
            {"cnt": 3454, "rate": 0.15, "tgi": -1, "tagValue": "5", "tagValueName": "电影派", "tagValueSortIndex": 4678},
            {"cnt": 1688, "rate": 0.08, "tgi": -1, "tagValue": "9", "tagValueName": "果粉", "tagValueSortIndex": 4682},
            {"cnt": 1887, "rate": 0.08, "tgi": -1, "tagValue": "28", "tagValueName": "书法家",
             "tagValueSortIndex": 4700},
            {"cnt": 1076673, "rate": 48.3, "tgi": -1, "tagValue": "-9999", "tagValueName": "未知",
             "tagValueSortIndex": 4673}], "sr": 0, "chartType": "bar",
         "tips": "在类目有强互动行为，如\n白富美：个护美妆，健身，贵重饰品女性\n高富帅：高档生活用品，汽车用品，男士护理，男装，健身男性\n美丽教主：个护，健身塑形，护肤美体，流行饰品女性\n时尚靓妹：精品女装，精品女鞋，服饰配件女性\n有型潮男：高端饰品，个护，男鞋男装男性",
         "valueSortType": 2},
        {"tagType": 1009, "tagTypeName": "省份", "tag": "common_receive_province_180d", "total": 2229145, "tagData": [
            {"cnt": 619462, "rate": 27.8, "tgi": -1, "tagValue": "-9999", "tagValueName": "未知",
             "tagValueSortIndex": 5055},
            {"cnt": 58397, "rate": 2.62, "tgi": -1, "tagValue": "1", "tagValueName": "安徽", "tagValueSortIndex": 5056},
            {"cnt": 10, "rate": 0, "tgi": -1, "tagValue": "2", "tagValueName": "澳门", "tagValueSortIndex": 5057},
            {"cnt": 49360, "rate": 2.21, "tgi": -1, "tagValue": "3", "tagValueName": "北京", "tagValueSortIndex": 5058},
            {"cnt": 48976, "rate": 2.2, "tgi": -1, "tagValue": "4", "tagValueName": "福建", "tagValueSortIndex": 5059},
            {"cnt": 21106, "rate": 0.95, "tgi": -1, "tagValue": "5", "tagValueName": "甘肃", "tagValueSortIndex": 5060},
            {"cnt": 251194, "rate": 11.27, "tgi": -1, "tagValue": "6", "tagValueName": "广东",
             "tagValueSortIndex": 5061},
            {"cnt": 62336, "rate": 2.8, "tgi": -1, "tagValue": "7", "tagValueName": "广西", "tagValueSortIndex": 5062},
            {"cnt": 50749, "rate": 2.28, "tgi": -1, "tagValue": "8", "tagValueName": "贵州", "tagValueSortIndex": 5063},
            {"cnt": 10439, "rate": 0.47, "tgi": -1, "tagValue": "9", "tagValueName": "海南", "tagValueSortIndex": 5064},
            {"cnt": 51254, "rate": 2.3, "tgi": -1, "tagValue": "10", "tagValueName": "河北", "tagValueSortIndex": 5065},
            {"cnt": 86113, "rate": 3.86, "tgi": -1, "tagValue": "11", "tagValueName": "河南",
             "tagValueSortIndex": 5066},
            {"cnt": 32997, "rate": 1.48, "tgi": -1, "tagValue": "12", "tagValueName": "黑龙江",
             "tagValueSortIndex": 5067},
            {"cnt": 60376, "rate": 2.71, "tgi": -1, "tagValue": "13", "tagValueName": "湖北",
             "tagValueSortIndex": 5068},
            {"cnt": 57805, "rate": 2.59, "tgi": -1, "tagValue": "14", "tagValueName": "湖南",
             "tagValueSortIndex": 5069},
            {"cnt": 18228, "rate": 0.82, "tgi": -1, "tagValue": "15", "tagValueName": "吉林",
             "tagValueSortIndex": 5070},
            {"cnt": 103122, "rate": 4.63, "tgi": -1, "tagValue": "16", "tagValueName": "江苏",
             "tagValueSortIndex": 5071},
            {"cnt": 60760, "rate": 2.73, "tgi": -1, "tagValue": "17", "tagValueName": "江西",
             "tagValueSortIndex": 5072},
            {"cnt": 32155, "rate": 1.44, "tgi": -1, "tagValue": "18", "tagValueName": "辽宁",
             "tagValueSortIndex": 5073},
            {"cnt": 17583, "rate": 0.79, "tgi": -1, "tagValue": "19", "tagValueName": "内蒙古",
             "tagValueSortIndex": 5074},
            {"cnt": 8162, "rate": 0.37, "tgi": -1, "tagValue": "20", "tagValueName": "宁夏", "tagValueSortIndex": 5075},
            {"cnt": 4373, "rate": 0.2, "tgi": -1, "tagValue": "21", "tagValueName": "青海", "tagValueSortIndex": 5076},
            {"cnt": 83064, "rate": 3.73, "tgi": -1, "tagValue": "22", "tagValueName": "山东",
             "tagValueSortIndex": 5077},
            {"cnt": 28882, "rate": 1.3, "tgi": -1, "tagValue": "23", "tagValueName": "山西", "tagValueSortIndex": 5078},
            {"cnt": 45186, "rate": 2.03, "tgi": -1, "tagValue": "24", "tagValueName": "陕西",
             "tagValueSortIndex": 5079},
            {"cnt": 50927, "rate": 2.29, "tgi": -1, "tagValue": "25", "tagValueName": "上海",
             "tagValueSortIndex": 5080},
            {"cnt": 86217, "rate": 3.87, "tgi": -1, "tagValue": "26", "tagValueName": "四川",
             "tagValueSortIndex": 5081},
            {"cnt": 28, "rate": 0, "tgi": -1, "tagValue": "27", "tagValueName": "台湾", "tagValueSortIndex": 5082},
            {"cnt": 13837, "rate": 0.62, "tgi": -1, "tagValue": "28", "tagValueName": "天津",
             "tagValueSortIndex": 5083},
            {"cnt": 1331, "rate": 0.06, "tgi": -1, "tagValue": "29", "tagValueName": "西藏", "tagValueSortIndex": 5084},
            {"cnt": 109, "rate": 0, "tgi": -1, "tagValue": "30", "tagValueName": "香港", "tagValueSortIndex": 5085},
            {"cnt": 12705, "rate": 0.57, "tgi": -1, "tagValue": "31", "tagValueName": "新疆",
             "tagValueSortIndex": 5086},
            {"cnt": 44293, "rate": 1.99, "tgi": -1, "tagValue": "32", "tagValueName": "云南",
             "tagValueSortIndex": 5087},
            {"cnt": 117315, "rate": 5.26, "tgi": -1, "tagValue": "33", "tagValueName": "浙江",
             "tagValueSortIndex": 5088},
            {"cnt": 39684, "rate": 1.76, "tgi": -1, "tagValue": "34", "tagValueName": "重庆",
             "tagValueSortIndex": 5089}], "sr": 0, "chartType": "china", "tips": "通过消费者的常住地计算出来的", "valueSortType": 0},
        {"tagType": 1010, "tagTypeName": "月均消费金额", "tag": "derive_pay_ord_amt_6m_015_range", "total": 2229145,
         "tagData": [{"cnt": 1002551, "rate": 44.97, "tgi": -1, "tagValue": "1", "tagValueName": "0-499元",
                      "tagValueSortIndex": 4979},
                     {"cnt": 215686, "rate": 9.68, "tgi": -1, "tagValue": "2", "tagValueName": "500-999元",
                      "tagValueSortIndex": 4980},
                     {"cnt": 96168, "rate": 4.31, "tgi": -1, "tagValue": "3", "tagValueName": "1000-1499元",
                      "tagValueSortIndex": 4981},
                     {"cnt": 52023, "rate": 2.33, "tgi": -1, "tagValue": "4", "tagValueName": "1500-1999元",
                      "tagValueSortIndex": 4982},
                     {"cnt": 52256, "rate": 2.34, "tgi": -1, "tagValue": "5", "tagValueName": "2000-2999元",
                      "tagValueSortIndex": 4983},
                     {"cnt": 46709, "rate": 2.1, "tgi": -1, "tagValue": "6", "tagValueName": "3000-5999元",
                      "tagValueSortIndex": 4984},
                     {"cnt": 14037, "rate": 0.63, "tgi": -1, "tagValue": "7", "tagValueName": "6000-10000元",
                      "tagValueSortIndex": 4985},
                     {"cnt": 10614, "rate": 0.48, "tgi": -1, "tagValue": "8", "tagValueName": "10000元以上",
                      "tagValueSortIndex": 4986},
                     {"cnt": 739164, "rate": 33.16, "tgi": -1, "tagValue": "-9999", "tagValueName": "未知",
                      "tagValueSortIndex": 4987}], "sr": 0, "chartType": "bar", "tips": "最近180天，消费者在淘宝天猫上的月均消费金额",
         "valueSortType": 0}, {"tagType": 1003, "tagTypeName": "人生阶段", "tag": "pred_life_stage", "total": 2229145,
                               "tagData": [
                                   {"cnt": 367286, "rate": 16.48, "tgi": -1, "tagValue": "1", "tagValueName": "单身",
                                    "tagValueSortIndex": 3855},
                                   {"cnt": 57622, "rate": 2.58, "tgi": -1, "tagValue": "2", "tagValueName": "恋爱期",
                                    "tagValueSortIndex": 3856},
                                   {"cnt": 66906, "rate": 3, "tgi": -1, "tagValue": "3", "tagValueName": "准备结婚期",
                                    "tagValueSortIndex": 3857},
                                   {"cnt": 393783, "rate": 17.66, "tgi": -1, "tagValue": "4", "tagValueName": "已婚未育",
                                    "tagValueSortIndex": 3858},
                                   {"cnt": 19904, "rate": 0.89, "tgi": -1, "tagValue": "5", "tagValueName": "育婴期",
                                    "tagValueSortIndex": 3859},
                                   {"cnt": 357424, "rate": 16.03, "tgi": -1, "tagValue": "6", "tagValueName": "已婚已育",
                                    "tagValueSortIndex": 3860},
                                   {"cnt": 14793, "rate": 0.66, "tgi": -1, "tagValue": "7", "tagValueName": "孝敬期",
                                    "tagValueSortIndex": 3861},
                                   {"cnt": 951563, "rate": 42.7, "tgi": -1, "tagValue": "-9999", "tagValueName": "未知",
                                    "tagValueSortIndex": 3862}], "sr": 0, "chartType": "vertical_bar", "tips": "",
                               "valueSortType": 0}]}
client = pymongo.MongoClient(host='192.168.0.47', port=27017)
db= client.DataBank

collection = db.dataMerge
creatTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
data_dict = {}
data_dict["_id"]="数据融合xunfei2" + str(creatTime)

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

    #预测性别
    pred_gender_data = data[0]["tagData"]
    data_dict["预测性别"]={}
    for i in pred_gender_data:
        data_dict["预测性别"][i["tagValueName"]] = str(i["rate"])+"%"

    #预测年龄
    pred_age_level_data =data[1]["tagData"]
    data_dict["预测年龄"] = {}
    for i in pred_age_level_data:
        data_dict["预测年龄"][i["tagValueName"]] = str(i["rate"]) + "%"

    #兴趣偏好
    interest_prefer_data = data[2]["tagData"]
    data_dict["兴趣偏好"] = {}
    for i in interest_prefer_data:
        data_dict["兴趣偏好"][i["tagValueName"]] = str(i["rate"]) + "%"

    #省份
    common_receive_province_180d_data = data[3]["tagData"]
    data_dict["省份"] = {}
    for i in common_receive_province_180d_data:
        data_dict["省份"][i["tagValueName"]] = str(i["rate"]) + "%"

    #月均消费金额
    derive_pay_ord_amt_6m_015_range_data = data[4]["tagData"]
    data_dict["月均消费金额"]={}
    for i in derive_pay_ord_amt_6m_015_range_data:
        data_dict["月均消费金额"][i["tagValueName"]] = str(i["rate"])+"%"

    #人生阶段
    pred_life_stage_data = data[5]["tagData"]
    data_dict["人生阶段"]={}
    for i in pred_life_stage_data:
        data_dict["人生阶段"][i["tagValueName"]] = str(i["rate"])+"%"

data_merge()

# 插入数据库
insert_item(db.dataMerge, data_dict)
