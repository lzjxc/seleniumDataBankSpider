import time

import pymongo
import xlsxwriter
from pymongo.errors import DuplicateKeyError

metaData ={"errCode":0,"errMsg":"ok","data":[{"tagType":1311,"tagTypeName":"视频类型偏好","tag":"information_vedio_category_prefer","total":85454,"tagData":[{"cnt":63643,"rate":74.48,"tgi":-1,"tagValue":"33","tagValueName":"影视剧","tagValueSortIndex":4728},{"cnt":33576,"rate":39.29,"tgi":-1,"tagValue":"25","tagValueName":"综艺","tagValueSortIndex":4723},{"cnt":30036,"rate":35.15,"tgi":-1,"tagValue":"1","tagValueName":"社会","tagValueSortIndex":4714},{"cnt":27086,"rate":31.7,"tgi":-1,"tagValue":"13","tagValueName":"娱乐","tagValueSortIndex":4742},{"cnt":22913,"rate":26.81,"tgi":-1,"tagValue":"5","tagValueName":"搞笑","tagValueSortIndex":4716},{"cnt":20856,"rate":24.41,"tgi":-1,"tagValue":"32","tagValueName":"时尚","tagValueSortIndex":4727},{"cnt":16025,"rate":18.75,"tgi":-1,"tagValue":"17","tagValueName":"体育","tagValueSortIndex":4719},{"cnt":12987,"rate":15.2,"tgi":-1,"tagValue":"18","tagValueName":"记录短片","tagValueSortIndex":4745},{"cnt":12238,"rate":14.32,"tgi":-1,"tagValue":"4","tagValueName":"汽车","tagValueSortIndex":4737},{"cnt":11120,"rate":13.01,"tgi":-1,"tagValue":"12","tagValueName":"美食","tagValueSortIndex":4720},{"cnt":11113,"rate":13,"tgi":-1,"tagValue":"9","tagValueName":"音乐","tagValueSortIndex":4721},{"cnt":9406,"rate":11.01,"tgi":-1,"tagValue":"27","tagValueName":"涨姿势","tagValueSortIndex":4746},{"cnt":9339,"rate":10.93,"tgi":-1,"tagValue":"3","tagValueName":"萌娃","tagValueSortIndex":4738},{"cnt":7211,"rate":8.44,"tgi":-1,"tagValue":"8","tagValueName":"科技","tagValueSortIndex":4717},{"cnt":6403,"rate":7.49,"tgi":-1,"tagValue":"34","tagValueName":"育儿","tagValueSortIndex":4729},{"cnt":6407,"rate":7.5,"tgi":-1,"tagValue":"2","tagValueName":"国际","tagValueSortIndex":4732},{"cnt":5264,"rate":6.16,"tgi":-1,"tagValue":"10","tagValueName":"萌宠","tagValueSortIndex":4722},{"cnt":5101,"rate":5.97,"tgi":-1,"tagValue":"6","tagValueName":"游戏","tagValueSortIndex":4735},{"cnt":3900,"rate":4.56,"tgi":-1,"tagValue":"31","tagValueName":"军事","tagValueSortIndex":4726},{"cnt":3748,"rate":4.39,"tgi":-1,"tagValue":"23","tagValueName":"语言类","tagValueSortIndex":4731},{"cnt":3157,"rate":3.69,"tgi":-1,"tagValue":"15","tagValueName":"健康","tagValueSortIndex":4740},{"cnt":3061,"rate":3.58,"tgi":-1,"tagValue":"14","tagValueName":"科学探索","tagValueSortIndex":4741},{"cnt":2274,"rate":2.66,"tgi":-1,"tagValue":"0","tagValueName":"动漫","tagValueSortIndex":4718},{"cnt":1755,"rate":2.05,"tgi":-1,"tagValue":"26","tagValueName":"房产","tagValueSortIndex":4730},{"cnt":1515,"rate":1.77,"tgi":-1,"tagValue":"11","tagValueName":"奇闻","tagValueSortIndex":4744},{"cnt":1416,"rate":1.66,"tgi":-1,"tagValue":"30","tagValueName":"旅游","tagValueSortIndex":4725},{"cnt":760,"rate":0.89,"tgi":-1,"tagValue":"28","tagValueName":"财经","tagValueSortIndex":4724},{"cnt":641,"rate":0.75,"tgi":-1,"tagValue":"20","tagValueName":"国内","tagValueSortIndex":4715},{"cnt":476,"rate":0.56,"tgi":-1,"tagValue":"7","tagValueName":"星座","tagValueSortIndex":4736},{"cnt":397,"rate":0.46,"tgi":-1,"tagValue":"21","tagValueName":"历史","tagValueSortIndex":4733},{"cnt":181,"rate":0.21,"tgi":-1,"tagValue":"29","tagValueName":"教育","tagValueSortIndex":4739},{"cnt":33,"rate":0.04,"tgi":-1,"tagValue":"19","tagValueName":"演讲","tagValueSortIndex":4734},{"cnt":240077,"rate":100,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":4747}],"sr":0,"chartType":"bar","tips":"","valueSortType":2},{"tagType":1314,"tagTypeName":"电影风格偏好","tag":"film_style_prefer","total":85454,"tagData":[{"cnt":104903,"rate":100,"tgi":-1,"tagValue":"24","tagValueName":"院线","tagValueSortIndex":4873},{"cnt":90463,"rate":100,"tgi":-1,"tagValue":"3","tagValueName":"动作","tagValueSortIndex":4861},{"cnt":86217,"rate":100,"tgi":-1,"tagValue":"43","tagValueName":"喜剧","tagValueSortIndex":4898},{"cnt":82192,"rate":96.18,"tgi":-1,"tagValue":"9","tagValueName":"剧情","tagValueSortIndex":4860},{"cnt":58387,"rate":68.33,"tgi":-1,"tagValue":"32","tagValueName":"冒险","tagValueSortIndex":4894},{"cnt":47494,"rate":55.58,"tgi":-1,"tagValue":"31","tagValueName":"言情","tagValueSortIndex":4878},{"cnt":43145,"rate":50.49,"tgi":-1,"tagValue":"33","tagValueName":"奇幻","tagValueSortIndex":4892},{"cnt":41458,"rate":48.52,"tgi":-1,"tagValue":"8","tagValueName":"推理","tagValueSortIndex":4864},{"cnt":38528,"rate":45.09,"tgi":-1,"tagValue":"4","tagValueName":"网络大电影","tagValueSortIndex":4904},{"cnt":37815,"rate":44.25,"tgi":-1,"tagValue":"29","tagValueName":"悬疑","tagValueSortIndex":4880},{"cnt":33941,"rate":39.72,"tgi":-1,"tagValue":"42","tagValueName":"恐怖","tagValueSortIndex":4897},{"cnt":32931,"rate":38.54,"tgi":-1,"tagValue":"44","tagValueName":"动画","tagValueSortIndex":4899},{"cnt":32631,"rate":38.19,"tgi":-1,"tagValue":"35","tagValueName":"科幻","tagValueSortIndex":4890},{"cnt":11596,"rate":13.57,"tgi":-1,"tagValue":"16","tagValueName":"战争","tagValueSortIndex":4871},{"cnt":8258,"rate":9.66,"tgi":-1,"tagValue":"36","tagValueName":"文艺","tagValueSortIndex":4889},{"cnt":7751,"rate":9.07,"tgi":-1,"tagValue":"10","tagValueName":"武侠","tagValueSortIndex":4865},{"cnt":7354,"rate":8.61,"tgi":-1,"tagValue":"22","tagValueName":"传记","tagValueSortIndex":4885},{"cnt":6766,"rate":7.92,"tgi":-1,"tagValue":"23","tagValueName":"历史","tagValueSortIndex":4884},{"cnt":2762,"rate":3.23,"tgi":-1,"tagValue":"14","tagValueName":"运动","tagValueSortIndex":4875},{"cnt":2078,"rate":2.43,"tgi":-1,"tagValue":"21","tagValueName":"纪录片","tagValueSortIndex":4902},{"cnt":1966,"rate":2.3,"tgi":-1,"tagValue":"25","tagValueName":"优酷出品","tagValueSortIndex":4887},{"cnt":1610,"rate":1.88,"tgi":-1,"tagValue":"17","tagValueName":"短片","tagValueSortIndex":4883},{"cnt":996,"rate":1.17,"tgi":-1,"tagValue":"12","tagValueName":"少儿","tagValueSortIndex":4876},{"cnt":256,"rate":0.3,"tgi":-1,"tagValue":"20","tagValueName":"警匪","tagValueSortIndex":4869},{"cnt":242,"rate":0.28,"tgi":-1,"tagValue":"11","tagValueName":"都市","tagValueSortIndex":4862},{"cnt":209,"rate":0.24,"tgi":-1,"tagValue":"40","tagValueName":"文化","tagValueSortIndex":4896},{"cnt":138,"rate":0.16,"tgi":-1,"tagValue":"13","tagValueName":"土豆出品","tagValueSortIndex":4877},{"cnt":129,"rate":0.15,"tgi":-1,"tagValue":"28","tagValueName":"偶像","tagValueSortIndex":4881},{"cnt":125,"rate":0.15,"tgi":-1,"tagValue":"30","tagValueName":"时装","tagValueSortIndex":4879},{"cnt":67,"rate":0.08,"tgi":-1,"tagValue":"5","tagValueName":"搞笑","tagValueSortIndex":4863},{"cnt":69,"rate":0.08,"tgi":-1,"tagValue":"45","tagValueName":"微电影","tagValueSortIndex":4874},{"cnt":61,"rate":0.07,"tgi":-1,"tagValue":"15","tagValueName":"娱乐","tagValueSortIndex":4872},{"cnt":31,"rate":0.04,"tgi":-1,"tagValue":"39","tagValueName":"军事","tagValueSortIndex":4895},{"cnt":5,"rate":0.01,"tgi":-1,"tagValue":"38","tagValueName":"古装","tagValueSortIndex":4888},{"cnt":5,"rate":0.01,"tgi":-1,"tagValue":"19","tagValueName":"青春","tagValueSortIndex":4868},{"cnt":3,"rate":0,"tgi":-1,"tagValue":"37","tagValueName":"穿越","tagValueSortIndex":4903},{"cnt":2,"rate":0,"tgi":-1,"tagValue":"18","tagValueName":"生活","tagValueSortIndex":4893},{"cnt":3,"rate":0,"tgi":-1,"tagValue":"1","tagValueName":"社会","tagValueSortIndex":4900},{"cnt":2,"rate":0,"tgi":-1,"tagValue":"27","tagValueName":"忍者","tagValueSortIndex":4882},{"cnt":4,"rate":0,"tgi":-1,"tagValue":"0","tagValueName":"神话","tagValueSortIndex":4867},{"cnt":6,"rate":0.01,"tgi":-1,"tagValue":"2","tagValueName":"吸血鬼","tagValueSortIndex":4886},{"cnt":4,"rate":0,"tgi":-1,"tagValue":"6","tagValueName":"家庭","tagValueSortIndex":4866},{"cnt":4,"rate":0,"tgi":-1,"tagValue":"7","tagValueName":"竞技","tagValueSortIndex":4901},{"cnt":5,"rate":0.01,"tgi":-1,"tagValue":"26","tagValueName":"网剧","tagValueSortIndex":4870},{"cnt":191165,"rate":100,"tgi":-1,"tagValue":"-9999","tagValueName":"未知","tagValueSortIndex":4905}],"sr":0,"chartType":"bar","tips":"","valueSortType":2}]}
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

    #视频类型偏好
    pred_gender_data = data[0]["tagData"]
    data_dict["视频类型偏好"]={}
    for i in pred_gender_data:
        data_dict["视频类型偏好"][i["tagValueName"]] = str(i["rate"])+"%"

    #电影风格偏好
    pred_age_level_data =data[1]["tagData"]
    data_dict["电影风格偏好"] = {}
    for i in pred_age_level_data:
        data_dict["电影风格偏好"][i["tagValueName"]] = str(i["rate"]) + "%"


data_merge()

# 插入数据库
insert_item(db.dataMerge, data_dict)
