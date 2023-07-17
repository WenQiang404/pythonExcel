import xlwings as xw
from datetime import datetime
from datetime import timedelta

app = xw.App(visible=True,add_book=False)
#打开已有的工作簿
wb=app.books.open(r'C:\\Users\\admin\Desktop\\副本19待处理.xlsx')
#打开sheet表，表不存在保存
sheet = wb.sheets['Sheet3']
#将整个表保存在二维数组中，每一行一个数组
Initsheet=sheet.range('E2:Q37181').value

    
#--------------------------------获取第几周星期几是几号---------------------------
srartMonth = 2
whichWeek = {"1":[2,0,0,0,16,17,18,19],"2":[2,20,21,22,23,24,25,26],"3":[2,27,28,1,2,3,4,5],
             "4":[3,6,7,8,9,10,11,12],"5":[3,13,14,15,16,17,18,19],"6":[3,20,21,22,23,24,25,26],"7":[3,27,28,29,30,31,1,2],
             "8":[4,3,4,5,6,7,8,9],"9":[4,10,11,12,13,14,15,16],"10":[4,17,18,19,20,21,22,23],"11":[4,24,25,26,27,28,29,30],
             "12":[5,1,2,3,4,5,6,7],"13":[5,8,9,10,11,12,13,14],"14":[5,15,16,17,18,19,20,21],"15":[5,22,23,24,25,26,27,28],
             "16":[5,29,30,31,1,2,3,4],"17":[6,5,6,7,8,9,10,11],"18":[6,12,13,14,15,16,17,18],"19":[6,19,20,21,22,23,24,25]}
weekend = {"星期一":"1", "星期二":"2", "星期三":"3", "星期四":"4", "星期五":"5", "星期六":"6", "星期日":"7"}
#根据星期数计算具体日期
def get_current_day(week,week_day): 
    day = whichWeek[week]
    month = day[0]
    count = int(weekend[week_day])
    Finally_Day = day[count]
    if Finally_Day < day[1] :
        month = month + 1
    if Finally_Day <= 9 :
        returnstr = '0' + str(Finally_Day)
    else:
        returnstr = str(Finally_Day)
    count = 0
    return [month,returnstr]
#result = get_current_day("7","星期日")


   
#新建一个数组用于存放最终数据
FinalList = []
for k in range(0,len(Initsheet)):
    intNum_week = int(Initsheet[k][6])
    intNum_week_day = Initsheet[k][7]
    result = get_current_day(str(intNum_week),intNum_week_day)
    Initsheet[k][11] = '20230'+str(result[0]) + result[1]
   
    FinalList.append(list(Initsheet[k]))

#--------------------------------获取第节课是几点---------------------------
sectionTime = {"1":"083000","2":"091500","3":"101000","4":"105500","5":"113500",
           "6":"133000","7":"141500","8":"151000","9":"155500","10":"180000",
           "11":"184500","12":"194000","13":"202500"}

FinalList2 = []

for n in range(0,len(FinalList)):
    section = FinalList[n][9]
    section_index = section.split('-')[0]
    time = sectionTime[str(section_index)]
    FinalList[n][12] = time
    FinalList2.append(list(FinalList[n]))



#将获取到的数据写入excel
wfinal=app.books.open(r'C:\\Users\\admin\Desktop\\19新增日期时间.xlsx')
ws = wfinal.sheets.add()
r = ws.range('E2')
r.value = FinalList2
