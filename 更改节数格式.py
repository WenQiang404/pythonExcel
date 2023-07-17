import xlwings as xw
import datetime

#创建操作对象
app=xw.App(visible=True,add_book=False)
#打开已有的工作簿
wb=app.books.open(r'C:\\Users\\admin\Desktop\\23所需数据.xlsx')

#打开总表
sht1=wb.sheets['Sheet2']
#将整个表保存在二维数组中，每一行一个数组
Initsheet_total=sht1.range('A2:H37181').value

list_final = []

for i in range(len(Initsheet_total)):
# 将字符串转换为datetime对象
    #dt = datetime.datetime.strptime(Initsheet_total[i][14], '%Y-%m-%d %H:%M:%S')
    s = Initsheet_total[i][5].strftime('%m-%d')
    #s = 01-02;09-10;11-12
    if s[0] == '0':
        if s[3] == '0':
            result = '[' + s[1] + '-' + s[4] + ']'
            Initsheet_total[i][5] = result
            list_final.append(list(Initsheet_total[i]))
        else:
            result ='[' +  s[1] + '-' + s[3] + s[4] + ']'
            Initsheet_total[i][5] = result
            list_final.append(list(Initsheet_total[i]))
    else:
        result ='[' + s + ']'
        Initsheet_total[i][5] = result
        list_final.append(list(Initsheet_total[i]))


wfinal=app.books.open(r'C:\\Users\\admin\Desktop\\表23最终版.xlsx')
ws = wfinal.sheets.add()
r = ws.range('A2')
r.value = list_final