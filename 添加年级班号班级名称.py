import xlwings as xw

#创建操作对象
app=xw.App(visible=True,add_book=False)
#打开已有的工作簿
wb=app.books.open(r'C:\\Users\\admin\Desktop\\19加入类别码.xlsx')

#打开总表
sht1=wb.sheets['Sheet2']
#将整个表保存在二维数组中，每一行一个数组
Initsheet_total=sht1.range('A2:T37181').value

wb2=app.books.open(r'C:\\Users\\admin\Desktop\\教学班信息2.xlsx')
#打开班级班号表
sht2=wb2.sheets['Sheet2']
#将整个表保存在二维数组中，每一行一个数组
Initsheet_class=sht2.range('A2:D2077').value
list_final = []
index = 0
start = 0
flag=0
for i in range(len(Initsheet_total)):
    for j in range(len(Initsheet_class)):
        if Initsheet_class[j][0] == Initsheet_total[i][8]:
            Initsheet_total[i][7] = Initsheet_class[j][3]
            Initsheet_total[i][6] = Initsheet_class[j][2]
            Initsheet_total[i][5] = Initsheet_class[j][1]
            list_final.append(list(Initsheet_total[i]))

        else:
            continue

    


      
wfinal=app.books.open(r'C:\\Users\\admin\Desktop\\相乘.xlsx')
ws = wfinal.sheets.add()
r = ws.range('A2')
r.value = list_final
