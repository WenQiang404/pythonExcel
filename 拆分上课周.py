import xlwings as xw

#创建操作对象
app=xw.App(visible=True,add_book=False)
#打开已有的工作簿
wb=app.books.open(r'C:\\Users\\admin\Desktop\\19待处理.xlsx')

#打开sheet表，表不存在保存
sht=wb.sheets['Sheet1']
#将整个表保存在二维数组中，每一行一个数组
Initsheet=sht.range('E2:O4565').value


#新建一个数组用于存放最终数据
FinalList = []
for i in range(0,len(Initsheet)):
    start = Initsheet[i][6]
    end = Initsheet[i][7]
    startInt = int(start)
    if end != None :
        num = int(end)-int(start)
        for j in range(0,num+1):
            Initsheet[i][6] = startInt + j
            FinalList.append(list(Initsheet[i]))
sheetHead = sht.range('A1:Q1').value


wfinal=app.books.open(r'C:\\Users\\admin\Desktop\\副本19待处理.xlsx')
ws = wfinal.sheets.add()
r = ws.range('E2')
r.value = FinalList

# #将表头加入新的单元格
# r2 = ws.range('A1:Q1')
# r2.value = sheetHead


