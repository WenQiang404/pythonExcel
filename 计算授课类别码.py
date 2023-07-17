import xlwings as xw



#创建操作对象
app=xw.App(visible=True,add_book=False)
#打开已有的工作簿
wb=app.books.open(r'C:\\Users\\admin\\Desktop\\授课类别.xlsx')

#打开sheet表，表不存在保存
sht=wb.sheets['sheet1']
#将整个表保存在二维数组中，每一行一个数组
Initsheet=sht.range('A2:C616').value
#创建类别码对应的字典
dict = {}
for i in range(len(Initsheet)):
    key = Initsheet[i][0]
    param = Initsheet[i][1]
    if '机房' in param:
        value = 4
        dict[key] = value
    elif '场' in param:
        value = 8
        dict[key] = value
    elif '教室' in param:
        value = 5
        dict[key] = value
    else:
        value = 99
        dict[key] = value



wb2=app.books.open(r'C:\\Users\\admin\\Desktop\\19新增日期时间.xlsx')
#打开sheet表，表不存在保存
sheet=wb2.sheets['Sheet2']
Initsheet2=sheet.range('A2:R37181').value

list_class_category = []
for i in range(len(Initsheet2)):
    class_name = Initsheet2[i][4]
    if class_name in dict:
        class_category = dict[class_name]
        Initsheet2[i][17] = class_category
        list_class_category.append(list(Initsheet2[i]))
    else:
        Initsheet2[i][17] = 99
        list_class_category.append(list(Initsheet2[i]))

wfinal=app.books.open(r'C:\\Users\\admin\Desktop\\19加入类别码.xlsx')
ws3 = wfinal.sheets.add()
r = ws3.range('A2')
r.value = list_class_category