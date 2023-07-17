import xlwings as xw
import re

#创建操作对象
app=xw.App(visible=True,add_book=False)
#打开已有的工作簿
wb=app.books.open(r'C:\\Users\\admin\Desktop\\教学班信息.xlsx')

#打开sheet表，表不存在保存
sht=wb.sheets['sheet1']
#将整个表保存在二维数组中，每一行一个数组
Initsheet=sht.range('A2:E1882').value
Finally_list = []
for i in range(len(Initsheet)):
    handle_str = Initsheet[i][0]
    if handle_str == '无':
        Initsheet[i][2] = '无'
        Initsheet[i][3] = '无'
        Initsheet[i][4] = '无'
        Finally_list.append(list(Initsheet[i]))
    elif handle_str == '2020' or handle_str == '2021' or handle_str == '2022' or handle_str == '2023':
        Initsheet[i][2] = handle_str
        Initsheet[i][3] = '无'
        Initsheet[i][4] = '无'
        Finally_list.append(list(Initsheet[i]))
    else:
        if ';' in handle_str:
            part_str_list = handle_str.split(';')
            for j in range(len(part_str_list)):
                Initsheet[i][4] = part_str_list[j][:-4]
                Initsheet[i][3] = part_str_list[j][-4:]
                Initsheet[i][2] = '20' + part_str_list[j][-4:][:2]
                Finally_list.append(list(Initsheet[i]))
        else:
            Initsheet[i][4] = handle_str[:-4]
            Initsheet[i][3] = handle_str[-4:]
            Initsheet[i][2] = '20' + handle_str[-4:][:2]
            Finally_list.append(list(Initsheet[i]))


wfinal=app.books.open(r'C:\\Users\\admin\Desktop\\教学班信息2.xlsx')
ws = wfinal.sheets.add()
r = ws.range('A2')
r.value = Finally_list