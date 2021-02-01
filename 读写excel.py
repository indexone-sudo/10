
import xlrd  #引入模块
#12313
#打开文件，获取excel文件的workbook（工作簿）对象

'''对workbook对象进行操作'''

#获取所有sheet的名字
names=workbook.sheet_names()
print(names) #  输出所有的表名，以列表的形式

#通过sheet索引获得sheet对象,fir-sheet
worksheet=workbook.sheet_by_index(0)
sheet0=worksheet
print(worksheet)  #<xlrd.sheet.Sheet object at 0x000001B98D99CFD0>
'''
#通过sheet名获得sheet对象
worksheet=workbook.sheet_by_name("123")
print(worksheet) #<xlrd.sheet.Sheet object at 0x000001B98D99CFD0>
'''
'''对sheet对象进行操作'''
name=worksheet.name  #获取表的姓名
print(name) #

nrows=worksheet.nrows  #获取该表总行数
print(nrows)  #32
ncols=worksheet.ncols  #获取该表总列数
print(ncols) #13


col_data=worksheet.col_values(1)  #获取第一列的内容
print(col_data)
print(col_data[1])#b1

#通过坐标读取表格中的数据
cell_value1=sheet0.cell_value(1,1)
print(cell_value1)
#123123123213213



# 导入xlwt模块
import xlwt

#创建一个Workbook对象，相当于创建了一个Excel文件
book=xlwt.Workbook(encoding="utf-8",style_compression=0)

'''
Workbook类初始化时有encoding和style_compression参数
encoding:设置字符编码，一般要这样设置：w = Workbook(encoding='utf-8')，就可以在excel中输出中文了。默认是ascii。
style_compression:表示是否压缩，不常用。
'''

# 创建一个sheet对象
sheet = book.add_sheet('test01', cell_overwrite_ok=True############)#
# 其中的test是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False


# 向表test中添加数据
sheet.write(0, 0, '0xff')  # 其中的'0-行, 0-列'指定表中的单元，''

# 最后，将以上操作保存到指定的Excel文件中
book.save('banch\\test145.xlsx')
