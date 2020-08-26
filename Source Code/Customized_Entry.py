from openpyxl import load_workbook
print('MADE BY KAUSHIL NAGRALE\n')
print('*************EASY PWD (https://ams.emahapwd.com/ams/login.do entry) *************\n')
print('*************NOTE ENTER DATA FROM FIRST ROW *************\n')
f= open("javasc.txt","w+")
print('Keep the File inside the same directory as this .py file\n')
shi=input('1. Enter the File name (Excel in file in which data is present)\n')
bo=input('2. Enter the worksheet name\n')
y=input('3. Enter the no of rows\n')
count=1
count=int(count)
y=int(y)
while count < y:
     f.write('addRows();\n')
     count=count+1

cou=y+1
x = 1

re=[]
wb = load_workbook(shi+".xlsx")
ws = wb.get_sheet_by_name(bo)
ha = input('1. Enter the Item Measurement column\n')
column = ws[ha]
re = [column[a].value for a in range(len(column))]

nu=[]
wb = load_workbook(shi+".xlsx")  # Work Book
ws = wb.get_sheet_by_name(bo)
hb = input('2. Enter the No. column\n')
column = ws[hb]  # Column
nu = [column[b].value for b in range(len(column))]

le=[]
wb = load_workbook(shi+".xlsx")  # Work Book
ws = wb.get_sheet_by_name(bo)
hc = input('3. Enter the Length column\n')
column = ws[hc]  # Column
le = [column[c].value for c in range(len(column))]

br=[]
wb = load_workbook(shi+".xlsx")  # Work Book
ws = wb.get_sheet_by_name(bo)
hd = input('4. Enter the Breadth column\n')
column = ws[hd]  # Column
br = [column[d].value for d in range(len(column))]

de=[]
wb = load_workbook(shi+".xlsx")  # Work Book
ws = wb.get_sheet_by_name(bo)
he = input('5. Enter the depth column\n') # Work Sheet
column = ws[he]  # Column
de = [column[e].value for e in range(len(column))]

cou=int(cou)
while x < cou:
  row = ('''document.getElementById('remarks%d').value = '%s';\n''' %(x,re[x-1]))
  if re[x-1]  != None:
       f.write(row)
  row = ('''document.getElementById('num%d').value = '%s';\n''' %(x,nu[x-1]))
  f.write(row)
  row = ('''document.getElementById('len%d').value = '%s';\n''' %(x,le[x-1]))
  f.write(row)
  row = ('''document.getElementById('bre%d').value = '%s';\n'''%(x,br[x-1]))
  f.write(row)
  row = ('''document.getElementById('dep%d').value = '%s';\n''' %(x,de[x-1]))
  f.write(row)
  row = ('''document.getElementById('hm%d').checked  = true;\n''' %x)
  f.write(row)
  x=x+1
count=1
while count < cou:
    f.write('calculateQty(%d);\n'%count)
    count=count+1
print('Sucessful copy the contents of the file of .txt to javascript console')
