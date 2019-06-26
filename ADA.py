from excel import OpenExcel
f = OpenExcel('test.xlsx')

#a = f.read() # read all
b = f.read('A')# read 'A' row
c = f.read(1)# f.read('1'), read '1' column
#d = f.read('A5')# read 'A5' position

#print (a)
print (b, type(b))
print (c, type(c))
#print (d)

nCols = len(b) # return number of columns
nRows = len(c) # return number of rows

print(nCols)
print(nRows)