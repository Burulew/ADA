from excel import OpenExcel

f = OpenExcel('test.xls')


class Line():
    def __init__(self, weight, layout, deltaX, deltaY, length, startX, startY, corner):
        self.weight = weight
        self.layout = layout
        self.deltaX = deltaX
        self.deltaY = deltaY
        self.length = length
        self.startX = startX
        self.startY = startY
        self.corner = corner


class Cirle():
    def __init__(self, weight, diameter, layout, coordinateX, coordinateY):
        self.weight = weight
        self.layout = layout
        self.diameter = diameter
        self.coordinateX = coordinateX
        self.coordinateY = coordinateY


class Text():
    def __init__(self, layout, coordinateX, coordinateY, height, value):
        self.layout = layout
        self.coordinateX = coordinateX
        self.coordinateY = coordinateY
        self.height = height
        self.value = value


class Point():
    def __init__(self, weight, layout, coordinateX, coordinateY):
        self.weight = weight
        self.layout = layout
        self.coordinateX = coordinateX
        self.coordinateY = coordinateY


class Ellipse():
    def __init__(self, weight, layout, coordinateX, coordinateY, bigVectorX, bigVectorY, smallVectorX, smallVectorY,
                 startCorner, axlesRelation, bigR_axles, smallR_axles):
        self.weight = weight
        self.layout = layout
        self.coordinateX = coordinateX
        self.coordinateY = coordinateY
        self.bigVectorX = bigVectorX
        self.bigVectorY = bigVectorY
        self.smallVectorX = smallVectorX
        self.smallVectorY = smallVectorY
        self.startCorner = startCorner     #I don't know maybe it's not necessary
        self.axlesRelation = axlesRelation #I don't know maybe it's not necessary
        self.bigR_axles = bigR_axles       #I don't know maybe it's not necessary
        self.smallR_axles = smallR_axles   #I don't know maybe it's not necessary


# a = f.read() # read all
b = f.read('A')  # read 'A' row
c = f.read(1)  # f.read('1'), read '1' column
# d = f.read('A5')# read 'A5' position

# print (a)
print(b, "\n", type(b), "\n\n")

print(c, "\n", type(c))
# print (d)

nCols = len(b)  # return number of columns
nRows = len(c)  # return number of rows

print(nCols)
print(nRows)
