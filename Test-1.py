import xlrd
import numpy as np
import matplotlib.pyplot as plt

loc = ("D:\StudioCode\Python\Sample Dataset.xls")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#Master array to be filled
MasterTable = np.full((4,3),0,dtype=object)

#for each cell
for mY in range(0,sheet.ncols,1):
    for mX in range(0,sheet.nrows,1):
        cell = sheet.cell_value(mX, mY)

        if isinstance(cell, str):
            #cleanup
            cell = cell.replace(",", "")
            cell = cell.replace("[", "")
            cell = cell.replace("]", "")
            length = len(cell)
            if cell[0] == " ":
                cell = cell[1:length-1]
            x = ""
            i = len(cell)-1
            if cell.count(".") >= 1:
                if cell[length-2] != ".":
                    while x != ".":
                        x = cell[i]
                        i -= 1
                    cell = cell[0:i+3]
                    length = len(cell)
            cell = cell + " "

            #fill out new list with cleaned data
            cell_List = []
            temp_Str = ""
            for i in cell:
                if i != " ":
                    temp_Str = temp_Str + i
                else:
                    if temp_Str[0].isdigit():
                        cell_List.append(float(temp_Str))
                    else:
                        cell_List.append(temp_Str)
                    temp_Str = ""
            #append cell to master table
            MasterTable[mX,mY] = cell_List
        elif isinstance(cell, float):
            #append cell to master table
            MasterTable[mX,mY] = cell

#Lists
x1 = MasterTable[1,1]
y1 = MasterTable[1,2][:len(MasterTable[1,1])]
x2 = MasterTable[2,1]
y2 = MasterTable[2,2][:len(MasterTable[2,1])]
x3 = MasterTable[3,1]
y3 = MasterTable[3,2][:len(MasterTable[3,1])]

#zero'ing all time values
minimum = int(x1[0]*10)
i=0
while i <= len(x1)-1:
    x1[i] = int(x1[i]*10) - minimum
    i += 1
i=0
minimum = int(x2[0]*10)
while i <= len(x2)-1:
    x2[i] = int(x2[i]*10) - minimum
    i += 1
i=0
minimum = int(x3[0]*10)
while i <= len(x3)-1:
    x3[i] = int(x3[i]*10) - minimum
    i += 1
i=0
while i <= len(x1)-1:
    x1[i] = float(x1[i])/10
    i += 1
i=0
while i <= len(x2)-1:
    x2[i] = float(x2[i])/10
    i += 1
i=0
while i <= len(x3)-1:
    x3[i] = float(x3[i])/10
    i += 1

#Plotting
plt.plot(x1, y1, label = "Use 1")
plt.xlabel('Time(s)')
plt.ylabel('Power(W)')
plt.savefig('Use1.pdf')

plt.clf()
plt.plot(x2, y2, label = "Use 2")
plt.xlabel('Time(s)')
plt.ylabel('Power(W)')
plt.savefig('Use2.pdf')

plt.clf()
plt.plot(x3, y3, label = "Use 3")
plt.xlabel('Time(s)')
plt.ylabel('Power(W)')
plt.savefig('Use3.pdf')

xx1 = x1[:len(x3)]
xx2 = x2[:len(x3)]
xx3 = x3[:len(x3)]
yy1 = y1[:len(x3)]
yy2 = y2[:len(x3)]
yy3 = y3[:len(x3)]

#Scale up
i=0
while i <= len(yy1)-1:
    yy1[i] = int(yy1[i]*10)
    i += 1
i=0
while i <= len(yy2)-1:
    yy2[i] = int(yy2[i]*10)
    i += 1
i=0
while i <= len(yy3)-1:
    yy3[i] = int(yy3[i]*10)
    i += 1

#Integrate
I1 = [0]
I2 = [0]
I3 = [0]
dt = 0.2 # or your time interval
for a in yy1:
    I1.append(I1[-1] + a*dt)
for a in yy2:
    I2.append(I2[-1] + a*dt)
for a in yy3:
    I3.append(I3[-1] + a*dt)

#Scale down
i=0
while i <= len(I1)-1:
    I1[i] = (I1[i]+5)/10
    i += 1
i=0
while i <= len(I2)-1:
    I2[i] = (I2[i]+5)/10
    i += 1
i=0
while i <= len(I3)-1:
    I3[i] = (I3[i]+5)/10
    i += 1
i=0

print(int(I1[len(I1)-1]))
print(int(I2[len(I2)-1]))
print(int(I3[len(I3)-1]))

#Plotting
plt.clf()
plt.plot(xx1, I1[:len(I1)-1], label = "Integral 1")
plt.xlabel('Time(s)')
plt.ylabel('Total Energy(j)')
plt.savefig('Integral1.pdf')

plt.clf()
plt.plot(xx2, I2[:len(I2)-1], label = "Integral 2")
plt.xlabel('Time(s)')
plt.ylabel('Total Energy(j)')
plt.savefig('Integral2.pdf')

plt.clf()
plt.plot(xx3, I3[:len(I3)-1], label = "Integral 3")
plt.xlabel('Time(s)')
plt.ylabel('Total Energy(j)')
plt.savefig('Integral3.pdf')

plt.clf()
plt.plot(xx1, I1[:len(I1)-1], label = "Integral (Use1)")
plt.plot(xx2, I2[:len(I2)-1], label = "Integral (Use2)")
plt.plot(xx3, I3[:len(I3)-1], label = "Integral (Use3)")
plt.xlabel('Time(s)')
plt.ylabel('Total Energy(j)')
plt.legend(loc="upper left")
plt.savefig('Integrals.pdf')

#40s heatup
x1_heatup = xx1[:200]
x2_heatup = xx2[:200]
x3_heatup = xx3[:200]
y1_heatup = I1[:200]
y2_heatup = I2[:200]
y3_heatup = I3[:200]
plt.clf()
plt.plot(x1_heatup, y1_heatup, label = "Use1")
plt.plot(x2_heatup, y2_heatup, label = "Use2")
plt.plot(x3_heatup, y3_heatup, label = "Use3")
plt.xlabel('Time(s)')
plt.ylabel('Total Energy(j)')
plt.legend(loc="upper left")
plt.savefig('Heatup40.pdf')

#140s heatup
x1_heatup = xx1[:740]
x2_heatup = xx2[:740]
x3_heatup = xx3[:740]
y1_heatup = I1[:740]
y2_heatup = I2[:740]
y3_heatup = I3[:740]
plt.clf()
plt.plot(x1_heatup, y1_heatup, label = "Use1")
plt.plot(x2_heatup, y2_heatup, label = "Use2")
plt.plot(x3_heatup, y3_heatup, label = "Use3")
plt.xlabel('Time(s)')
plt.ylabel('Total Energy(j)')
plt.legend(loc="upper left")
plt.savefig('Heatup140.pdf')