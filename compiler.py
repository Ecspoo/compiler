import os
import xlrd
import xlwt
import time
import random
import datetime
people1=['a', 'b', 'c', 'd', 'e']
fileName=os.getcwd()+'\\compilee.xls'
resultFileName=os.getcwd()+'\\result.xls'
wb=xlrd.open_workbook(fileName)
sh=wb.sheet_by_index(0)
nr=sh.nrows
nc=sh.ncols
mat=[]
outputMat=[]
for i in range(nr):
    mat=mat+[[""]*nc]
for i in range(nr):
    for j in range(nc):
       mat[i][j]=sh.cell(i, j).value
for i in range(2, nr):
    outputHead=[mat[i][0], mat[i][4], mat[i][2]]
    num=int(mat[i][3])
    numt=num
    numl=[]
    while numt>0:
        rant=random.randint(0,100)
        if rant<80:
            b=1
        elif rant<90:
            b=2
        else:
            b=3
        numl=numl+[b]
        numt=numt-b
    if numt==-1:
        numl[-1]=numl[-1]-1
    if numt==-2:
        numl[-1]=numl[-1]-2
    if num==230:
        numl=[230]
    timeList=mat[i][4].split('.')
    currentTime=datetime.datetime(int(timeList[0]), int(timeList[1]), int(timeList[2]))
    outputLines=[]
    for numNow in numl:
        deltaDay=random.randint(1, 5)
        currentTime=currentTime+datetime.timedelta(days=deltaDay)
        cMonth=currentTime.month
        if cMonth<6:
            xPeople=random.choice(people1)
        elif cMonth<9:
            xPeople=random.choice(people2)
        else:
            xPeople=random.choice(people3)
        outputLines=outputLines+[[numNow, currentTime.strftime("%Y.%m.%d"), xPeople]]
    outputMat=outputMat+[[outputHead, outputLines]]
wbk=xlwt.Workbook()
sheet1=wbk.add_sheet('sheet1')
linePointer=0
for i in outputMat:
    print(i)
    print('')
    lineNum=len(i[1])
    sheet1.write_merge(linePointer, linePointer+lineNum-1, 0, 0, i[0][0])
    sheet1.write_merge(linePointer, linePointer+lineNum-1, 1, 1, i[0][1])
    sheet1.write_merge(linePointer, linePointer+lineNum-1, 2, 2, i[0][2])
    for j in i[1]:
        sheet1.write(linePointer, 3, j[0])
        sheet1.write(linePointer, 4, j[1])
        sheet1.write(linePointer, 5, j[2])
        linePointer=linePointer+1
    linePointer=linePointer+1
wbk.save(resultFileName)
print('DONE')
time.sleep(1.7)
