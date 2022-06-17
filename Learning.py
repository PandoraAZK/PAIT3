import pandas as pd
import math
import xlsxwriter
import os

# Memakai KNN

# Mengambil data dari excel
def extractExcel():
    destination = "D:/Kampus/Sem4/Pengantar Kecerdasan Buatan/Tugas3/traintest.xlsx"
    learnSource = pd.read_excel(destination,sheet_name=0)
    testSource = pd.read_excel(destination,sheet_name=1)
    learnArr = []
    learnTemp = []
    testArr = []
    testTemp = []
    for i in range(296):
        learnTemp.append(learnSource["x1"].iloc[i])
        learnTemp.append(learnSource["x2"].iloc[i])
        learnTemp.append(learnSource["x3"].iloc[i])
        learnTemp.append(learnSource["y"].iloc[i])
        learnArr.append(learnTemp)
        learnTemp = []
    
    for i in range(10):
        testTemp.append(testSource["x1"].iloc[i])
        testTemp.append(testSource["x2"].iloc[i])
        testTemp.append(testSource["x3"].iloc[i])
        testTemp.append("X")
        testArr.append(testTemp)
        testTemp = []

    return testArr, learnArr

# Membuat training model
def trainModel(learnArray):
    positiveArr = []
    negativeArr = []

    for i in learnArray:
        if i[3] == 1:
            positiveArr.append(i)
        elif i[3] == 0:
            negativeArr.append(i)
    
    return positiveArr, negativeArr

# Menyimpan training model
def trainSave(learnArr):
    trainArr = []
    positiveArr, negativeArr = trainModel(learnArr)

    for i in negativeArr:
        trainArr.append(i)

    for i in positiveArr:
        trainArr.append(i)

    return trainArr

# Testing antara hasil training dengan test subject
def testingProcess(trainArr, testSubject):
    distanceArr = []
    for a in trainArr:
        # Memakai |x1-x2|+|y1-y2|+...+|n1-n2|
        distanceNum = abs(a[0]-testSubject[0])+abs(a[1]-testSubject[1])+abs(a[2]-testSubject[2])
        distanceArr.append(distanceNum)

    return distanceArr

# Testing data
def testing(trainArr, testArr, k):
    for i in testArr:
        result = testingProcess(trainArr, i)
        minimum = min(result)
        basket = []
        while len(basket) != k:
            for j in range(len(result)):
                if len(basket) == k:
                    break
                elif result[j] == minimum:
                    basket.append(j)
            minimum += 1
            
        trainBin = []
        for j in basket:
            trainBin.append(trainArr[j][3])

        countOnes = 0
        for j in range(len(trainBin)):
            if trainBin[j] == 1:
                countOnes += 1
        if countOnes >= math.ceil(k/2):
            i[3] = 1
        else:
            i[3] = 0
    return testArr

# Simpan hasil
def saveRes(testArr, trainArr):
    path = os.path.relpath("Desktop/hasil_learning.xlsx")
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Training")
    row = 0
    worksheet.write(row,0,"Id")
    worksheet.write(row,1,"x1")
    worksheet.write(row,2,"x2")
    worksheet.write(row,3,"x3")
    worksheet.write(row,4,"y")

    for i in trainArr:
        row += 1
        worksheet.write(row,0,row)
        worksheet.write(row,1,i[0])
        worksheet.write(row,2,i[1])
        worksheet.write(row,3,i[2])
        worksheet.write(row,4,i[3])

    worksheet = workbook.add_worksheet("Test")
    row = 0
    worksheet.write(row,0,"Id")
    worksheet.write(row,1,"x1")
    worksheet.write(row,2,"x2")
    worksheet.write(row,3,"x3")
    worksheet.write(row,4,"y")

    for i in testArr:
        row += 1
        worksheet.write(row,0,row)
        worksheet.write(row,1,i[0])
        worksheet.write(row,2,i[1])
        worksheet.write(row,3,i[2])
        worksheet.write(row,4,i[3])

    workbook.close()

k = 9
testArr, learnArr = extractExcel()
trainArr = trainSave(learnArr)
print(trainArr)
testArr = testing(trainArr,testArr,k)
saveRes(testArr,trainArr)
