import os
import xlrd
import xlwt
import pandas as pd

class Info:

    def __init__(self,Target,Sample,Cq_Mean):
        self.Target = Target
        self.Sample = Sample
        self.Cq_Mean = Cq_Mean

    def __str__(self):
        return "PCR:Target:" + self.Target + ",Sample:" + self.Sample + ",Cq_Mean:" + self.Cq_Mean

filepath = "./files"

def getData(filename):
    # 获取一个工作簿
    workbook = xlrd.open_workbook(filename)
    
    # 获取所有的工作表
    sheet_names = workbook.sheet_names()
    
    # 获得第一个工作表 也就是PCR生成的表
    sheet = workbook.sheet_by_name(sheet_names[0])
    
    i = 1
    data = []
   # data.append(["Target","Sample","Cq Mean","Target","Sample","Cq Mean","Target","Sample","Cq Mean"])
    while(i < sheet.nrows):
        rowData = sheet.row_values(i)
        data.append([rowData[3],rowData[5],rowData[7]])
        i = i + 1
        
    # 处理脏数据
    for d in data:
        if(d[1] == ""):
            d[1] = "Neg Ctrl"
        if(d[2] == ""):
            d[2] = 0.00
            
    for index in range(9):
        data.append(["","",""])
            
            
    return data

def getValidData(data):
    length = 9
    newData = []
    newData.append(["Target","Sample","Cq Mean","Target","Sample","Cq Mean","Target","Sample","Cq Mean"])
    for j in range(9):
        newRow = []
        newRow.extend(data[j])
        newRow.extend(data[j + length])
        newRow.extend(data[j + length * 2])
        newData.append(newRow)
    
    for j in range(27,36):
        newRow = []
        newRow.extend(data[j])
        newRow.extend(data[j + length])
        newRow.extend(data[j + length * 2])
        newData.append(newRow)
        
    for j in range(54,63):
        newRow = []
        newRow.extend(data[j])
        newRow.extend(data[j + length])
        newRow.extend(data[j + length * 2])
        newData.append(newRow)
    
    return newData
    

def getFiles(path):
    return os.listdir(path)

def main():
    root = "./files/"
    files = getFiles(filepath)
    Data = [];
    for filename in files:
        name = filename
        filename = root + filename
        data = getValidData(getData(filename))
        Data.append([name,"","","","","","","",""])
        Data.extend(data)
    
    print(len(Data))
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('res')
    for i in range(len(Data)):
        for j in range(len(Data[1])):
            worksheet.write(i,j,Data[i][j])
    workbook.save('./'+"res"+'.xls')
   

    
    
    
    
    
    
    
        


if __name__ == "__main__":
    main()