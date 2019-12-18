import numpy as np
import math
import os
import xlrd
import xlwt
import pandas as pd
import re


class OriginData:

    """
        [00:00.0,Entry,0,Ready,Time]
    """

    def __init__(self,csv_data):
        self.testTime = csv_data[0]
        self.status = csv_data[1]
        self.integer = csv_data[2]
        self.text = csv_data[3]
        self.comment = csv_data[4]

    def __str__(self):
        return "[testTime:%s,text:%s]"%(self.testTime,self.text)

class TransformData:
    """
        filename: 那个csv文件
        textNameCount: 某列出现的次数
        earliestTime: 某列最早出现时间
    """
    def __init__(self,filename,textNameCount,earliestTime):
        self.filename = filename
        self.textNameCount = textNameCount
        self.earliestTime = earliestTime

    def __str__(self):
        return "[filename:%s,textNameCount:%s,earliestTime:%s]"%(self.filename,self.textNameCount,self.earliestTime)

def getData(filename,textname="On1A2"):
    csv_data = pd.read_csv(filename).iloc[6:,1:]
    originData = None
    originDataList = []
    for i in range(0,len(csv_data)):
        originData = OriginData(csv_data.iloc[i])
        if(originData.text == textname):
            originDataList.append(originData)

    result = None
    transformdata = TransformData(filename,'','')
    pattern = re.compile(r'\d+-\d+')
    filename_new = " "+pattern.findall(filename)[0]
    try:
        transformdata =  TransformData(filename,len(originDataList),originDataList[0].testTime)
        result = {"filename":filename_new,"count":transformdata.textNameCount,"earliestTime":transformdata.earliestTime}
    except Exception:
        result = {"filename":filename_new,"count":'0',"earliestTime":'5:00.00'}
    
    return result

def createDataCSV(dataList):
    pass


def main(text="On1A2"):
    filelist = os.listdir()
    csvlist = []
    for filename in filelist:
        if("csv" in filename):
            csvlist.append(filename)

    datalist = []
    for i in range(0,len(csvlist)):
        datalist.append(getData(csvlist[i],text))

    names = ["filename","count","earliestTime"]
    df = pd.DataFrame(datalist,columns=names)
    df.to_csv("result.csv",index=False)
    print("已经生成result.csv")

if __name__ == "__main__":
    main()