import xlrd
import numpy as np
import math
import xlwt

class PCR:
    
    def __init__(self,Well,Fluor,Target,Content,Sample,Biological_Set_Name,Cq,Cq_Mean,Cq_Std_Dev):
        self.Well = Well
        self.Fluor = Fluor
        self.Target = Target
        self.Content = Content
        self.Sample = Sample
        self.Biological_Set_Name = Biological_Set_Name
        self.Cq = Cq
        self.Cq_Mean = Cq_Mean
        self.Cq_Std_Dev = Cq_Std_Dev

    @staticmethod
    def isInValid(pcr):
        # if(pcr.Well == ""):
        #     return True
        # if(pcr.Fluor == ""):
        #     return True
        if(pcr.Target == ""):
            return True
        # if(pcr.Content == ""):
        #     return True
        # if(pcr.Sample == ""):
        #     return True
        # if(pcr.Biological_Set_Name == ""):
        #     return True
        # if(pcr.Cq == ""):
        #     return True
        # if(pcr.Cq_Mean == ""):
        #     return True
        # if(pcr.Cq_Std_Dev == ""):
        #     return True
        return False

    def __str__(self):
        return "PCR:Target:" + self.Target + ",Sample:" + self.Sample

    @staticmethod
    def createPCRByList(list):
        return PCR(list[1],list[2],list[3],list[4],list[5],list[6],list[7],list[8],list[9])


filename = input('请输入文件名\n')

# 获取一个工作簿
workbook = xlrd.open_workbook(filename)

# 获取所有的工作表
sheet_names= workbook.sheet_names()

# 获得第一个工作表 也就是PCR生成的表
sheet = workbook.sheet_by_name(sheet_names[0])

pcrs = []
data_dict = {}
i = 1
while(i < sheet.nrows):
    pcr = PCR.createPCRByList(sheet.row_values(i))
    if(not PCR.isInValid(pcr)):
        # pcrs.append(pcr)
        if(pcr.Sample not in data_dict):
            data_dict[pcr.Sample] = {}
            if(pcr.Target not in data_dict[pcr.Sample]):
                data_dict[pcr.Sample][pcr.Target] = []
                data_dict[pcr.Sample][pcr.Target].append(pcr.Cq)
            else:
                data_dict[pcr.Sample][pcr.Target].append(pcr.Cq)
        else:
            if(pcr.Target not in data_dict[pcr.Sample]):
                data_dict[pcr.Sample][pcr.Target] = []
                data_dict[pcr.Sample][pcr.Target].append(pcr.Cq)
            else:
                data_dict[pcr.Sample][pcr.Target].append(pcr.Cq)
    i = i + 1

# 计算GAPAVR delta_ct
for key in list(data_dict.keys()):
    data_dict[key]['GAPAVR'] = np.mean(data_dict[key]['GAPDH'])
    for key2 in list(data_dict[key].keys()):
        if(key2 == 'GAPDH' or key2 == 'GAPAVR' or ('delta' in key2)):
            pass
        else:
            data_dict[key][key2 + '_delta_ct'] =  np.array(data_dict[key][key2]) - np.full((len(data_dict[key][key2])),data_dict[key]['GAPAVR'])
            # data_dict[key][key2 + '_double_delta_ct'] =  data_dict[key][key2 + '_delta_ct'] - data_dict[key]['C-CON_delta_ct']

# 计算double_delta_ct 
for key in list(data_dict.keys()):
    for key2 in list(data_dict[key].keys()):
        if(key2 == 'GAPDH' or key2 == 'GAPAVR' or ('delta' in key2)):
            pass
        else:
            data_dict[key][key2 + '_double_delta_ct'] =  data_dict[key][key2 + '_delta_ct'] - data_dict['C-CON'][key2 + '_delta_ct']


# 结果表格表头的顺序 先写死
resTitle = ['C-CON','C-H','C-N','C-L','A-CON','A-H','A-N','A-L']

# 计算最终结果 2的多少多少次方
for key in resTitle:
    for key2 in list(data_dict[key]):
        if('double' in key2):
            data_dict[key][key2+'_terminal'] = []
            for j in data_dict[key][key2]:
                data_dict[key][key2+'_terminal'].append(math.pow(2,-j))


# for key in resTitle:
#     for key2 in list(data_dict[key]):
#         if('terminal' in  key2):
#             print(key,key2,data_dict[key][key2])


# 根据title得到最后的转置数组
def getResult(title):
    res = []
    for key in resTitle:
        for key2 in list(data_dict[key]):
            if(title + '_double_delta_ct_terminal' == key2):
                res.append(data_dict[key][key2])
    
    return np.transpose(res)



def main():
    title = input('请输入title\n')
    res = getResult(title)
    print(res)
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet('res')

    # 添加表格title
    for i in range(len(resTitle)):
        worksheet.write(0,i,resTitle[i])

    for i in range(res.shape[0]):
        for j in range(res.shape[1]):
            worksheet.write(i+1,j,res[i][j])
    workbook.save('./RES.xls')
    print("文件已经保存到D:/RES.xls")

if __name__ == "__main__":
    main()
            