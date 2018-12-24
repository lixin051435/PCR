import xlrd
import numpy as np

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



# 获取一个工作簿
workbook = xlrd.open_workbook(u'1.xlsx')

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

for key in list(data_dict.keys()):
    data_dict[key]['GAPAVR'] = np.mean(data_dict[key]['GAPDH'])
    for key2 in list(data_dict[key].keys()):
        if(key2 == 'GAPDH' or key2 == 'GAPAVR'):
            pass
        else:
            data_dict[key][key2 + '_delta_ct'] =  np.array(data_dict[key][key2]) - np.full((len(data_dict[key][key2])),data_dict[key]['GAPAVR'])
            data_dict[key][key2 + '_double_delta_ct'] =  np.array(data_dict[key][key2 + '_delta_ct']) - np.array(data_dict[key]['C-CON_delta_ct'])

for key in data_dict:
    print(data_dict[key])