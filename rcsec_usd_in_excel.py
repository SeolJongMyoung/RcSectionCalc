#import xlwings as xw
import openpyxl as op
from openpyxl import load_workbook, Workbook
from string import ascii_uppercase
from sec_back_ver02 import *

#--------------------------------------------------
#    직사각형 단철근보 데이터 입력 (도로교설계기준 2010)
#--------------------------------------------------

class Inputdatafromexcel:
    def __init__(self):
        self.wb = load_workbook("D:/Python/Excel/Calc_As_input.xlsx",read_only=False)
        self.ws_in = self.wb['Sheet1']
        self.wb.close()

    def input_data_translation_list(self):    
        data_list_0 = []
        data_list_1 = []
        data_list_2 = []
        data_list_3 = []
        data_temp_0 = self.ws_in['C4:L4']                   #Openpyxl은 범위 입력시 튜플로 입력됨
        for data_temp_1 in data_temp_0 :                  #튜플형을 리스트형으로 변환 
            for data_temp_2 in data_temp_1 :
                data_list_0.append(data_temp_2.value)
        data_temp_3 = self.ws_in['C8:K8']
        for data_temp_4 in data_temp_3 :                 #튜플형을 리스트형으로 변환 
            for data_temp_5 in data_temp_4 :
                data_list_1.append(data_temp_5.value)
        data_temp_6 = self.ws_in['C11:H11']
        for data_temp_7 in data_temp_6 :                 #튜플형을 리스트형으로 변환 
            for data_temp_8 in data_temp_7 :
                data_list_2.append(data_temp_8.value)


if __name__ == "__main__":
    calc_1 = Sec_back(data_list_0, data_list_1, data_list_2)