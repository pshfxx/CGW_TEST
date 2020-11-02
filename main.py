import os
import re
import pandas as pd
from pandas import ExcelWriter
from tqdm import tqdm
import time
import Excel_Load as EL
import Make_CANType as MC

if __name__ == '__main__':
    RDB_list = EL.search_file() # Excel파일 path 불러오기
    #appDirect, appDirectSplit, appIndirect, diagCcp, eol_1, eol_4, eol_5, eol_6, eol_8, eol_9 = EL.Excel_Load(RDB_list) # 각 시트별로 Dataframe 생성
    appDirect = EL.Excel_Load(RDB_list)
    MC.CAPL_CODE(appDirect)
    print("test")
