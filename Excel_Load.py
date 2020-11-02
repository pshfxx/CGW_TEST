import os
import win32com.client
import re
import pandas as pd
from tqdm import tqdm


def listit(t):
    return list(map(listit, t)) if isinstance(t, (list, tuple)) else t


def Excel_Load(i_file):
    excel = win32com.client.Dispatch("Excel.Application")
    #excel.Visible = False
    excel_file = excel.Workbooks.Open(i_file)

    sheet_app_direct = excel_file.Sheets('1.CAN APP DIRECT')
    data_app_direct = sheet_app_direct.UsedRange
    data_app_direct = data_app_direct.Value
    data_app_direct = listit(data_app_direct)

    sheet_app_direct_split = excel_file.Sheets('2.CAN APP EXPAN DIRECT SPLIT')
    data_app_direct_split = sheet_app_direct_split.UsedRange
    data_app_direct_split = data_app_direct_split.Value
    data_app_direct_split = listit(data_app_direct_split)

    sheet_app_indirect = excel_file.Sheets('3.CAN APP INDIRECT')
    data_app_indirect = sheet_app_indirect.UsedRange
    data_app_indirect = data_app_indirect.Value
    data_app_indirect = listit(data_app_indirect)

    sheet_diag_ccp = excel_file.Sheets('4.CAN DIAG&CCP')
    data_diag_ccp = sheet_diag_ccp.UsedRange
    data_diag_ccp = data_diag_ccp.Value
    data_diag_ccp = listit(data_diag_ccp)

    sheet_eol_1 = excel_file.Sheets('5.EOL1')
    data_eol_1 = sheet_eol_1.UsedRange
    data_eol_1 = data_eol_1.Value
    data_eol_1 = listit(data_eol_1)

    sheet_eol_4 = excel_file.Sheets('5.EOL4')
    data_eol_4 = sheet_eol_4.UsedRange
    data_eol_4 = data_eol_4.Value
    data_eol_4 = listit(data_eol_4)

    sheet_eol_5 = excel_file.Sheets('5.EOL5')
    data_eol_5 = sheet_eol_5.UsedRange
    data_eol_5 = data_eol_5.Value
    data_eol_5 = listit(data_eol_5)

    sheet_eol_6 = excel_file.Sheets('5.EOL6')
    data_eol_6 = sheet_eol_6.UsedRange
    data_eol_6 = data_eol_6.Value
    data_eol_6 = listit(data_eol_6)

    sheet_eol_8 = excel_file.Sheets('5.EOL8')
    data_eol_8 = sheet_eol_8.UsedRange
    data_eol_8 = data_eol_8.Value
    data_eol_8 = listit(data_eol_8)

    sheet_eol_9 = excel_file.Sheets('5.EOL9')
    data_eol_9 = sheet_eol_9.UsedRange
    data_eol_9 = data_eol_9.Value
    data_eol_9 = listit(data_eol_9)

    excel_file.Close(True)
    excel.Quit()

    print('APP_DIRECT Loading: ')
    df_APP_DIRECT = make_data_frame(data_app_direct)

    return df_APP_DIRECT
"""    print('\nAPP_SPLIT Loading:')
    df_APP_SPLIT = make_data_frame(data_app_direct_split)
    print('\nDIAG_CCP Loading:')
    df_DIAG_CCP = make_data_frame(data_diag_ccp)
    print('\nAPP_INDIRECT Loading:')
    df_APP_INDIRECT = make_data_frame(data_app_indirect)
    print('\nEOL_1 Loading:')
    df_EOL_1 = make_data_frame(data_eol_1)
    print('\nEOL_4 Loading:')
    df_EOL_4 = make_data_frame(data_eol_4)
    print('\nEOL_5 Loading:')
    df_EOL_5 = make_data_frame(data_eol_5)
    print('\nEOL_6 Loading:')
    df_EOL_6 = make_data_frame(data_eol_6)
    print('\nEOL_8 Loading:')
    df_EOL_8 = make_data_frame(data_eol_8)
    print('\nEOL_9 Loading:')
    df_EOL_9 = make_data_frame(data_eol_9)

    print('\nExcel Data Load Complete!')

    return df_APP_DIRECT, df_APP_SPLIT, df_APP_INDIRECT, df_DIAG_CCP, df_EOL_1, df_EOL_4, df_EOL_5, df_EOL_6, df_EOL_8, df_EOL_9
"""

def make_data_frame(in_data):
    fieldBuffer = []
    for index, temp in enumerate(tqdm(in_data)):
        if index == 2:
            tempBuffer = ['SN_Message', 'SN_ID', 'SN_Length', 'SN_CycleTime', 'SN_CanType', 'DN_Message', 'DN_ID',
                          'DN_Length', 'DN_CycleTime', 'DN_CanType', 'ECU_Name', 'Destination', 'WU', 'ChangePoint']
            fieldBuffer.append(pd.Series(tempBuffer))
        elif index >= 3:
            try:
                while True:
                    temp[temp.index(None)] = '-'
                    temp = change_des(temp)

            except ValueError:
                pass
            fieldBuffer.append(pd.Series(temp))
        df = pd.DataFrame.from_records(fieldBuffer)
    return df


def search_file():
    file_path = os.getcwd() + '\\' + 'RDB'
    file_list = os.listdir(file_path)
    filtered_file_list = [file for file in file_list if not(file.startswith("~$")) and file.endswith(".xlsx")]
    file = file_path + '\\' + filtered_file_list[0]

    return file


def change_des(i_data):
    s = ''
    temprow= []
    for i, temp in enumerate(i_data):
        if i == 11:
            s = str(temp)
            s = s.replace("FD-", "")
            s = s.replace('HS-', "")
            s = s.replace("2", "to")
            temprow.append(s)
        else:
            temprow.append(temp)

    return temprow

