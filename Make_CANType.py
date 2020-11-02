import os
import win32com.client
import re
import pandas as pd
from pandas import DataFrame
from tqdm import tqdm
from os import unlink
import numpy as np


def CAPL_CODE(i_data):
    #unlink("test.txt")
    f = open("test.txt", 'w', encoding="UTF8")
    print("\n========== Message Variable Generation Start! ========== ")
    tx_CCANFD_Var, tx_ECANFD_Var, tx_GCANFD_Var, tx_PCANFD_Var, tx_HSBCAN_Var, tx_HSMCAN_Var, tx_HSICAN_Var, tx_HSDCAN_Var = Gen_Variables(i_data)
    #파일 쓰기 시작
    f.write(Tx_CAN_include)
    f.write(Tx_CAN_on_start)
    f.write(Tx_CAN_variables_Start)
    f.write(tx_CCANFD_Var)
    f.write(Tx_CAN_variables_End)
    print("\n========== Message Variable Generation End! ========== ")
    print("\n========== Testcase Generation Start! ========== ")
    tx_CCANFD_Var, tx_ECANFD_Var, tx_GCANFD_Var, tx_PCANFD_Var, tx_HSBCAN_Var, tx_HSMCAN_Var, tx_HSICAN_Var, tx_HSDCAN_Var = Gen_TxCycleTime(i_data)
    f.write(TxCycleTime_Start)
    f.write(tx_CCANFD_Var)
    f.write(TxCycleTime_End)
    print("\n========== Testcase Generation End! ========== ")
    f.close()

    return 0


def Gen_TxCycleTime(i_data):
    # Variables의 message 선언
    tx_CCANFD_Tc = ''
    tx_ECANFD_Tc = ''
    tx_GCANFD_Tc = ''
    tx_PCANFD_Tc = ''
    tx_HSBCAN_Tc = ''
    tx_HSMCAN_Tc = ''
    tx_HSICAN_Tc = ''
    tx_HSDCAN_Tc = ''
    test_list = i_data.values.tolist()
    for i, temp in tqdm(enumerate(test_list)):
        if i != 0:
            if temp[4] == 'FD-C':
                tx_CCANFD_Tc += make_Testcase_Text(i, temp)
            elif temp[4] == 'FD-E':
                tx_ECANFD_Tc += make_Testcase_Text(i, temp)
            elif temp[4] == 'FD-G':
                tx_GCANFD_Tc += make_Testcase_Text(i, temp)
            elif temp[4] == 'FD-P':
                tx_PCANFD_Tc += make_Testcase_Text(i, temp)
            elif temp[4] == 'HS-B':
                tx_HSBCAN_Tc += make_Testcase_Text(i, temp)
            elif temp[4] == 'HS-I':
                tx_HSICAN_Tc += make_Testcase_Text(i, temp)
            elif temp[4] == 'HS-M':
                tx_HSMCAN_Tc += make_Testcase_Text(i, temp)
            elif temp[4] == 'HS-D':
                tx_HSDCAN_Tc += make_Testcase_Text(i, temp)

    return tx_CCANFD_Tc, tx_ECANFD_Tc, tx_GCANFD_Tc, tx_PCANFD_Tc, tx_HSBCAN_Tc, tx_HSMCAN_Tc, tx_HSICAN_Tc, tx_HSDCAN_Tc


def Gen_Variables(i_data):
    # Variables의 message 선언
    tx_msg_CCANFD = []
    tx_msg_ECANFD = []
    tx_msg_GCANFD = []
    tx_msg_PCANFD = []
    tx_msg_HSBCAN = []
    tx_msg_HSMCAN = []
    tx_msg_HSICAN = []
    tx_msg_HSDCAN = []
    t_list = len(i_data[0]) - 1

    for i in tqdm(range(t_list)):
        if i != 0:
            temp = '  message ' + i_data[0][i] + ' ' + i_data[0][i] + ';'
            if i_data[4][i] == 'FD-C':
                tx_msg_CCANFD.append(temp)
            elif i_data[4][i] == 'FD-E':
                tx_msg_ECANFD.append(temp)
            elif i_data[4][i] == 'FD-G':
                tx_msg_GCANFD.append(temp)
            elif i_data[4][i] == 'FD-P':
                tx_msg_PCANFD.append(temp)
            elif i_data[4][i] == 'HS-B':
                tx_msg_HSBCAN.append(temp)
            elif i_data[4][i] == 'HS-I':
                tx_msg_HSICAN.append(temp)
            elif i_data[4][i] == 'HS-M':
                tx_msg_HSMCAN.append(temp)
            elif i_data[4][i] == 'HS-D':
                tx_msg_HSDCAN.append(temp)
    # 중복데이터 제거
    tx_msg_CCANFD = del_equal_data(tx_msg_CCANFD)
    tx_msg_ECANFD = del_equal_data(tx_msg_ECANFD)
    tx_msg_GCANFD = del_equal_data(tx_msg_GCANFD)
    tx_msg_PCANFD = del_equal_data(tx_msg_PCANFD)
    tx_msg_HSBCAN = del_equal_data(tx_msg_HSBCAN)
    tx_msg_HSMCAN = del_equal_data(tx_msg_HSMCAN)
    tx_msg_HSICAN = del_equal_data(tx_msg_HSICAN)
    tx_msg_HSDCAN = del_equal_data(tx_msg_HSDCAN)
    # 오른차순 정렬
    tx_msg_CCANFD.sort()
    tx_msg_ECANFD.sort()
    tx_msg_GCANFD.sort()
    tx_msg_PCANFD.sort()
    tx_msg_HSBCAN.sort()
    tx_msg_HSMCAN.sort()
    tx_msg_HSICAN.sort()
    tx_msg_HSDCAN.sort()

    tx_CCANFD_Var = make_Variables_Text(tx_msg_CCANFD)
    tx_ECANFD_Var = make_Variables_Text(tx_msg_ECANFD)
    tx_GCANFD_Var = make_Variables_Text(tx_msg_GCANFD)
    tx_PCANFD_Var = make_Variables_Text(tx_msg_PCANFD)
    tx_HSBCAN_Var = make_Variables_Text(tx_msg_HSBCAN)
    tx_HSMCAN_Var = make_Variables_Text(tx_msg_HSMCAN)
    tx_HSICAN_Var = make_Variables_Text(tx_msg_HSICAN)
    tx_HSDCAN_Var = make_Variables_Text(tx_msg_HSDCAN)

    return tx_CCANFD_Var, tx_ECANFD_Var, tx_GCANFD_Var, tx_PCANFD_Var, tx_HSBCAN_Var, tx_HSMCAN_Var, tx_HSICAN_Var, tx_HSDCAN_Var


def make_Variables_Text(i_list):
    mystr = ''
    for i in tqdm(range(len(i_list))):
        mystr += str(i_list[i]) + '\n'

    return mystr


def make_Testcase_Text(index, i_data):
    mystr = ''
    mystr += '    case ' + str(index) + ':\n'
    mystr += '	  ' + 'TC_' + i_data[0] + '_' + i_data[11] + '(' + i_data[8] + ');\n'
    mystr += """	  break;\n"""

    return mystr


def del_equal_data(i_list):
    new_list = []
    for v in i_list:
        if v not in new_list:
            new_list.append(v)
    return new_list


Tx_CAN_include = """includes
{

}
\n
"""

Tx_CAN_variables_Start = """variables
{
  char Replay_Cycle = 3;
  word Testcase_Num = 0;  
  word Test_Cycle = 0;
  word Tx_Cnt = 0;
  msTimer TxCycleTime;
  char toggle_data;
  /************** Message List **************/
"""
Tx_CAN_variables_End = """}\n"""

Tx_CAN_on_start = """on start
{
  setTimer(Send_timer, 2000);	
}
\n"""

TxCycleTime_Start = """\non timer TxCycleTime
{
  switch (Testcase_Num){
"""
TxCycleTime_End = """    default:
      stop();
      break;
  }
}"""
