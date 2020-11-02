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
    f.write(Tx_CAN_variables_Start)
    f.write(tx_CCANFD_Var)
    f.write(tx_ECANFD_Var)
    f.write(tx_GCANFD_Var)
    f.write(tx_PCANFD_Var)
    f.write(tx_HSBCAN_Var)
    f.write(tx_HSMCAN_Var)
    f.write(tx_HSICAN_Var)
    f.write(Tx_CAN_variables_End)
    f.write(Tx_CAN_on_start)
    print("\n========== Message Variable Generation End! ========== ")
    print("\n========== Testcase Generation Start! ========== ")
    tx_CCANFD_Var, tx_ECANFD_Var, tx_GCANFD_Var, tx_PCANFD_Var, tx_HSBCAN_Var, tx_HSMCAN_Var, tx_HSICAN_Var, tx_HSDCAN_Var = Gen_TxCycleTime(i_data)
    f.write(TxCycleTime_Start)
    f.write(tx_CCANFD_Var)
    f.write(tx_ECANFD_Var)
    f.write(tx_GCANFD_Var)
    f.write(tx_PCANFD_Var)
    f.write(tx_HSBCAN_Var)
    f.write(tx_HSMCAN_Var)
    f.write(tx_HSICAN_Var)
    f.write(TxCycleTime_End)
    print("\n========== Testcase Generation End! ========== ")
    f.write(TxTestcode_Code)
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
    t_cnt = 0
    test_list = i_data.values.tolist()
    for i, temp in tqdm(enumerate(test_list)):
        if i != 0:
            if temp[4] == 'FD-C' and temp[13] != '삭제':
                tx_CCANFD_Tc += make_Testcase_Text(i, temp, 1)
            elif temp[4] == 'FD-E' and temp[13] != '삭제':
                tx_ECANFD_Tc += make_Testcase_Text(i, temp, 1)
            elif temp[4] == 'FD-G' and temp[13] != '삭제':
                tx_GCANFD_Tc += make_Testcase_Text(i, temp, 1)
            elif temp[4] == 'FD-P' and temp[13] != '삭제':
                tx_PCANFD_Tc += make_Testcase_Text(i, temp, 1)
            elif temp[4] == 'HS-B' and temp[13] != '삭제':
                tx_HSBCAN_Tc += make_Testcase_Text(i, temp, 0)
            elif temp[4] == 'HS-I' and temp[13] != '삭제':
                tx_HSICAN_Tc += make_Testcase_Text(i, temp, 0)
            elif temp[4] == 'HS-M' and temp[13] != '삭제':
                tx_HSMCAN_Tc += make_Testcase_Text(i, temp, 0)
            elif temp[4] == 'HS-D' and temp[13] != '삭제':
                tx_HSDCAN_Tc += make_Testcase_Text(i, temp, 0)

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
            if i_data[4][i] == 'FD-C' and i_data[13][i] != '삭제':
                tx_msg_CCANFD.append(temp)
            elif i_data[4][i] == 'FD-E' and i_data[13][i] != '삭제':
                tx_msg_ECANFD.append(temp)
            elif i_data[4][i] == 'FD-G' and i_data[13][i] != '삭제':
                tx_msg_GCANFD.append(temp)
            elif i_data[4][i] == 'FD-P' and i_data[13][i] != '삭제':
                tx_msg_PCANFD.append(temp)
            elif i_data[4][i] == 'HS-B' and i_data[13][i] != '삭제':
                tx_msg_HSBCAN.append(temp)
            elif i_data[4][i] == 'HS-I' and i_data[13][i] != '삭제':
                tx_msg_HSICAN.append(temp)
            elif i_data[4][i] == 'HS-M' and i_data[13][i] != '삭제':
                tx_msg_HSMCAN.append(temp)
            elif i_data[4][i] == 'HS-D' and i_data[13][i] != '삭제':
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


def make_Testcase_Text(index, i_data, i_brs):
    mystr = ''
    mystr += '    case ' + str(index) + ':\n'
    mystr += '	  ' + 'txTestcase' + '(' + i_data[0] + ', ' + i_data[2] + ', ' + i_data[3] + ', ' + i_data[8] + ', ' + str(i_brs) + ');\n'
    mystr += """	  break;\n"""

    return mystr


def del_equal_data(i_list):
    new_list = []
    for v in i_list:
        if v not in new_list:
            new_list.append(v)
    return new_list


Tx_CAN_include = """/*@!Encoding:949*/
includes
{

}
"""

Tx_CAN_variables_Start = """variables
{
  char Replay_Cycle = 1;
  word Testcase_Num = 0;  
  word Test_Cycle = 0;
  word Tx_Cnt = 0;
  msTimer TxCycleTime;
  char toggle_data;
  /************** Message List **************/
"""
Tx_CAN_variables_End = """}\n"""

Tx_CAN_on_start = """
on start
{
  setTimer(TxCycleTime, 2000);	
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

TxTestcode_Code = """
void txTestcase(message * i_message, char unit, word tx_CycleTime, word rx_CycleTime, char i_brs)
{
  word cycleTime = 0;
  i_message.BRS = i_brs;
  
  if(rx_CycleTime != 0){
    cycleTime = rx_CycleTime/tx_CycleTime;
    setTimer(TxCycleTime, tx_CycleTime);
  }
  else{
    cycleTime = 1;
    setTimer(TxCycleTime, 10);
  }
  
  if(Test_Cycle<Replay_Cycle){
    if(toggle_data == 0){
      Set_Data_Val(i_message,0x0,unit,cycleTime); 
    }
    else if (toggle_data == 1){
      Set_Data_Val(i_message,0x1,unit,cycleTime); 
    }
    else if (toggle_data == 2){
      Set_Data_Val(i_message,0x2,unit,cycleTime);  
    }
    else if (toggle_data == 3){
      Set_Data_Val(i_message,0x3,unit,cycleTime); 
    }
    else if (toggle_data == 4){
      Set_Data_Val(i_message,0x4,unit,cycleTime);
    }
    else{
      toggle_data=0;
      Test_Cycle++;
    }
  }
  else{
    Testcase_Num++;
    Test_Cycle=0;
  }
}

void Set_Data_Val(message * i_message, char val, char unit, char cycle)
{
  char i;
  for(i=0;i<unit;i++){
    i_message.byte(i) = val;
  }
  output(i_message);
  Tx_Cnt++;
  if(Tx_Cnt>=cycle){
    toggle_data++;
    Tx_Cnt=0;
  }
}"""