import pandas as pd
import numpy as np
import string
import os
#%% Assign src and destination path
def get_input_path():
    return input("Enter the input file path: ")

def get_output_path():
    return input("Enter the output file path: ")

def excel_column_name(n):
    letters = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def generate_new_column_names(num_columns):
    columns = []
    for i in range(1, num_columns + 1):
        column_name = ""
        while i > 0:
            i, remainder = divmod(i - 1, 26)
            column_name = chr(65 + remainder) + column_name
        columns.append(column_name)
    return columns

def create_empty_dataframe(num_columns):
    column_names = generate_new_column_names(num_columns)
    df = pd.DataFrame(columns=column_names)
    return df

# Call the function with the desired number of columns (117 in this case)
num_columns = 117
dest_df = create_empty_dataframe(num_columns)

def main():
    src_file_path = get_input_path()
    src_file_path= src_file_path[1:-1]
    dest_file_path = get_output_path()
    dest_file_path=dest_file_path[1:-1]
    # Check if the input file exists
    if not os.path.exists(src_file_path):
        print("Input file does not exist.")
        return
    
    col_A_BE=["(LRU)ORIGINE","(LRU)RPT_ORG_NAME","(LRU)RPT_ORG_CODE","(LRU)SHOP_LOC_TEXT","(LRU)REC_STAT",
          "(LRU)OPER_CODE","OPER_NAME","(LRU)AC_MDL_NO","(LRU)AC_SERIES_NO","AC_ENGINE_TYPE",
          "ENG_POSITION_CODE","(LRU)AC_MFG_SER_NO","(LRU)AC_REG_NO","(LRU)AC_ID_NO","CUM_TOT_FLGT_HR",
          "CUM_TOT_LNDG_QT","(LRU)REMV_REASON_TEXT","MAINT_ACTION_TEXT","AC_MSG_TEXT","AC_MSG_CODE",
          "(LRU)REMV_DATE","STATION_CODE","OPER_EVENT_ID","(LRU)REMV_MANUF_PART_NO",
          "(LRU)REMV_OPER_PART_NO","(LRU)REMV_MANUF_SER_NO","(LRU)REMV_OPER_SER_NO",
          "(LRU)REMV_MANUF_CODE","REMV_MOD_LVL","REMV_SOFT_MOD_LVL","(LRU)REMV_TYPE_CODE",
          "(LRU)REMV_TYPE_TEXT","NHA_MANUF_PART_NO","NHA_OPER_PART_NO","NHA_MANUF_SER_NO",
          "NHA_OPER_SER_NO","PART_POSITION_CODE","ATA_NO","INST_MANUF_PART_NO","INST_OPER_PART_NO",
          "INST_MANUF_SER_NO","INST_OPER_SER_NO","INST_MANUF_CODE","INST_MOD_LVL","INST_SOFT_MOD_LVL",
          "TSN_HRS","TSI_HRS","TSR_HRS","TSC_HRS","TSV_HRS","TSO_HRS","CSN_LDGS","CSI_LDGS","CSR_LDGS",
          "CSC_LDGS","CSV_LDGS","CSO_LDGS"]
    
    col_BF_end=["RMV_TRACK_ID","RLTD_SHOP_RCRD_ID","SHOP_RCRD_ID","ORIGINE","RPT_ORG_NAME","RPT_ORG_CODE","SHOP_LOC_TEXT",
          "REC_STAT","RCVD_DATE","PURCHASER_NAME","PURCHASER_CODE","OPER_CODE","AC_MDL_NO","AC_SERIES_NO",
          "REG_NO","AC_REG_NO","AC_ID_NO","REMV_DATE","REMV_MANUF_PART_NO","REMV_OPER_PART_NO","REMV_MANUF_SER_NO",
          "V_OPER_SER_NO","REMV_MANUF_CODE","REMV_TYPE_TEXT","RECV_MANUF_PART_NO","RECV_OPERATOR_PART_NO","MANUF_PART_DESC","ACRONYM",
          "RECV_MANUF_SER_NO","RECV_OPERATOR_SER_NO","REC_MANUF_CODE","INC_MOD_LVL","INC_SOFT_MOD_LVL","REMV_REASON_TEXT",
          "INC_INSP_TEXT","SHOP_ACT_TEXT","REMV_TYPE_CODE","FLR_FND_IN","FLR_INDC_IN","FLR_CNFRM_IN","AC_MSG_CNFRM_IN",
          "AC_BITE_CNFRM_IN","FLR_TYPE_IN","SHIP_MANUF_PART_NO","SHIP_OPER_PART_NO","SHIP_MANUF_SER_NO","SHIP_OPER_SER_NO",
          "SHIP_MANUF_CODE","SHIP_DATE","SHIP_MOD_LVL","SHIP_SOFT_MOD_LVL","MOD_INC_TEXT","SOFT_MOD_INC_TXT","SHOP_FNDG_CODE",
          "INT_INFO","C","F","WORK_ORDER","NOTIFICATION_NUMBER","SB_APPLIED","","","DP"]

    out_file_col=col_A_BE+col_BF_end

    src_file_df=pd.read_excel(src_file_path,sheet_name=0,index_col=None)
    src_file_col=src_file_df.columns

    new_column_names = [f'col_{excel_column_name(i)}' for i in range(1, len(src_file_df.columns)+1)]
    src_file_df.columns=new_column_names

    dest_df["BN"]=pd.to_datetime(src_file_df['col_F']).dt.strftime('%d/%m/%Y')
    dest_df["BW"]=pd.to_datetime(src_file_df['col_B']).dt.strftime('%d/%m/%Y')
    dest_df["BR"]=src_file_df['col_Y']
    dest_df["BT"]=src_file_df['col_Z']
    dest_df["BT"].fillna("UNK",inplace=True)
    dest_df["BX"],dest_df["CD"],dest_df["CW"]=src_file_df['col_H'],src_file_df['col_H'],src_file_df['col_H']
    dest_df["BZ"],dest_df["CH"],dest_df["CY"]=src_file_df['col_M'],src_file_df['col_M'],src_file_df['col_M']
    dest_df["CM"]=src_file_df['col_T']
    dest_df["CN"]=src_file_df['col_X']

    dest_df["CO"]=src_file_df['col_AX']
    dest_df["DB"]=pd.to_datetime(src_file_df['col_AV']).dt.strftime('%d/%m/%Y')
    dest_df["DH"]=src_file_df['col_AY']
    dest_df["DK"]=src_file_df['col_A']
    dest_df["BF"]=src_file_df['col_C']
    dest_df["BH"]=src_file_df['col_E']
    dest_df["BI"]="IFE_AUG23_SAPUS"
    dest_df["BO"]=src_file_df['col_AK']

    #% Encode AG column to be filled in BQ of the output file:

    oper_code_encoding=src_file_df[['col_AG','col_AK']]
    oper_code_encoding['col_AG'].value_counts()

    oper_code = {"col_AG": {"JETBLUE AIRWAYS": "JBU", "AIR CANADA": "ACN", "UNITED AIRLINES": "UAL", 
                        "BRITISH AIRWAYS PLC": "BAW","SAUDI ARABIAN AIRLINES": "SVA", 
                        "QATAR AIRWAYS GROUP -TECH ACCT": "QATAR", "OMAN AIR":"OMA","AIR INDIA LIMITED":
                        "AIC","JAZZ AVIATION LP":"JAZ","GULF AIR COMPANY G.S.C.":"GAC",
                        "ETHIOPIAN AIRLINES":"ETH","Polskie Linie Lotnicze LOT S.A":"LOT",
                        "JAPAN AIRLINES CO,LTD":"JAL","CHINA EASTERN AVIATION":"CEA","ROYAL AIR MAROC":
                        "RAM","AIR CHINA IMPORT/EXPORT":"ACIE","BEIJING CAPITAL AIRLINES CO, L":"BCA",
                        "ROYAL JORDANIAN AIRLINE":"RJA","ETHIOPIAN AIRLINES CORP":"ETH",
                        "SINGAPORE AIRCRAFT LEASI":"SNGA","ETIHAD AIRWAYS":"ETH",
                        "Total Havacilik Ic ve Dis Tic.":"AHY","THALES AEROSPACE (BEIJING) CO.":"AIC",
                        "HAINAN AIRLINES CO LTD":"HAC","TURKISH AIRLINES TECHNIC":"THY"}}  
    oper_df = oper_code_encoding.replace(oper_code)
    oper_df['col_AG'].fillna("XXX",inplace=True)  

    dest_df["BQ"]=oper_df['col_AG']
    dest_df["CB"]="0LWJ2"
    dest_df["CJ"]="0LWJ2"   
    dest_df["DA"]="0LWJ2"  
    dest_df["DN"]=np.nan
    dest_df["DO"]=np.nan
    dest_df["DP"]=pd.Series(range(1,len(src_file_df['col_F'])))



    dest_df.columns=out_file_col

    dest_df.to_csv(dest_file_path+"/output_file.csv",index=False)




if __name__ == "__main__":
    main()

