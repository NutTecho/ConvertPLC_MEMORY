import pandas as pd
import numpy as np
import struct
import csv
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

memory_export = "R_ZR1.csv"
mark_file = "mark_file1.xlsx"
excel_import = "output1.xlsx"
plctype = "GX WORK3"
zrstart = 0

def STRING_convert(strdata , nlength):
    data = strdata.encode("utf-8")
    xs = bytearray(data)
    i = 0
    list_data = list(data)
    while (len(list_data) < (nlength * 2)):
        xs.append(0)
        list_data.append(0)
    
    output = []
    # while (i < len(list_data) - 1):
    #     output.append(int.from_bytes([list_data[i+1] , list_data[i]]))
    #     i +=2
    output = struct.unpack('h'*int(len(xs)/2),xs)
    return output

def DWORD_convert(dworddata):
    packed_bytes = struct.pack('<I', dworddata)

    hex_list = [f"{i:02x}" for i in packed_bytes]
    # print(hex_list)
    
    byte_data1 = bytes.fromhex(f'{hex_list[0]}{hex_list[1]}')
    byte_data2 = bytes.fromhex(f'{hex_list[2]}{hex_list[3]}')
    unsigned_int_le1 = int.from_bytes(byte_data1, byteorder='little', signed=True)
    unsigned_int_le2 = int.from_bytes(byte_data2, byteorder='little', signed=True)
    # print([unsigned_int_le1 , unsigned_int_le2])
    return [unsigned_int_le1 , unsigned_int_le2]

def FLOAT_convert(dworddata):
    packed_bytes = struct.pack('<f', dworddata)

    hex_list = [f"{i:02x}" for i in packed_bytes]
    # print(hex_list)
    
    byte_data1 = bytes.fromhex(f'{hex_list[0]}{hex_list[1]}')
    byte_data2 = bytes.fromhex(f'{hex_list[2]}{hex_list[3]}')
    unsigned_int_le1 = int.from_bytes(byte_data1, byteorder='little', signed=True)
    unsigned_int_le2 = int.from_bytes(byte_data2, byteorder='little', signed=True)
    # print([unsigned_int_le1 , unsigned_int_le2])
    return [unsigned_int_le1 , unsigned_int_le2]

def excel_to_memory(df,markdata,zrstart,plctype):
    xcolumns = ["'+0","'+1","'+2","'+3","'+4","'+5","'+6","'+7","'+8","'+9"]
    df2 =  pd.DataFrame([{}] ,columns=xcolumns)
    df2 = df2.fillna(0)
    df2 = df2.astype(int)

    # sumlength = markdata['LENGTH'].sum()
    # markdata[markdata['DATATYPE'] != 'WORD']['LENGTH'].sum()

    sumlength = markdata.loc[
                        (markdata['DATATYPE'] == 'WORD') | (markdata['DATATYPE'] == 'STRING'),
                        'LENGTH'
                        ].sum() + (markdata.loc[
                        (markdata['DATATYPE'] != 'WORD') & (markdata['DATATYPE'] != 'STRING'),
                        'LENGTH'
                        ].sum()*2)
    indexnum = 0

    for data_row in range(0, len(df)):
        print(f"read row : {data_row}")
        # status_convert_excel_to_mem.set(f"read row : {data_row}")
        # status_convert_excel_to_mem.set("=== Convert complete ===")
        full_data = []
        mark_row = 0
        use_row = 0
        while mark_row < (len(markdata)):
            
            if (markdata.iloc[mark_row,3] == "STRING" and markdata.iloc[mark_row,0] != "BLANK"):
                data = STRING_convert(df.iloc[data_row,use_row],markdata.iloc[mark_row,2])
                use_row += 1

            if (markdata.iloc[mark_row,3] == "DWORD" and markdata.iloc[mark_row,0] != "BLANK"):
                data = DWORD_convert(df.iloc[data_row,use_row])
                use_row += 1

            if (markdata.iloc[mark_row,3] == "FLOAT" and markdata.iloc[mark_row,0] != "BLANK"):
                data = FLOAT_convert(df.iloc[data_row,use_row])
                use_row += 1

            if(markdata.iloc[mark_row,3] == "WORD" and markdata.iloc[mark_row,0] != "BLANK"):
                data = [df.iloc[data_row,use_row]]
                use_row += 1

            if (markdata.iloc[mark_row,0] == "BLANK"):
                data = [0]*markdata.iloc[mark_row,2]

            mark_row += 1
            full_data += list(data)
        
        i = 0
        while (i < sumlength):
            if ( (indexnum + i)%10 == 0):
                df2.loc[len(df2)] = pd.Series(dtype='int64')
                df2 = df2.fillna(0)
                df2 = df2.astype(int)
                # print(len(df2))
                # print(df2)

            row = (indexnum + i)//10
            col = (indexnum + i)%10
            df2.iloc[row,col] = int(full_data[i])

            i += 1
        indexnum += sumlength
    
    zrdata = []
    for i in range(len(df2)):
        zrdata.append("ZR" + str(int(zrstart) +(i*10)))
    
    if plctype == "GX_WORK3":
        df2.insert(0,"Device Name",zrdata)
    if plctype == "GX_WORK2":
        df2.insert(0,"",zrdata)
    return df2

def convert_excel_to_mem():
    try:

        markdata = pd.read_excel(mark_file,engine='openpyxl')
        markdata['STARTAD'] = markdata['STARTAD'].astype('int')
        markdata['LENGTH'] = markdata['LENGTH'].astype('int')

        if plctype == "GX_WORK3":
            header = {  "GenerateMemory" : "",
                        "Data Display Format" : "16-bit [Signed]" ,
                        "Value" : "DEX" , 
                        "Bit Order" : "0-F" ,
                        "Number of Device Points to Display in 1 Row" : 10, 
                        "Export only the rows in which devices already set are included" : "Disable",
                        "": ""}
            
            # df_header = pd.DataFrame.from_dict(header)
            df_header = pd.DataFrame({'col1': list(header.keys()) ,'col2': list(header.values()) } )
            # df_header = df_header.reset_index(drop=True)
            df_header.to_csv(memory_export,mode='w',sep='\t', encoding='UTF-16LE' ,index=False , header=False)

        df = pd.read_excel(excel_import , engine="openpyxl",index_col=0)
        print(df)

        for index , row in markdata.iterrows():
            if (row['DATATYPE'] == "WORD" and row['COLNAME'] != "BLANK"):
                df[row['COLNAME']] = df[row['COLNAME']].astype(int) 

        df2 = excel_to_memory(df,markdata,zrstart,plctype)
        
        if plctype == "GX_WORK3":
            df2.to_csv(memory_export, mode='a', sep='\t', encoding='UTF-16LE' ,index=False)
           

        if plctype == "GX_WORK2":
            df2.to_excel(memory_export,engine='openpyxl', index=False)
           
        print("====convert excel to memory complete")
        # status_convert_excel_to_mem.set("=== Convert complete ===")
    except Exception as e:
        # status_convert_excel_to_mem.set(f"=== Convert Error {e} ===")
        print(e)

if __name__ == "__main__":
    convert_excel_to_mem()
