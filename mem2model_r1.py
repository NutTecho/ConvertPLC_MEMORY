import pandas as pd
import numpy as np
import struct
import csv
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

memory_file = r"D:\VSCODE\MyPython\mc1\MAIN_ZR1.csv"
mark_file = r"D:\VSCODE\MyPython\mc1\mark_file1.xlsx"
excel_export = r"D:\VSCODE\MyPython\mc1\runcode\output1.xlsx"
plctype = "GX_WORK3"

def tohex_twos_complement(val, nbits):
    # Mask to keep only the relevant bits
    return hex(val & ((1 << nbits) - 1))

def readstring(data):
	# sumdata = ""
	# convert = list(map(lambda x : struct.pack('H',x).decode("UTF-8"),data)) 
	# for y in convert:
	# 	sumdata += y
    sumdata = ""
    sum1 = pd.Series(data).sum()
    if sum1 != 0:
        hex_list = [tohex_twos_complement(n,16)[2:].zfill(4) for n in data]
        for i in hex_list:
            setstr = f'{i[2:]}{i[:2]}'
            # print(setstr)
            sumdata += bytes.fromhex(f'{setstr}').decode("Windows-1252'")
        # sumdata = struct.pack('H'*len(data),*data).decode("UTF-8")
    return sumdata

def readfloat(data):
	# sumdata = ""
	# convert = list(map(lambda x : struct.pack('H',x).decode("UTF-8"),data)) 
	# for y in convert:
	# 	sumdata += y
    sumdata = ""
    sum1 = pd.Series(data).sum()
    if sum1 != 0:
        # sumdata = struct.pack('i'*len(data),*data).decode("UTF-8")
        hex_list = [tohex_twos_complement(n,16)[2:].zfill(4) for n in data]
        byte_data = bytes.fromhex(f'{hex_list[1]}{hex_list[0]}')
        sumdata = struct.unpack('>f', byte_data)[0]
    return sumdata

def readdword(data):
	# sumdata = ""
	# convert = list(map(lambda x : struct.pack('H',x).decode("UTF-8"),data)) 
	# for y in convert:
	# 	sumdata += y
    sumdata = ""
    sum1 = pd.Series(data).sum()
    if sum1 != 0:
        hex_list = [tohex_twos_complement(n,16)[2:].zfill(4) for n in data]
        # sumdata = struct.pack('f'*len(data),*data).decode("UTF-8")
        sumdata = int(f'0x{hex_list[1]}{hex_list[0]}',0)
    return sumdata

def offsetindex(df,indexnum , datalength):  

    if datalength == 1 :
        row = indexnum//10
        col = indexnum%10
        return df.iat[row,col]
    
    else:
        i = 0
        dataout = []
        while ( i < datalength):
            row = (indexnum + i)//10
            col = (indexnum + i)%10
            dataout.append(df.iat[row,col])
            i+= 1
        return dataout

def memory_to_excel(memory_file,markdata,plctype):
    if plctype == "GX_WORK3":
        df = pd.read_csv(memory_file,sep="\t",encoding="UTF-16LE",engine='python' ,
                        skiprows=7,header=0,
                        usecols=[1,2,3,4,5,6,7,8,9,10])
    
    if plctype == "GX_WORK2":
        df = pd.read_excel(memory_file, usecols=[1,2,3,4,5,6,7,8,9,10] , engine='xlrd')
    
    df = df.fillna(0)
    df = df.convert_dtypes(convert_integer=True)
    
    newdata = []
    indexrow = 0
    ncount = len(df) * 10

    try:
        while (indexrow < ncount -1):
        # mcode = readstring(df.iloc[row,colstart:colstart + 3].values.tolist())
            dict_data = {}
            for index , row in markdata.iterrows():
                if (row['DATATYPE'] == "STRING"):
                    data = readstring(offsetindex(df,indexrow,int(row['LENGTH'])))
                    indexrow += int(row['LENGTH'])

                if (row['DATATYPE'] == "WORD"):
                    data = offsetindex(df,indexrow,int(row['LENGTH']))
                    indexrow += int(row['LENGTH'])

                if (row['DATATYPE'] == "DWORD"):
                    data = readdword(offsetindex(df,indexrow,int(row['LENGTH'])*2))
                    indexrow += int(row['LENGTH'])*2

                if (row['DATATYPE'] == "FLOAT"):
                    data = readfloat(offsetindex(df,indexrow,int(row['LENGTH'])*2))
                    indexrow += int(row['LENGTH'])*2

                if (row['COLNAME'] != "BLANK"):
                    dict_data[row['COLNAME']] = data
                    
            # print(dict_data)
            newdata.append(dict_data)
        
    except Exception as e:
        print(e)
        pass

    df2 = pd.DataFrame(newdata)
    print(df2)
    for col in df2.columns:
        if df2[col].dtype == 'object': # Check if the column is of string/object type
            df2[col] = df2[col].astype(str).apply(lambda x: ILLEGAL_CHARACTERS_RE.sub('', x))

    df2.index = np.arange(1, len(df2) + 1)
    return df2

def convert_mem_to_excel():
    try:

        markdata = pd.read_excel(mark_file,engine='openpyxl')
        markdata['STARTAD'] = markdata['STARTAD'].astype('int')
        markdata['LENGTH'] = markdata['LENGTH'].astype('int')
        
        df2 = memory_to_excel(memory_file,markdata,plctype)
        df2.to_excel(excel_export , engine='openpyxl')
        print(df2)
        print("==== convert memory to excel complete ====")
        # status_convert_mem_to_excel.set("=== Convert complete ===")
    except Exception as e:
        # status_convert_mem_to_excel.set(f"=== Convert Error {e} ===")
        print(e)

if __name__ == "__main__":
    convert_mem_to_excel()