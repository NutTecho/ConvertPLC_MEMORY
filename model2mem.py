import struct
import pandas as pd


# ======== Parameter setting ==============
input_model = "output1.xlsx"
output_file = "datamem.csv"

# After export to data memory
# Copy header format from previous file to this file

# =========================================

def byteconvert(strdata , nlength):
    data = strdata.encode("utf-8")
    i = 0
    list_data = list(data)
    while (len(list_data) < (nlength * 2)):
        list_data.append(0)
    
    output = []
    # print(list_data)
    while (i < len(list_data) - 1):
        output.append(int.from_bytes([list_data[i+1] , list_data[i]]))
        # data2 =  struct.pack('HHH', data)
        # print(data)
        i +=2
    return output

def main():
    df = pd.read_excel(input_model)
    xcolumns = ["'+0","'+1","'+2","'+3","'+4","'+5","'+6","'+7"]
    df2 =  pd.DataFrame([{}] ,columns=xcolumns)
    df2 = df2.fillna(0)
    df2 = df2.astype(int)
    
    indexnum = 0
    for row in df.itertuples():
        mcode = byteconvert(row[4],3)
        mname = byteconvert(row[5],10)
        pattern = row[6]
        H_type= row[7]
        di_W = row[8]
        di_L = row[9]
        di_H = row[10]
        work_max = row[11]
        layer_max = row[12]
        prog_rb = row[13]
        prog_iai = row[14]
        spare1 = row[15]
        spare2 = row[16]

        full_data = mcode + [0] + mname + [pattern] + [H_type] + [di_W] + [di_L] + [di_H] + [work_max] + [layer_max] + [prog_rb] + [prog_iai] + [spare1] + [spare2]

        i = 0
        print(row[4])
        while (i < 25):
            row = (indexnum + i)//8
            col = (indexnum + i)%8
            if (i > 0 and (indexnum + i)%7 == 0):
                # df.append(pd.Series(), ignore_index=True)
                df2.loc[len(df2)] = pd.Series(dtype='int64')
                df2 = df2.fillna(0)
                df2 = df2.astype(int)
            df2.iloc[row,col] = int(full_data[i])

            i += 1
        indexnum += 25
        
    zrdata = []
    for i in range(len(df2)):
        zrdata.append("ZR" + str(i*8))
    df2.insert(0,"Device Name",zrdata)
    print(df2)
    df2.to_csv(output_file,sep='\t', encoding='UTF-16LE' ,index=False)
    
if __name__ == "__main__":
    main()