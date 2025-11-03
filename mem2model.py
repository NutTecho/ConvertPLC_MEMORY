import pandas as pd
import struct

# ======== Parameter setting =========
input_memory = "plcdata.csv"
output_model = "output.xlsx"

# =================================


def readstring(data):
	sumdata = ""
	convert = list(map(lambda x : struct.pack('H',x).decode("UTF-8"),data)) 
	for y in convert:
		sumdata += y
	return sumdata

def offsetindex(df,indexnum , datalength):   
    if datalength == 1 :
        row = indexnum//8
        col = indexnum%8
        return df.iat[row,col]
    else:
        i = 0
        dataout = []
        while ( i < datalength):
            row = (indexnum + i)//8
            col = (indexnum + i)%8
            dataout.append(df.iat[row,col])
            i+= 1
        return dataout
                   
def main():
    df = pd.read_csv(input_memory,sep="\t",encoding="cp1252",engine='python' ,
                     header=0,usecols=[1,2,3,4,5,6,7,8])
    # print(df)
    newdata = []
    indexrow = 0
    ncount = len(df) * 8
    while (indexrow < ncount -1):
        # mcode = readstring(df.iloc[row,colstart:colstart + 3].values.tolist())

        try:
            zstart = indexrow
            mcode = readstring(offsetindex(df,indexrow,3))
            indexrow += 4
            # mname = readstring(df.iloc[1,4:9].values.tolist() +  df.iloc[1 + i,1:6].values.tolist())
            mname = readstring(offsetindex(df,indexrow,10)).replace('\x00','')
            indexrow += 10

            pattern = offsetindex(df,indexrow,1)
            indexrow += 1

            H_type = offsetindex(df,indexrow,1)
            indexrow += 1

            di_W = offsetindex(df,indexrow,1)
            indexrow += 1

            di_L = offsetindex(df,indexrow,1)
            indexrow += 1

            di_H =  offsetindex(df,indexrow,1)
            indexrow += 1

            work_max = offsetindex(df,indexrow,1)
            indexrow += 1

            layer_max =  offsetindex(df,indexrow,1)
            indexrow += 1

            prog_rb =  offsetindex(df,indexrow,1)
            indexrow += 1

            prog_iai =  offsetindex(df,indexrow,1)
            indexrow += 1

            spare1 = offsetindex(df,indexrow,1)
            indexrow += 1

            spare2 = offsetindex(df,indexrow,1)
            indexrow += 1

            zend = indexrow

            dict_data = { "zrstart" : zstart , "zend": zend  , "mcode" : mcode , "mname" : mname, "pattern": pattern ,
                            "H_type" : H_type , "di_W" : di_W , "di_L" : di_L , "di_H" : di_H,
                            "work_max" : work_max , "layer_max" : layer_max , "prog_rb" : prog_rb,
                            "prog_iai" : prog_iai, "spare1" : spare1 , "spare2" : spare2 }
            newdata.append(dict_data)
        
        except Exception as e:
            print(indexrow)
            print(e)

    df2 = pd.DataFrame(newdata)
    df2.to_excel(output_model,sheet_name="data1")
    print(df2)

if __name__ == "__main__":
    main()