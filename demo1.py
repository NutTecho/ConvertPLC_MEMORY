import struct
import pandas as pd
# data = tuple(28836,16709)
# packed_string = struct.pack("h", 16709)
# # unpacked_float = struct.unpack("f", packed_string)[0]

# print(packed_string)

def readstring(data):
	sumdata = ""
	convert = list(map(lambda x : struct.pack('H',x).decode("UTF-8"),data)) 
	for y in convert:
		sumdata += y
	return sumdata.replace('\x00','')

def byteconvert(strdata , nlength):
    data = strdata.encode("utf-8")
    xs = bytearray(data)
    i = 0
    list_data = list(data)

    #  ====== concept 1 =======
    while (len(list_data) < (nlength * 2)):
        xs.append(0)
        list_data.append(0)
    output = struct.unpack('h' * int(len(xs) / 2) , xs)
    
    #  ====== concept 2 =========

    while (i < len(list_data) - 1):
        output.append(int.from_bytes([list_data[i] , list_data[i+1]],'little',signed=True))
        # data2 =  struct.pack('HHH', data)
        # print(data)
        i +=2

    return output

def main():
    df = pd.read_csv('MAIN_D2.csv',sep="\t",encoding="utf16",
                     skiprows=7, header=0,
                     usecols=[1,2,3,4,5,6,7,8])
    df = df.fillna(0)
    df = df.astype(int)
    print(df)
    for i in df.itertuples():
      mcode = readstring(i[1:7])
      print(mcode)

def main2():
    info = ["Thailand" , "Japan" , "Spain" , "Canada" , "Maxico"]
    xcolumns = ["'+0","'+1","'+2","'+3","'+4","'+5","'+6","'+7"]
    df2 =  pd.DataFrame([{}] ,columns=xcolumns)
    df2 = df2.fillna(0)
    df2 = df2.astype(int)
    for i,v in enumerate(info):
        cv =  byteconvert(v,8)
        for y,z in  enumerate(cv):
            df2.iloc[i,y] = int(z)
        df2.loc[len(df2)] = pd.Series(dtype='int64')
        df2 = df2.fillna(0)
        df2 = df2.astype(int)
         

    zrdata = []
    for i in range(len(df2)):
        zrdata.append("D" + str(i*8))
    df2.insert(0,"Device Name",zrdata)
    print(df2)
    # df2.to_csv('test2.csv',sep='\t', encoding='UTF-16LE' ,index=False)

def main3():
    # data = list(map(lambda x : struct.pack('H',x).decode("UTF-8"),data)) 
    # pack1 = struct.pack('s',*data)
    strdata = "Thailand"
    data = strdata.encode("utf-8")
    # y = [84, 104, 97, 105, 108, 97, 110, 100]
    # y = [84, 104]
    # x = struct.unpack('c',data)
    print(x)
    x = struct.unpack()
    print(int.from_bytes(bytes=['A','B'],byteorder='little',signed=True))


if __name__ == "__main__":
    # ==== read CSV from data memory and convert to string ====
    main()

    # ==== convert string to WORD format for data memory
    # main2()

