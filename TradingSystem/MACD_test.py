import numpy as np
import pandas as pd
import openpyxl

df = pd.read_excel('input.xlsx')
numOfSamples = df.__len__()
data=np.mat(df)

infor_1 = 0
infor_2 = 0
infor_3 = 0

outData1=np.zeros((numOfSamples,5))
outData2=np.zeros((numOfSamples,3))
outData3=np.zeros((numOfSamples,6))

for i in range(numOfSamples):
    if i<11:
        EMA_12=data[i,1]
    else:
        EMA_12=EMA_12*11/13+data[i,1]*2/13
    print('12:',EMA_12)

    if i<25:
        EMA_26=data[i,1]
    else:
        EMA_26=EMA_26*25/27+data[i,1]*2/27
    print('26:',EMA_26)

    DIF=EMA_12-EMA_26
    print('DIF:',DIF)

    if i<34:
        DEA=DIF
    else:
        DEA=DEA*8/10+DIF*2/10
    print('DEA:',DEA)

    MACD=DIF-DEA
    print('MACD:',MACD)

    outData1[i] = [EMA_12,EMA_26,DIF,DEA,MACD]


for i in range(numOfSamples):
    if i>0:
        if ((outData1[i,2]>0.0) & (outData1[i,3]>0.0) & (outData1[i,2]>outData1[i-1,2]) & (outData1[i,3]>outData1[i-1,3])):
            infor_1 = 2
        if ((outData1[i,2]<0.0) & (outData1[i,3]<0.0) & (outData1[i,2]<outData1[i-1,2]) & (outData1[i,3]<outData1[i-1,3])):
            infor_1 = -2
        if ((outData1[i,2]>0.0) & (outData1[i,3]>0.0) & (outData1[i,2]<outData1[i-1,2]) & (outData1[i,3]<outData1[i-1,3])):
            infor_1 = -1
        if ((outData1[i,2]<0.0) & (outData1[i,3]<0.0) & (outData1[i,2]>outData1[i-1,2]) & (outData1[i,3]>outData1[i-1,3])):
            infor_1 = 1
        if ((outData1[i,2]>0.0) & (outData1[i,3]>0.0) & (outData1[i,2]>outData1[i,3]) & (outData1[i-1,2]<outData1[i-1,3])):
            infor_2 = 2
        if ((outData1[i,2]<0.0) & (outData1[i,3]<0.0) & (outData1[i,2]>outData1[i,3]) & (outData1[i-1,2]<outData1[i-1,3])):
            infor_2 = 1
        if ((outData1[i,2]>0.0) & (outData1[i,3]>0.0) & (outData1[i,2]<outData1[i,3]) & (outData1[i-1,2]>outData1[i-1,3])):
            infor_2 = -1
        if ((outData1[i,2]<0.0) & (outData1[i,3]<0.0) & (outData1[i,2]<outData1[i,3]) & (outData1[i-1,2]>outData1[i-1,3])):
            infor_2 = -2
        if ((outData1[i,4]>0.0) & (outData1[i,4]>outData1[i-1,4])):
            infor_3 = 2
        if ((outData1[i,4]<0.0) & (outData1[i,4]<outData1[i-1,4])):
            infor_3 = -2
        if ((outData1[i,4]>0.0) & (outData1[i,4]<outData1[i-1,4])):
            infor_3 = -1
        if ((outData1[i,4]<0.0) & (outData1[i,4]>outData1[i-1,4])):
            infor_3 = 1
    outData2[i] = [infor_1, infor_2, infor_3]



outData3=np.concatenate((outData1[:,2:5],outData2[:,:]),axis=1)

df = pd.DataFrame(outData3)
df.to_excel('output.xlsx', sheet_name='Sheet1')

# outData3 = pd.ExcelWriter('output.xlsx')
# df.to_excel(outData3)
# outData3.save()


def test12(a,b):
    print(a*11/13+b*2/13)