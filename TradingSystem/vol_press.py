import numpy as np
import pandas as pd
import datetime


df = pd.read_excel('input_000016.xlsx')
da = np.mat(df)

numOfSamples = df.__len__()
infor_1 = np.zeros((numOfSamples, 1))
infor_2 = np.zeros((numOfSamples, 1))
obvma = np.zeros((numOfSamples, 1))
date = np.zeros((numOfSamples, 1))
close = np.zeros((numOfSamples, 1))

for i in range(0, numOfSamples):
    obvma[i, 0] = np.mean(da[0:i+1, 3])
    if i > 1:
        if (da[i, 1] > da[i - 1, 1]) and (da[i, 2] > da[i - 1, 2]):
            infor_1[i, 0] = 1

            date_time1 = datetime.datetime.strptime(da[i, 0], "'%Y%m%d'") #把时间转成字符串
            str_date_time1 = date_time1.strftime('%Y%m%d')
            date[i, 0] = str_date_time1

            close[i, 0] = da[i, 1]
        if (da[i, 1] < da[i - 1, 1]) and (da[i, 2] < da[i - 1, 2]):
            infor_1[i, 0] = -1

            date_time1 = datetime.datetime.strptime(da[i, 0], "'%Y%m%d'")
            str_date_time1 = date_time1.strftime('%Y%m%d')
            date[i, 0] = str_date_time1

            close[i, 0] = da[i, 1]
        if (da[i, 3] > obvma[i, 0]) and (da[i - 1, 3] < obvma[i - 1, 0]):
            infor_2[i, 0] = 1

            date_time1 = datetime.datetime.strptime(da[i, 0], "'%Y%m%d'")  # 把时间转成字符串
            str_date_time1 = date_time1.strftime('%Y%m%d')
            date[i, 0] = str_date_time1

            close[i, 0] = da[i, 1]
        if (da[i, 3] < obvma[i, 0]) and (da[i - 1, 3] < obvma[i - 1, 0]):
            infor_2[i, 0] = -1

            date_time1 = datetime.datetime.strptime(da[i, 0], "'%Y%m%d'")  # 把时间转成字符串
            str_date_time1 = date_time1.strftime('%Y%m%d')
            date[i, 0] = str_date_time1

            close[i, 0] = da[i, 1]

df1 = pd.DataFrame(np.column_stack((infor_1, infor_2,)),
                   columns=['infor_1', 'infor_2'])
df1.to_excel('table1_output.xlsx')

df2 = pd.DataFrame(np.column_stack((infor_1, infor_2, date, close)),
                   columns=['infor_1', 'infor_2', 'date', 'close'])

df2.to_excel('table2_output.xlsx')