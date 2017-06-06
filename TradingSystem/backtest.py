import numpy as np
import pandas as pd
import math
import datetime

df = pd.read_excel('huice_input.xlsx')
da = np.mat(df)

rf = 0.02
numofDays = 0
numOfSamples = df.__len__()
unit = np.zeros((numOfSamples, 1))
acc_p = np.zeros((numOfSamples, 1))
acc_r = np.zeros((numOfSamples, 1))
sharp_r = np.zeros((numOfSamples, 1))
d = np.zeros((numOfSamples, 1))
IR = np.zeros((numOfSamples, 1))
beta = np.zeros((numOfSamples, 1))
alpha = np.zeros((numOfSamples, 1))
unit_cost = np.zeros((numOfSamples, 1))
ann_r = np.zeros((numOfSamples, 1))
ann_p = np.zeros((numOfSamples, 1))
u = np.zeros((numOfSamples, 1))

# 初始unit第一个元素为1
unit[0, 0] = 1.0
for i in range(1, numOfSamples):
    if da[i, 2] == da[i - 1, 2]:  # m不变的情况
        unit[i, 0] = 0
    else:
        unit[i, 0] = da[i, 1] - da[i - 1, 1]

unit_cost = np.multiply(unit, da[:, 4])
unitall_price = np.multiply(da[:, 1], da[:, 5])

for i in range(0, numOfSamples):
    acc_p[i, 0] = (da[i, 5] - da[0, 5]) / da[0, 5]

    m = da[i, 2]  # 求m
    # unit和cost的累加
    unitpluscost = sum(unit_cost[0:m, 0])
    acc_r[i, 0] = (da[i, 1] * da[i, 5] - unitpluscost) / unitpluscost

    # if i > 1:
    #     date_time1 = datetime.datetime.strptime(da[i - 1, 0], "'%Y%m%d'")
    #     date_time2 = datetime.datetime.strptime(da[i, 0], "'%Y%m%d'")
    #     # str_date_time1 = date_time1.strftime('%Y%m%d')
    #     # str_date_time2 = date_time2.strftime('%Y%m%d')
    #     numofDays = (date_time2-date_time1).days  # 实际天数
    # else:
    #     numofDays = 1
    # ann_r[i, 0] = math.pow(1 + acc_r[i, 0], numofDays / 365) - 1

    value_max = np.max(unitall_price[0:i + 1, 0] - unitpluscost)

    d[i, 0] = np.max(1 - (unitall_price[0:i + 1, 0] / value_max))

    alpha[i, 0] = acc_r[i, 0] - acc_p[i, 0]

# 求sharp_r和IR,初始sharp_r和IR的第一个元素为1
sharp_r[0, 0] = 1.0
IR[0, 0] = 1.0
for i in range(1, numOfSamples):
    date_time1 = datetime.datetime.strptime(da[i - 1, 0], "'%Y%m%d'")
    date_time2 = datetime.datetime.strptime(da[i, 0], "'%Y%m%d'")
    numofDays = (date_time2 - date_time1).days  # 实际天数
    ann_r[i, 0] = math.pow(1 + acc_r[i, 0], numofDays / 365) - 1

    dv = np.std(ann_r[0:i + 1, 0])
    print('dv:', dv)
    sharp_r[i, 0] = (ann_r[i, 0] - rf) / dv

    ann_p[i, 0] = math.pow(1 + acc_p[i, 0], numofDays / 365) - 1
    u[i, 0] = ann_r[i, 0] - ann_p[i, 0]
    IR[i, 0] = np.mean(u[0:i, 0]) / np.std(u[0:i + 1, 0])

    beta[i, 0] = np.cov((ann_r[i, 0], ann_p[i, 0])) / np.var(ann_p[0:i + 1, 0])

df1 = pd.DataFrame(np.column_stack((acc_p, acc_r, sharp_r, d, IR, beta, alpha)),
                   columns=['acc_p', 'acc_r', 'sharp_r', 'd', 'IR', 'beta', 'alpha'])

df1.to_excel('huice_output.xlsx')
