import numpy as np
import pandas as pd
import xlrd

Account=1000000
bk = xlrd.open_workbook('index.xls')

for sheet in bk.sheets():   #读每一个表格

    df=pd.read_excel('index.xls',sheet.name)
    da=np.mat(df)
    numOfSamples=da.shape[0]    #表格行数

    t=np.empty(numOfSamples,dtype=object)

    N = np.zeros(numOfSamples)#第一个值没有意义
    unit_all = np.zeros(numOfSamples)
    cost= np.zeros(numOfSamples)
    unit = np.zeros(numOfSamples) #交易单位（下标为0的没有意义）
    print (numOfSamples)
    j=0 #统计m的

    for i in range(1,20):   #i=1----19
        N[i]=max(da[i, 1] - da[i, 2], abs(da[i, 1] - da[i - 1, 4]), abs(da[i - 1, 4] - da[i, 2]))
        #真实波幅：max(high-low,abs(high-close(-1)),abs(low-close(-1)))
        unit[i] = float(0.1 * Account / N[i])
            #交易单位（股）
    for i in range(20,numOfSamples): #i=20----266
        truerange = max(da[i,1]-da[i,2],abs(da[i,1]-da[i-1,4]),abs(da[i-1,4]-da[i,2]))
        N[i]=(sum(N[i-19:i])+truerange)/20  #20天移动平均
        unit[i] = float(0.1*Account/N[i])

        if unit_all[j] != 0:    #加仓/平仓
            if da[i,4] >= cost[j]+0.5*N[i]: #cost[j]:上一笔交易的交易成本
                j=j+1
                # m[i]=m[i]+1
                # m[i]=sum(m)
                unit_all[j]=unit_all[j-1]+unit[i]   #加仓：1一个交易单位
                cost[j] = da[i,4]   #收盘的时候加仓--交易成本：收盘价
                t[j] = da[i, 0]
                # timeArray = time.strptime(da[i, 0], "'%Y%m%d'")
                # t[j] = float(time.mktime(timeArray))

                print u'%d Add %.2f'%(i,unit_all[j])

            if da[i,4]<cost[j]-2*N[i] or da[i,4]<np.min(da[i-20:i,2],axis=0):  #平仓
                j=j+1
                # m[i]=m[i]+1
                # m[i] = sum(m)
                unit[i]=-unit_all[j-1]
                unit_all[j]=unit_all[j-1]+unit[i]   
                cost[j]=da[i,4]
                t[j] = da[i, 0]

                print u'%d Clear %.2f'%(i,unit_all[j])
            else:
                print u'%d No Action %.2f'%(i,unit_all[j])

        else:   #开仓
            if da[i,4]>np.max(da[i-20:i,1],axis=0):
                #close>max(high(-20:0))
                j = j + 1
                # m[i]=m[i]+1
                cost[j]=da[i,4]
                unit_all[j]=unit[i]
                t[j] = da[i, 0]
                # timeArray = time.strptime(da[i,0], "'%Y%m%d'")
                # t[j] = float(time.mktime(timeArray))

                print '%d Open %.2f'%(i,unit_all[j])
            else:
                print '%d No action %.2f'%(i,unit_all[j])

    df1 = pd.DataFrame(np.column_stack((unit_all,cost,t)),columns=['unit_all','cost','m_date'])

    df1.to_excel('.\\turtle_out\\'+sheet.name + '_out.xls')




