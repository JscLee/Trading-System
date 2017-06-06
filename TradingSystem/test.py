import numpy as np
import pandas as pd
import xlrd

Account=1000000
n=20
bk = xlrd.open_workbook('index.xlsx')

for sheet in bk.sheets():

    df=pd.read_excel('index.xlsx',sheet.name)
    da=np.mat(df)
    numOfSamples=da.shape[0]

    t=np.empty(numOfSamples,dtype=object)

    N = np.zeros(numOfSamples)#第一个值没有意义
    unit_all = np.zeros(numOfSamples)
    cost= np.zeros(numOfSamples)
    unit = np.zeros(numOfSamples) #第一个值没有意义
    print (numOfSamples)
    j=0 #统计m的

    for i in range(1,n):
        N[i]=max(da[i, 1] - da[i, 2], abs(da[i, 1] - da[i - 1, 4]), abs(da[i - 1, 4] - da[i, 2]))
        # print(i) #i=1----19
        unit[i] = float(0.1 * Account / N[i])

    for i in range(n,numOfSamples): #i=20----266
        truerange = max(da[i,1]-da[i,2],abs(da[i,1]-da[i-1,4]),abs(da[i-1,4]-da[i,2]))
        N[i]=(sum(N[i-19:i])+truerange)/20  #dont change
        unit[i] = float(0.1*Account/N[i])

        if unit_all[j] != 0:
            if da[i,4] >= cost[j]+0.5*N[i]:
                j=j+1
                # m[i]=m[i]+1
                # m[i]=sum(m)
                unit_all[j]=unit_all[j-1]+unit[i]
                cost[j] = da[i,4]
                t[j] = da[i, 0]
                # timeArray = time.strptime(da[i, 0], "'%Y%m%d'")
                # t[j] = float(time.mktime(timeArray))

                print(i, '0.5N',unit_all[j])

            if da[i,4]<cost[j]-2*N[i]:
                j=j+1
                # m[i]=m[i]+1
                # m[i] = sum(m)
                unit[i]=-unit_all[j-1]
                cost[j]=-da[i,4]
                t[j] = da[i, 0]

                print(i, '2N',unit_all[j])
            else:
                if da[i,4]<np.min(da[i-n:i,2],axis=0):
                    j = j + 1
                    # m[i]=m[i]+1
                    # m[i] = sum(m)
                    # unit[i]=-unit_all[j-1]
                    unit_all[j] = 0
                    cost[j]=-da[i,4]
                    t[j] = da[i, 0]
                    # timeArray = time.strptime(da[i, 0], "'%Y%m%d'")
                    # t[j] = float(time.mktime(timeArray))
                    print(i, 'min',unit_all[j])
                else:
                    print(i, '!min', unit_all[j])

        else:
            if da[i,4]>np.max(da[i-n:i,1],axis=0): # not sub 1
                j = j + 1
                # m[i]=m[i]+1
                cost[j]=da[i,4]
                unit_all[j]=unit[i]
                t[j] = da[i, 0].__str__()

                # timeArray = time.strptime(da[i,0], "'%Y%m%d'")
                # t[j] = float(time.mktime(timeArray))

                print(i, 'max',unit_all[j])
            else:
                print(i,'!max',unit_all[j])

    df1 = pd.DataFrame(np.column_stack((unit,unit_all,cost,t)),columns=['unit','unit_all','cost','m_date'])

    df1.to_excel(sheet.name + '_out.xlsx')




