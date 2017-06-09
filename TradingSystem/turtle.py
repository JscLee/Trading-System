import numpy as np
import pandas as pd
import xlrd

amount_=1000000
surplus_=amount_

bk = xlrd.open_workbook('index.xls')
fee_ratio=0.003

for sheet in bk.sheets():   #读每一个表格
    df=pd.read_excel('index.xls',sheet.name)
    da=np.mat(df)
    numOfSamples=da.shape[0]    #表格行数

    N = np.zeros(numOfSamples)#下标为0不使用
    unit_all = np.zeros(numOfSamples)
    cost= np.zeros(numOfSamples)
    unit = np.zeros(numOfSamples) #交易单位
    delta_unit = np.zeros(numOfSamples) #当日交易量
    trade_flag = np.zeros(numOfSamples) #当日交易量
    surplus= np.zeros(numOfSamples) #空闲资金
    amount= np.zeros(numOfSamples) #账户总资金
    print (numOfSamples)
    j=0 #记录最近交易下标

    amount[19]=amount_
    surplus[19]=amount_

    for i in range(1,20):   #i=1----19
        N[i]=max(da[i, 1] - da[i, 2], abs(da[i, 1] - da[i - 1, 4]), abs(da[i - 1, 4] - da[i, 2]))
        #真实波幅：max(high-low,abs(high-close(-1)),abs(low-close(-1)))
        unit[i] = int(float(0.01 * amount_ / N[i])/100)*100  #交易单位（股）
        
    for i in range(20,numOfSamples): #i=20----266
        truerange = max(da[i,1]-da[i,2],abs(da[i,1]-da[i-1,4]),abs(da[i-1,4]-da[i,2]))
        N[i]=(sum(N[i-19:i])+truerange)/20  #20天移动平均
        unit[i] = int(float(0.01 * amount_ / N[i])/100)*100

        if unit_all[j] != 0:    #加仓/平仓
            if (da[i,4] >= cost[j]+0.5*N[i]) and (surplus[i-1]>0): #加仓
                delta_unit[i]=unit[i]
                unit_all[i]=unit_all[j]+delta_unit[i]   #加仓：1一个交易单位
                cost[i] = da[i,4]   #收盘的时候加仓--交易成本：收盘价
                
                j=i
                trade_flag[i]=1
                
                surplus[i]=surplus[i-1]-delta_unit[i]*cost[i]-fee_ratio*abs(delta_unit[i]*cost[i])
                amount[i]=surplus[i]+unit_all[i]*cost[i]        
    
                print u'%d Add %.2f'%(i,unit_all[j])
            else:
                if da[i,4]<cost[j]-2*N[i] or da[i,4]<np.min(da[i-20:i,2],axis=0):  #平仓
                    delta_unit[i]=-unit_all[j]
                    unit_all[i]=unit_all[j]+delta_unit[i]   
                    cost[i]=da[i,4]

                    j=i
                    trade_flag[i]=1

                    #从总额中除去交易费率
                    surplus[i]=surplus[i-1]-delta_unit[i]*cost[i]-fee_ratio*abs(delta_unit[i]*cost[i])
                    amount[i]=surplus[i]+unit_all[i]*cost[i]       
                    print u'%d Clear %.2f'%(i,unit_all[j])
                else:
                    unit_all[i]=unit_all[j]
                    surplus[i]=surplus[i-1]
                    amount[i]=amount[i-1]
                    print u'%d No Action %.2f'%(i,unit_all[i])

        else:   #开仓
            if (da[i,4]>np.max(da[i-20:i,1],axis=0)) and (surplus[i-1]>0):
                delta_unit[i]=unit[i]   
                cost[i]=da[i,4]
                unit_all[i]=delta_unit[i]

                j=i
                trade_flag[i]=1

                #从总额中除去交易费率
                surplus[i]=surplus[i-1]-delta_unit[i]*cost[i]-fee_ratio*abs(delta_unit[i]*cost[i])
                amount[i]=surplus[i]+unit_all[i]*cost[i] 
      
                print '%d Open %.2f'%(i,unit_all[j])
            else:
                unit_all[i]=unit_all[j]

                surplus[i]=surplus[i-1]
                amount[i]=amount[i-1]
                print '%d No action %.2f'%(i,unit_all[i])
    
    surplus_=surplus[i]
    amount_=amount[i]

    df1 = pd.DataFrame(np.column_stack((N[20:],unit[20:],unit_all[20:],cost[20:],da[20:,0],trade_flag[20:],delta_unit[20:],surplus[20:],amount[20:])),columns=['N','unit','unit_all','cost','date','trade','delta','surplus','amount'])

    df1.to_excel('.\\turtle_out\\'+sheet.name + '_out.xls')



