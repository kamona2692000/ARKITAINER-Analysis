
import math
import pandas as pd
import numpy as np


IPE = pd.read_excel (r'C:\Users\Kamona\Desktop\pythondata.xlsx', sheet_name='IPE')
print(IPE.shape)
#pd.set_option('display.max_columns',10)
IPE_avgprice = IPE['Average price']
IPEavg = IPE_avgprice.mean()
print("this is avg price",IPEavg)
IPEPivot=IPE.pivot_table(index="Weeks",aggfunc="sum",values=["purchases(kg)","sales(kg)"])
#Get stock of each week
IPEPivot['Stock']=0
len_of_table = len(IPEPivot['Stock'])
IPEStock = np.zeros(len_of_table)
IPEStock[0]=IPEPivot.iloc[0][0]-IPEPivot.iloc[0][1]
for i in range(len_of_table-1):
    IPEStock[i+1]=IPEStock[i]+IPEPivot.iloc[i+1][0]-IPEPivot.iloc[i+1][1]
IPEPivot['Stock']=IPEStock
#print(IPEPivot)
#Get Holding cost of each week
IPEHolding = np.zeros(len_of_table)
IPEHolding[0]=IPEPivot.iloc[0][0]*0.2
for i in range(len_of_table-1):
    IPEHolding[i+1]=(IPEStock[i]+IPEPivot.iloc[i+1][0])*0.2
IPEPivot['Holding']=IPEHolding
#print(IPEPivot)
#to get avg sales price
ITEM_avgsalesprice = IPE['Average price for sales']
total_price = 0
total_count = 0
for i in range(len(ITEM_avgsalesprice)):
    if ITEM_avgsalesprice[i] > 0:
        total_price += ITEM_avgsalesprice[i]
        total_count += 1
IPEsalesavg = total_price / total_count

#Get order cost of each week
IPEOrdercost = np.zeros(len_of_table)
for i in range(len_of_table):
    IPEOrdercost[i]=IPEPivot.iloc[i][0]*IPEavg*0.05
IPEPivot['Order cost']=IPEOrdercost
print(IPEPivot)
positive_purchases = []
#Get average quantity per order
for i in range(len(IPE)):
    if IPE.iloc[i][3] > 0:
        positive_purchases.append(IPE.iloc[i][3])
avgQuantityPerOrder = sum(positive_purchases)/len(positive_purchases)
#Fixed cost for placing and receiving an order
F = avgQuantityPerOrder*IPEavg*0.05
#Annual sales in units
S = sum(IPEPivot['sales(kg)'])
"Carrying cost expressed as a percentage"
C = 0.2
"Perchase price per unit"
P = IPEavg

EOQ_IPE = math.sqrt((2*F*S)/(C*P))
print("EOQ ",EOQ_IPE)

Lead_time = 4
Weakly_average_usage = S / len_of_table
Re_order_point = Lead_time * Weakly_average_usage
print("Re-order point ",Re_order_point)
print("after using EOQ.....")
print("")
print("")
order_EOQ = np.zeros(len_of_table)

order_EOQ = np.zeros(len_of_table)
# fill frist 4 index for eoq
order_EOQ[0] = EOQ_IPE
order_EOQ[1] = 0
order_EOQ[2] = 0
order_EOQ[3] = 0

stock_EOQ = np.zeros(len_of_table)
stock_EOQ[0] = order_EOQ[0] - IPEPivot.iloc[0][1]
for i in range(3):
    if stock_EOQ[i] > 0:
        stock_EOQ[i + 1] = stock_EOQ[i] + order_EOQ[i + 1] - IPEPivot.iloc[i + 1][1]
    else:
        stock_EOQ[i + 1] = order_EOQ[i + 1] - IPEPivot.iloc[i + 1][1]
# to make stock and purcheses eoq
for i in range(len_of_table - 4):
    if order_EOQ[i + 1] != EOQ_IPE and order_EOQ[i + 2] != EOQ_IPE and order_EOQ[i + 3] != EOQ_IPE and stock_EOQ[
        i] <= Re_order_point:
        order_EOQ[i + 4] = EOQ_IPE
    if stock_EOQ[i + 3] > 0:
        stock_EOQ[i + 4] = stock_EOQ[i + 3] + order_EOQ[i + 4] - IPEPivot.iloc[i + 4][1]
    if stock_EOQ[i + 3] <= 0:
        stock_EOQ[i + 4] = order_EOQ[i + 4] - IPEPivot.iloc[i + 4][1]
holding_EOQ = np.zeros(len_of_table)
holding_EOQ[0] = order_EOQ[0] * 0.2
shortage_cost = IPEsalesavg - IPEavg - 0.2 - (0.05 * IPEavg)
# to get holding and shortage cost
for i in range(len_of_table - 1):
    if stock_EOQ[i] > 0 and stock_EOQ[i + 1] > 0:
        holding_EOQ[i + 1] = (stock_EOQ[i] + order_EOQ[i + 1]) * 0.2
    elif stock_EOQ[i] <= 0 and stock_EOQ[i + 1] > 0:
        holding_EOQ[i + 1] = (order_EOQ[i + 1]) * 0.2
    elif stock_EOQ[i] > 0 and stock_EOQ[i + 1] <= 0:
        holding_EOQ[i + 1] = ((stock_EOQ[i] + order_EOQ[i + 1]) * 0.2) + (stock_EOQ[i + 1] * -shortage_cost)
    elif stock_EOQ[i] <= 0 and stock_EOQ[i + 1] <= 0:
        holding_EOQ[i + 1] = (order_EOQ[i + 1] * 0.2) + (stock_EOQ[i + 1] * -shortage_cost)
# to get order cost
order_cost_EOQ = np.zeros((len_of_table), float)
for i in range(len_of_table):
    order_cost_EOQ[i] = order_EOQ[i] * P * 0.05

total_cost_EOQ = sum(order_cost_EOQ) + sum(holding_EOQ)
total_cost = sum(IPEOrdercost) + sum(IPEHolding)
saving = total_cost - total_cost_EOQ

IPEPivot['Order EOQ'] = order_EOQ
IPEPivot['order_cost_EOQ'] = order_cost_EOQ
print("Total cost without EOQ", total_cost)
print("Total cost with EOQ", total_cost_EOQ)
print("this is saving", saving)
print(shortage_cost)
# --------------------------------------
