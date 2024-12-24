
import math
import pandas as pd
import numpy as np


HEB = pd.read_excel (r'C:\Users\Kamona\Desktop\pythondata.xlsx', sheet_name='HEB')
print(HEB.shape)
#pd.set_option('display.max_columns',10)
HEB_avgprice = HEB['Average price']
HEBavg = HEB_avgprice.mean()
print("this is avg price",HEBavg)
HEBPivot=HEB.pivot_table(index="Weeks",aggfunc="sum",values=["purchases(kg)","sales(kg)"])
#Get stock of each week
HEBPivot['Stock']=0
len_of_table = len(HEBPivot['Stock'])
HEBStock = np.zeros(len_of_table)
HEBStock[0]=HEBPivot.iloc[0][0]-HEBPivot.iloc[0][1]
for i in range(len_of_table-1):
    HEBStock[i+1]=HEBStock[i]+HEBPivot.iloc[i+1][0]-HEBPivot.iloc[i+1][1]
HEBPivot['Stock']=HEBStock
#print(HEBPivot)
#Get Holding cost of each week
HEBHolding = np.zeros(len_of_table)
HEBHolding[0]=HEBPivot.iloc[0][0]*0.2
for i in range(len_of_table-1):
    HEBHolding[i+1]=(HEBStock[i]+HEBPivot.iloc[i+1][0])*0.2
HEBPivot['Holding']=HEBHolding
#print(HEBPivot)
#to get avg sales price
ITEM_avgsalesprice = HEB['Average price for sales']
total_price = 0
total_count = 0
for i in range(len(ITEM_avgsalesprice)):
    if ITEM_avgsalesprice[i] > 0:
        total_price += ITEM_avgsalesprice[i]
        total_count += 1
HEBsalesavg = total_price / total_count

#Get order cost of each week
HEBOrdercost = np.zeros(len_of_table)
for i in range(len_of_table):
    HEBOrdercost[i]=HEBPivot.iloc[i][0]*HEBavg*0.05
HEBPivot['Order cost']=HEBOrdercost
print(HEBPivot)
positive_purchases = []
#Get average quantity per order
for i in range(len(HEB)):
    if HEB.iloc[i][3] > 0:
        positive_purchases.append(HEB.iloc[i][3])
avgQuantityPerOrder = sum(positive_purchases)/len(positive_purchases)
#Fixed cost for placing and receiving an order
F = avgQuantityPerOrder*HEBavg*0.05
#Annual sales in units
S = sum(HEBPivot['sales(kg)'])
"Carrying cost expressed as a percentage"
C = 0.2
"Perchase price per unit"
P = HEBavg

EOQ_HEB = math.sqrt((2*F*S)/(C*P))
print("EOQ ",EOQ_HEB)

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
order_EOQ[0] = EOQ_HEB
order_EOQ[1] = 0
order_EOQ[2] = 0
order_EOQ[3] = 0

stock_EOQ = np.zeros(len_of_table)
stock_EOQ[0] = order_EOQ[0] - HEBPivot.iloc[0][1]
for i in range(3):
    if stock_EOQ[i] > 0:
        stock_EOQ[i + 1] = stock_EOQ[i] + order_EOQ[i + 1] - HEBPivot.iloc[i + 1][1]
    else:
        stock_EOQ[i + 1] = order_EOQ[i + 1] - HEBPivot.iloc[i + 1][1]
# to make stock and purcheses eoq
for i in range(len_of_table - 4):
    if order_EOQ[i + 1] != EOQ_HEB and order_EOQ[i + 2] != EOQ_HEB and order_EOQ[i + 3] != EOQ_HEB and stock_EOQ[
        i] <= Re_order_point:
        order_EOQ[i + 4] = EOQ_HEB
    if stock_EOQ[i + 3] > 0:
        stock_EOQ[i + 4] = stock_EOQ[i + 3] + order_EOQ[i + 4] - HEBPivot.iloc[i + 4][1]
    if stock_EOQ[i + 3] <= 0:
        stock_EOQ[i + 4] = order_EOQ[i + 4] - HEBPivot.iloc[i + 4][1]
holding_EOQ = np.zeros(len_of_table)
holding_EOQ[0] = order_EOQ[0] * 0.2
shortage_cost = HEBsalesavg - HEBavg - 0.2 - (0.05 * HEBavg)
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
total_cost = sum(HEBOrdercost) + sum(HEBHolding)
saving = total_cost - total_cost_EOQ

HEBPivot['Order EOQ'] = order_EOQ
HEBPivot['order_cost_EOQ'] = order_cost_EOQ
print("Total cost without EOQ", total_cost)
print("Total cost with EOQ", total_cost_EOQ)
print("this is saving", saving)
print(shortage_cost)
# --------------------------------------
