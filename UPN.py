
import math
import pandas as pd
import numpy as np


UPN = pd.read_excel (r'C:\Users\Kamona\Desktop\pythondata.xlsx', sheet_name='UPN')
print(UPN.shape)
#pd.set_option('display.max_columns',10)
UPN_avgprice = UPN['Average price']
UPNavg = UPN_avgprice.mean()
print("this is avg price",UPNavg)
UPNPivot=UPN.pivot_table(index="Weeks",aggfunc="sum",values=["purchases(kg)","sales(kg)"])
#Get stock of each week
UPNPivot['Stock']=0
len_of_table = len(UPNPivot['Stock'])
UPNStock = np.zeros(len_of_table)
UPNStock[0]=UPNPivot.iloc[0][0]-UPNPivot.iloc[0][1]
for i in range(len_of_table-1):
    UPNStock[i+1]=UPNStock[i]+UPNPivot.iloc[i+1][0]-UPNPivot.iloc[i+1][1]
UPNPivot['Stock']=UPNStock
#print(UPNPivot)
#Get Holding cost of each week
UPNHolding = np.zeros(len_of_table)
UPNHolding[0]=UPNPivot.iloc[0][0]*0.2
for i in range(len_of_table-1):
    UPNHolding[i+1]=(UPNStock[i]+UPNPivot.iloc[i+1][0])*0.2
UPNPivot['Holding']=UPNHolding
#print(UPNPivot)
#to get avg sales price
ITEM_avgsalesprice = UPN['Average price for sales']
total_price = 0
total_count = 0
for i in range(len(ITEM_avgsalesprice)):
    if ITEM_avgsalesprice[i] > 0:
        total_price += ITEM_avgsalesprice[i]
        total_count += 1
UPNsalesavg = total_price / total_count

#Get order cost of each week
UPNOrdercost = np.zeros(len_of_table)
for i in range(len_of_table):
    UPNOrdercost[i]=UPNPivot.iloc[i][0]*UPNavg*0.05
UPNPivot['Order cost']=UPNOrdercost
print(UPNPivot)
positive_purchases = []
#Get average quantity per order
for i in range(len(UPN)):
    if UPN.iloc[i][3] > 0:
        positive_purchases.append(UPN.iloc[i][3])
avgQuantityPerOrder = sum(positive_purchases)/len(positive_purchases)
#Fixed cost for placing and receiving an order
F = avgQuantityPerOrder*UPNavg*0.05
#Annual sales in units
S = sum(UPNPivot['sales(kg)'])
"Carrying cost expressed as a percentage"
C = 0.2
"Perchase price per unit"
P = UPNavg

EOQ_UPN = math.sqrt((2*F*S)/(C*P))
print("EOQ ",EOQ_UPN)

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
order_EOQ[0] = EOQ_UPN
order_EOQ[1] = 0
order_EOQ[2] = 0
order_EOQ[3] = 0

stock_EOQ = np.zeros(len_of_table)
stock_EOQ[0] = order_EOQ[0] - UPNPivot.iloc[0][1]
for i in range(3):
    if stock_EOQ[i] > 0:
        stock_EOQ[i + 1] = stock_EOQ[i] + order_EOQ[i + 1] - UPNPivot.iloc[i + 1][1]
    else:
        stock_EOQ[i + 1] = order_EOQ[i + 1] - UPNPivot.iloc[i + 1][1]
# to make stock and purcheses eoq
for i in range(len_of_table - 4):
    if order_EOQ[i + 1] != EOQ_UPN and order_EOQ[i + 2] != EOQ_UPN and order_EOQ[i + 3] != EOQ_UPN and stock_EOQ[
        i] <= Re_order_point:
        order_EOQ[i + 4] = EOQ_UPN
    if stock_EOQ[i + 3] > 0:
        stock_EOQ[i + 4] = stock_EOQ[i + 3] + order_EOQ[i + 4] - UPNPivot.iloc[i + 4][1]
    if stock_EOQ[i + 3] <= 0:
        stock_EOQ[i + 4] = order_EOQ[i + 4] - UPNPivot.iloc[i + 4][1]
holding_EOQ = np.zeros(len_of_table)
holding_EOQ[0] = order_EOQ[0] * 0.2
shortage_cost = UPNsalesavg - UPNavg - 0.2 - (0.05 * UPNavg)
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
total_cost = sum(UPNOrdercost) + sum(UPNHolding)
saving = total_cost - total_cost_EOQ

UPNPivot['Order EOQ'] = order_EOQ
UPNPivot['order_cost_EOQ'] = order_cost_EOQ
print("Total cost without EOQ", total_cost)
print("Total cost with EOQ", total_cost_EOQ)
print("this is saving", saving)
print(shortage_cost)
# --------------------------------------