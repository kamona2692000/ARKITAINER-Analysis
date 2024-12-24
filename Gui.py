from tkinter import *
import tkinter as tk
from tkinter.ttk import *
from tkinter.filedialog import askopenfile
from openpyxl import load_workbook
from PIL import ImageTk, Image
import os
import math
import pandas as pd
import numpy as np

root = tk.Tk()
root.title("Arkitainer Inventory model")

root.minsize(600, 800)
root.configure(bg='white')
# Set arkitainer icon
root.iconbitmap('icon.ico')
# -----------------------------
#------------------------------
logo = ImageTk.PhotoImage(Image.open('logo.png'))
logolabel = Label(image = logo)
logolabel.grid(row=1,column=1,columnspan=3)
#Items pictures
anglespic = ImageTk.PhotoImage(Image.open('angles.png'))
upnpic = ImageTk.PhotoImage(Image.open('upn.png'))
ipepic = ImageTk.PhotoImage(Image.open('ipe.png'))
hebpic = ImageTk.PhotoImage(Image.open('heb.png'))
heapic = ImageTk.PhotoImage(Image.open('hea.png'))
item_list = [anglespic,upnpic,ipepic,hebpic,heapic]
#items lables
ipelabel = Label(image = ipepic)
angleslabel = Label(image = anglespic)
healabel = Label(image = heapic)
heblabel = Label(image = hebpic)
upnlabel = Label(image = upnpic)


#Preview items functions
def preview_IPE():
    global sheet_name
    sheet_name = "IPE"
    ipelabel.grid(row=4,column=1,columnspan=3)
    upnlabel.grid_forget()
    heblabel.grid_forget()
    healabel.grid_forget()
    angleslabel.grid_forget()
    btn2 = Button(root, text='Get EOQ', command=get_EOQ)
    btn2.grid(row=6, column=2)

    return
def preview_UPN():
    global sheet_name
    sheet_name = "UPN"
    upnlabel.grid(row=4,column=1,columnspan=3)
    heblabel.grid_forget()
    healabel.grid_forget()
    angleslabel.grid_forget()
    ipelabel.grid_forget()
    btn2 = Button(root, text='Get EOQ', command=get_EOQ)
    btn2.grid(row=6, column=2)
    return
def preview_HEB():
    global sheet_name
    sheet_name = "HEB"
    heblabel.grid(row=4,column=1,columnspan=3)
    healabel.grid_forget()
    angleslabel.grid_forget()
    ipelabel.grid_forget()
    upnlabel.grid_forget()
    btn2 = Button(root, text='Get EOQ', command=get_EOQ)
    btn2.grid(row=6, column=2)
    return
def preview_HEA():
    global sheet_name
    sheet_name = "HEA"
    healabel.grid(row=4,column=1,columnspan=3)
    angleslabel.grid_forget()
    ipelabel.grid_forget()
    upnlabel.grid_forget()
    heblabel.grid_forget()
    btn2 = Button(root, text='Get EOQ', command=get_EOQ)
    btn2.grid(row=6, column=2)
    return
def preview_ANGLES():
    global sheet_name
    sheet_name = "angles"
    angleslabel.grid(row=4,column=1,columnspan=3)
    upnlabel.grid_forget()
    heblabel.grid_forget()
    healabel.grid_forget()
    ipelabel.grid_forget()
    btn2 = Button(root, text='Get EOQ', command=get_EOQ)
    btn2.grid(row=6, column=2)
    return

st = Style()
st.configure('W.TButton', background='#345', foreground='black', font=('Arial', 12 ))

buttom_ipe = Button(root,text="IPE",command=lambda : preview_IPE(),style='W.TButton')
buttom_ipe.grid(row=3,column=0)
buttom_heb = Button(root,text="HEB",command=lambda : preview_HEB(),style='W.TButton')
buttom_hea = Button(root,text="HEA",command=lambda : preview_HEA(),style='W.TButton')
buttom_hea.grid(row=3,column=2)
buttom_upn = Button(root,text="UPN",command=lambda : preview_UPN(),style='W.TButton')
buttom_upn.grid(row=3,column=3)
buttom_angles = Button(root,text="ANGLES",command=lambda : preview_ANGLES(),style='W.TButton')
buttom_angles.grid(row=3,column=4)



def open_file():

    file = askopenfile(mode ='r', filetypes =[('Excel Files', '*.xlsx *.xlsm *.sxc *.ods *.csv *.tsv')]) # To open the file that you want.
    #' mode='r' ' is to tell the filedialog to read the file
    # 'filetypes=[()]' is to filter the files shown as only Excel files

    wb = load_workbook(filename = file.name) # Load into openpyxl
    wb2 = wb.active
    filepath = os.path.abspath(file.name)

    #Whatever you want to do with the WorkSheet
    return filepath

buttom_heb.grid(row=3,column=1)
Output="Hi"

def get_EOQ():
    ITEM = pd.read_excel(open_file(),sheet_name=sheet_name)
    # pd.set_option('display.max_columns',10)
    ITEM_avgprice = ITEM['Average price']
    ITEMavg = ITEM_avgprice.mean()
    ITEMPivot = ITEM.pivot_table(index="Weeks", aggfunc="sum", values=["purchases(kg)", "sales(kg)"])
    ITEM_avgsalesprice = ITEM['Average price for sales']
    total_price = 0
    total_count = 0
    for i in range (len(ITEM_avgsalesprice)):
        if ITEM_avgsalesprice[i]>0:
            total_price+=ITEM_avgsalesprice[i]
            total_count+=1
    ITEMsalesavg=total_price/total_count
    # Get stock of each week
    ITEMPivot['Stock'] = 0
    len_of_table = len(ITEMPivot['Stock'])
    ITEMStock = np.zeros(len_of_table)
    ITEMStock[0] = ITEMPivot.iloc[0][0] - ITEMPivot.iloc[0][1]
    for i in range(len_of_table - 1):
        ITEMStock[i + 1] = ITEMStock[i] + ITEMPivot.iloc[i + 1][0] - ITEMPivot.iloc[i + 1][1]
    ITEMPivot['Stock'] = ITEMStock
    # Get Holding cost of each week
    ITEMHolding = np.zeros(len_of_table)
    ITEMHolding[0] = ITEMPivot.iloc[0][0] * 0.2
    for i in range(len_of_table - 1):
        ITEMHolding[i + 1] = (ITEMStock[i] + ITEMPivot.iloc[i + 1][0]) * 0.2
    ITEMPivot['Holding'] = ITEMHolding
    # Get order cost of each week
    ITEMOrdercost = np.zeros(len_of_table)
    for i in range(len_of_table):
        ITEMOrdercost[i] = ITEMPivot.iloc[i][0] * ITEMavg * 0.05
    ITEMPivot['Order cost'] = ITEMOrdercost
    positive_purchases = []
    # Get average quantity per order
    for i in range(len(ITEM)):
        if ITEM.iloc[i][3] > 0:
            positive_purchases.append(ITEM.iloc[i][3])
    avgQuantityPerOrder = sum(positive_purchases) / len(positive_purchases)
    # Fixed cost for placing and receiving an order
    F = avgQuantityPerOrder * ITEMavg * 0.05
    # Annual sales in units
    S = sum(ITEMPivot['sales(kg)'])
    "Carrying cost expressed as a percentage"
    C = 0.2
    "Perchase price per unit"
    P = ITEMavg
    EOQ_ITEM = math.sqrt((2 * F * S) / (C * P))
    Lead_time = 4
    Weakly_average_usage = S / len_of_table
    Re_order_point = Lead_time * Weakly_average_usage
#--------------------------------------------------------
    order_EOQ = np.zeros(len_of_table)
    # fill frist 4 index for eoq
    order_EOQ[0] = EOQ_ITEM
    order_EOQ[1] = 0
    order_EOQ[2] = 0
    order_EOQ[3] = 0

    stock_EOQ = np.zeros(len_of_table)
    stock_EOQ[0] = order_EOQ[0] - ITEMPivot.iloc[0][1]
    for i in range(3):
        if stock_EOQ[i] > 0:
            stock_EOQ[i + 1] = stock_EOQ[i] + order_EOQ[i + 1] - ITEMPivot.iloc[i + 1][1]
        else:
            stock_EOQ[i + 1] = order_EOQ[i + 1] - ITEMPivot.iloc[i + 1][1]
    # to make stock and purcheses eoq
    for i in range(len_of_table - 4):
        if order_EOQ[i + 1] != EOQ_ITEM and order_EOQ[i + 2] != EOQ_ITEM and order_EOQ[i + 3] != EOQ_ITEM and stock_EOQ[
            i] <= Re_order_point:
            order_EOQ[i + 4] = EOQ_ITEM
        if stock_EOQ[i + 3] > 0:
            stock_EOQ[i + 4] = stock_EOQ[i + 3] + order_EOQ[i + 4] - ITEMPivot.iloc[i + 4][1]
        if stock_EOQ[i + 3] <= 0:
            stock_EOQ[i + 4] = order_EOQ[i + 4] - ITEMPivot.iloc[i + 4][1]
    holding_EOQ = np.zeros(len_of_table)
    holding_EOQ[0] = order_EOQ[0] * 0.2
    shortage_cost = ITEMsalesavg - ITEMavg - 0.2 - (0.05*ITEMavg)
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
    total_cost = sum(ITEMOrdercost) + sum(ITEMHolding)
    saving = total_cost - total_cost_EOQ

    ITEMPivot['Order EOQ'] = order_EOQ
    ITEMPivot['order_cost_EOQ'] = order_cost_EOQ
    print("Total cost without EOQ", total_cost)
    print("Total cost with EOQ", total_cost_EOQ)
    print("this is saving", saving)
    print(shortage_cost)
    #--------------------------------------
    Output1 = "EOQ:"
    Output2 = EOQ_ITEM
    Output2 = "{:,.2f}".format(Output2)
    Output3 = "Re-order point:"
    Output4 = Re_order_point
    Output4 = "{:,.2f}".format(Output4)
    Output5 = "Cost without EOQ"
    Output6 = total_cost
    Output6 = "{:,.2f}".format(Output6)
    Output7 = "Cost with EOQ"
    Output8 = total_cost_EOQ
    Output8 = "{:,.2f}".format(Output8)
    Output9 = "Saving"
    Output10 = total_cost - total_cost_EOQ
    Output10 = "{:,.2f}".format(Output10)
    output1 = Label(root,text=Output1,font=('Arial', 12 ))
    output1.grid(row=7,column=0)
    output2 = Label(root,text=Output2,font=('Arial', 12 ))
    output2.grid(row=7,column=1)
    output3 = Label(root,text=Output3,font=('Arial', 12 ))
    output3.grid(row=7,column=3)
    output4 = Label(root,text=Output4,font=('Arial', 12 ))
    output4.grid(row=7,column=4)

    output5 = Label(root,text=Output5,font=('Arial', 12 ))
    output5.grid(row=8,column=0)
    output6 = Label(root,text=Output6,font=('Arial', 12 ))
    output6.grid(row=8,column=1)
    output7 = Label(root,text=Output7,font=('Arial', 12 ))
    output7.grid(row=8,column=3)
    output8 = Label(root,text=Output8,font=('Arial', 12 ))
    output8.grid(row=8,column=4)
    output9 = Label(root, text=Output9, font=('Arial', 12))
    output9.grid(row=9, column=2)
    output10 = Label(root, text=Output10, font=('Arial', 12))
    output10.grid(row=10, column=2)


root.mainloop()


