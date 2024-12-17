import pandas as pd
import openpyxl as pyxl

wb = pyxl.load_workbook('leopards.xlsx')
sheet= wb.active
# Name = xl['Name'].last_valid_index()
xl = pd.read_excel('leopards.xlsx')
Leop_id = xl['ID'].tolist()
Status = xl['Status'].tolist()
Orders = xl['Name'].tolist()
write = xl['test'].tolist()
for a in Orders:
    if(a in Leop_id):
        # Getting order index
        Order_index = Orders.index(a)
        leop_id_index = Leop_id.index(a)
        # writing status in new collumn
        Deliver_status = sheet[f'A{leop_id_index+2}'].value
        W_status = sheet[f'D{Order_index+2}'].value = Deliver_status # add two becouse openpxl works not in header & start from 1
wb.save("new.xlsx")

