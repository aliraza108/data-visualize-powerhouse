import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl as xl
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.offline as pyo
import plotly.graph_objects as go
import plotly.offline as po



# adding file in pandas
file = pd.read_excel("data.xlsx")
Status = file['Status']

product = file["Products"].tolist()

# opeing excel in pyexcel
wb = xl.load_workbook(r'data.xlsx')
sheet = wb.active


labels_x = [
"Tripod 3110",
"Ring Light",
"RGB Strip Light",
"Ring Light with 7ft Tripod",
"4in1 Selfie Stick Tripod",
"K9 Microphone",
"Tripod Stand 3366"
]

total =1
total_R =1
delivered_total =0
Return_Total =0
Deliver_List=[]
Return_List=[]

for item in labels_x:
# Getting Total Number of Delivered Products
    for l in Status:
            Cell_product = sheet[f'C{total}'].value
            if(l=='Delivered'):
                if(Cell_product==item):
                    delivered_total = delivered_total+1
            total=total+1
    Deliver_List.append(delivered_total)
    delivered_total = 0
    total =1
    # Getting Total Number of Return Products
    for st in Status:
                Cell_product = sheet[f'C{total_R}'].value
                if(st=='Returned to shipper'):
                    if(Cell_product ==item):
                        Return_Total = Return_Total+1
                total_R=total_R+1
    Return_List.append(Return_Total)
    Return_Total=0
    total_R=1
print(Deliver_List)
print(Return_List)

#add retrn besides the deliver
plot_deliver = px.bar(y=Deliver_List, x=labels_x, color=labels_x, title="delivered")





# Getting highest number
p_hiest_list = [
"Tripod 3110",
"Ring Light",
"RGB Strip Light",
"Ring Light with 7ft Tripod",
"4in1 Selfie Stick Tripod"
]

tc_list = []

for a in p_hiest_list:
    tc = 1
    hight_count = 0
    for i in Status:  # Assuming 'Status' is a list of delivery statuses
        Cell_product = sheet[f'C{tc}'].value  # Assuming 'sheet' is the worksheet
        if Cell_product == a:
            if i == 'Delivered':
                hight_count += 1
        tc += 1

    tc_list.append(hight_count)

# Get the highest order number and the corresponding product
max_count = max(tc_list)  # Get the highest count
max_index = tc_list.index(max_count)  # Find the index of the highest count
max_product = p_hiest_list[max_index]  # Get the corresponding product
print(max_count)
print(max_product)



tc_list_return = []

for a in p_hiest_list:
    tc = 1
    hight_count = 0
    for i in Status:  # Assuming 'Status' is a list of delivery statuses
        Cell_product = sheet[f'C{tc}'].value  # Assuming 'sheet' is the worksheet
        if Cell_product == a:
            if i == 'Returned to shipper':
                hight_count += 1
        tc += 1

    tc_list_return.append(hight_count)

# Get the highest order number and the corresponding product
print(tc_list_return)
max_count_return = max(tc_list_return)  # Get the highest count
max_index_return = tc_list_return.index(max_count_return)  # Find the index of the highest count
print(max_index_return)
max_product_return = p_hiest_list[max_index_return]  # Get the corresponding product

print(max_count)
print(max_product)
print(max_count_return)
print(max_product_return)

# fig_pie = px.pie(names=labels_x, values=Deliver_List, color=Deliver_List, hole=0.1, title='Delivered ')












# Pie Chart Gettin Total Returns 
# Count Return
return_Status_count = 0
for i in Status:
    if(i=="Returned to shipper"):
        return_Status_count=return_Status_count+1


# Count Pending
pending_Status_count = 0
for i in Status:
    if(i=="Pending"):
        pending_Status_count=pending_Status_count+1


# Count Pickup
pickupRQ_Status_count = 0
for i in Status:
    if(i=="Pickup Request Sent"):
        pickupRQ_Status_count=pickupRQ_Status_count+1


# Count Delivered
Delivered_Status_count = 0
for i in Status:
    if(i=="Delivered"):
        Delivered_Status_count=Delivered_Status_count+1

# Count Cancelled
Cancelled_Status_count = 0
for i in Status:
    if(i=="Cancelled"):
        Cancelled_Status_count=Cancelled_Status_count+1

# Count Pickup Request not Send
NotPickUpRQ_Status_count = 0
for i in Status:
    if(i=="Pickup Request not Send"):
        NotPickUpRQ_Status_count=NotPickUpRQ_Status_count+1

# Count Being Return
BeingRT_Status_count = 0
for i in Status:
    if(i=="Being Return"):
        BeingRT_Status_count=BeingRT_Status_count+1


Delivered_data = [Delivered_Status_count,Cancelled_Status_count,return_Status_count,pickupRQ_Status_count,NotPickUpRQ_Status_count,pickupRQ_Status_count]

label =["Deliver","Cencel","Return","Pickup","NotPickUpRQ_Status_count","pickupRQ_Status_count"]












## converting data to data frame with pandas
# data = pd.DataFrame(Deliver_List,Return_List)


plot_report = px.pie(names=label, values=Delivered_data, color=label, title="Pie Delivered")
# plot_deliver = px.line(data, y=Deliver_List, x=labels_x, title="Delivered")
# # plot_return = px.bar(y=[20, 30, 40, 50], x=labels_x, color=labels_x, title="Return")


fig = make_subplots(
    rows=1, cols=2,  
    specs=[[{"type": "bar"}, {"type": "pie"}]], 
    subplot_titles=("Bar Chart (Delivered & Returned)", "Pie Chart")
)

# fig =  go.Figure()
# fig =  go.Figure(data=[go.pie(labels=labels_x,values=Delivered_data)])

# Add delivered bar chart
fig.add_trace(go.Bar(
    x=labels_x,
    y=Deliver_List,
    name="Delivered",
    marker_color='rgb(55, 83, 109)'  # Choose color for delivered
),row=1, col=1)

# Add returned bar chart
# fig.add_trace(go.Bar(
#     x=labels_x,
#     y=Return_List,
#     name="Returns",
#     marker_color='rgb(26, 118, 255)'  # Choose color for returns
# ),row=1,col=1)


fig.add_trace(plot_deliver,row=1,col=2)

fig.update_layout(
    barmode='group',  # Group the bars side by side
    title_text="Dashboard with Grouped Bar Charts and Other Plots"
)
# Save the Dashboard as an HTML File
po.plot(fig, filename="test.html")

