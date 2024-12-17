import plotly.graph_objects as go
import plotly.offline as po


# Define your data
labels_x = [
    "Tripod 3110",
    "Ring Light",
    "RGB Strip Light",
    "Ring Light with 7ft Tripod",
    "4in1 Selfie Stick Tripod",
    "K9 Microphone",
    "Tripod Stand 3366"
]

Deliver_List = [382, 53, 231, 859, 502, 19, 18]  # Example delivery data
Return_List = [132, 13, 80, 317, 200, 6, 6]      # Example return data

# Create a grouped bar chart
fig = go.Figure()

# Add delivered bar chart
fig.add_trace(go.Bar(
    x=labels_x,
    y=Deliver_List,
    name="Delivered",
    marker_color='rgb(55, 83, 109)'  # Choose color for delivered
))

# Add returned bar chart
fig.add_trace(go.Bar(
    x=labels_x,
    y=Return_List,
    name="Returns",
    marker_color='rgb(26, 118, 255)'  # Choose color for returns
))

# Update layout for the grouped bar chart
fig.update_layout(
    title="Delivered vs Returns by Product",
    xaxis_title="Products",
    yaxis_title="Count",
    barmode='group',  # 'group' puts the bars side by side
    height=600,
    width=900,
    template="plotly_white",  # Light theme for better readability
    font=dict(
        family="Arial, sans-serif",
        size=14,
        color="Black"
    )
)

# Show the plot
# fig.show()

# Save as HTML
po.plot(fig,"grouped_bar_chart.html")
