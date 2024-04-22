import streamlit as st
import pandas as pd
import plotly.express as px
import webbrowser as wb
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io
import calendar
from datetime import datetime

today_date = datetime.now().strftime('%Y-%m-%d')

st.set_page_config(page_title="Dashboard", page_icon=":bar_chart:", layout="wide")

#st.header("File Upload")
data_file1 = st.file_uploader(".xlsx file1",type=['xlsx'])
data_file2 = st.file_uploader(".xlsx file2",type=['xlsx'])
data_file3 = st.file_uploader(".xlsx file3",type=['xlsx'])
st.markdown("#")

st.header(f":bar_chart: Dashboard")

def process_data_sheets(data_file):
    # Helper function to process individual sheets
    def process_sheet(df):
        df.drop(df.index[:4], inplace=True)
        df.columns = df.iloc[0]
        df = df[1:]
        df.reset_index(drop=True, inplace=True)
        return df

    # Read the sheets into DataFrames
    df1 = pd.read_excel(data_file, sheet_name="Web")
    df2 = pd.read_excel(data_file, sheet_name="Shopee")
    df3 = pd.read_excel(data_file, sheet_name="Lazada")
    df4 = pd.read_excel(data_file, sheet_name="TikTok")

    # Process each DataFrame
    df1 = process_sheet(df1)
    df2 = process_sheet(df2)
    df3 = process_sheet(df3)
    df4 = process_sheet(df4)

    # Concatenate all DataFrames
    df = pd.concat([df1, df2, df3, df4])
    return df


df_A = process_data_sheets(data_file1)
df_B = process_data_sheets(data_file2)
df_C = process_data_sheets(data_file3)


df_ALL = pd.concat([df_A,df_B,df_C])

# Convert 'Date Added' to datetime format
df_ALL['Date Added'] = pd.to_datetime(df_ALL['Date Added'], errors='coerce')

# Check for any NaT values which indicate conversion failures
print(df_ALL[df_ALL['Date Added'].isna()])

# Assuming the conversion was successful, extract the month
df_ALL['month'] = df_ALL['Date Added'].dt.month
df_ALL['month'].ffill(inplace=True)
df_ALL['month'] = df_ALL['month'].apply(lambda x: calendar.month_name[int(x)])
#df_ALL
df_ALL['Order Source'].ffill(inplace=True)
df_ALL['Order Status'].ffill(inplace=True)

df_ALL_copy=df_ALL.copy()
############################################################################################################################################
st.sidebar.write("Filter Here")
# Start a form in the sidebar
with st.sidebar.form(key='filter_form'):
    # Get a list of unique values for each category
    months = df_ALL_copy['month'].unique().tolist()
    marketplaces = df_ALL_copy['Order Source'].unique().tolist()
    status = df_ALL_copy['Order Status'].unique().tolist()

    # Create multiselect widgets inside the form for each category
    selected_month = st.selectbox('Month:', options=months)
    selected_marketplaces = st.multiselect('Marketplace:', marketplaces)
    selected_status = st.multiselect('Status:', status)

    # Dictionary to map month name to its number
    month_to_number = {name: num for num, name in enumerate(calendar.month_name) if name}

    # Your current month name variable
    current_month_name = selected_month # Example

    # Get the current month number
    current_month_number = month_to_number[current_month_name]

    # Calculate the previous month number
    previous_month_number = current_month_number - 1 if current_month_number > 1 else 12

    # Map the previous month number to its name
    number_to_month = {num: name for name, num in month_to_number.items()}
    previous_month_name = number_to_month[previous_month_number]

    state_option= st.selectbox('Focus:', options=['Sales', 'Profit', 'Units'])

    # Form submit button
    submitted = st.form_submit_button('Apply')

# Check if the form has been submitted
if submitted:
    # Filter the DataFrame based on the selected values
    filtered_df = df_ALL_copy[df_ALL_copy['month'].isin([selected_month])]
    filtered_df = filtered_df[filtered_df['Order Source'].isin(selected_marketplaces)]
    filtered_df = filtered_df[filtered_df['Order Status'].isin(selected_status)]

    df=filtered_df

    df_prev = df_ALL_copy[df_ALL_copy['month'].isin([previous_month_name])]
    df_prev = df_prev[df_prev['Order Source'].isin(selected_marketplaces)]
    df_prev = df_prev[df_prev['Order Status'].isin(selected_status)]

else:
    df=df_ALL
    df_prev=df_ALL
    # Display the filtered DataFrame
    #st.write(filtered_df)

##############################################################################################################################################################
# Sample DataFrame with 'state', 'latitude', and 'longitude' for locations in Malaysia
states_coordinates = pd.DataFrame({
    'Shipping State': ['Johor', 'Kedah', 'Kelantan', 'Malacca', 'Negeri Sembilan', 'Pahang', 'Penang', 'Perak', 'Perlis', 'Sabah', 'Sarawak', 'Selangor', 'Terengganu'],
    'lat': [1.4854, 6.1184, 6.1254, 2.1896, 2.7258, 3.8126, 5.4164, 4.5921, 6.4449, 5.9804, 1.5533, 3.0738, 5.3117],
    'lon': [103.7618, 100.3682, 102.2386, 102.2501, 101.9424, 103.3256, 100.3308, 101.0901, 100.2048, 116.0753, 110.3441, 101.5183, 103.1324]
})

df.loc[(df['Shipping State'] == "Kuala Lumpur") | (df['Shipping State'] == "W.P. Kuala Lumpur") | (df['Shipping State'] == "Wp Kuala Lumpur") | (df['Shipping State'] == "Putrajaya") | (df['Shipping State'] == "W.P. Putrajaya") | (df['Shipping State'] == "Wp Putrajaya") | (df['Shipping State'] == "Wilayah Persekutuan Putrajaya"), 'Shipping State'] = "Selangor"
df.loc[(df['Shipping State'] == "Labuan") | (df['Shipping State'] == "W.P. Labuan") | (df['Shipping State'] == "Wp Labuan"), 'Shipping State'] = "Sabah"
df.loc[(df['Shipping State'] == "Melaka"), 'Shipping State'] = "Malacca"
df_state=df['Shipping State'].ffill(inplace=True)
df_state = df.merge(states_coordinates, on='Shipping State', how='left')
#df_state

# Assuming 'df' is your DataFrame with 'Shipping State', 'lat', 'lon', 'product', and 'Unit Total' columns

# First, group your data by 'Shipping State', 'lat', and 'lon', and sum the 'Unit Total'
df_state['Unit Total'] = pd.to_numeric(df_state['Unit Total'], errors='coerce')
df_state['Margin Per Item'] = pd.to_numeric(df_state['Margin Per Item'], errors='coerce')
df_state['Quantity'] = pd.to_numeric(df_state['Quantity'], errors='coerce')
df_state_sum = df_state.groupby(['Shipping State', 'lat', 'lon'], as_index=False).agg({
    'Unit Total': 'sum',
    'Margin Per Item': 'sum',
    'Quantity': 'sum'
})
df_state_sum.rename(columns={
    'Unit Total': 'Sales',
    'Margin Per Item': 'Profit',
    'Quantity': 'Units'
}, inplace=True)
#df_state_sum

##############################################################################################################################################
df['Unit Total'] = df['Unit Total'].replace('-', 0)
total_sales = df['Unit Total'].sum()
total_sales = int(total_sales)
df_prev['Unit Total'] = df_prev['Unit Total'].replace('-', 0)
total_sales_prev = df_prev['Unit Total'].sum()
total_sales_prev = int(total_sales_prev)
compare_sales = total_sales - total_sales_prev

df['Margin Per Item'] = df['Margin Per Item'].replace('-', 0)
total_profit = df['Margin Per Item'].sum()
total_profit = int(total_profit)
df_prev['Margin Per Item'] = df_prev['Margin Per Item'].replace('-', 0)
total_profit_prev = df_prev['Margin Per Item'].sum()
total_profit_prev = int(total_profit_prev)
compare_profit = total_profit - total_profit_prev

total_qty = df['Quantity'].sum()
total_qty_prev = df_prev['Quantity'].sum()
compare_total_qty = total_qty - total_qty_prev

df['Order ID'].ffill(inplace=True)
average_margin = df.groupby('Order ID')['Margin Per Item'].sum().reset_index()
average_margin = round(average_margin['Margin Per Item'].mean(),2)
df_prev['Order ID'].ffill(inplace=True)
average_margin_prev = df_prev.groupby('Order ID')['Margin Per Item'].sum().reset_index()
compare_average_margin = average_margin - average_margin_prev
compare_average_margin = round(compare_average_margin['Margin Per Item'].mean(),2)

profit_per = int(total_profit/total_sales *100)
if total_sales_prev is not None and total_sales_prev != 0:
    profit_per_prev = int(total_profit_prev/total_sales_prev *100)
    compare_profit_per = profit_per - profit_per_prev
else:
    compare_profit_per = 0
###########################################################################################################################################
def format_value(value):
    # Check if the original value is negative
    is_negative = value < 0
    # Work with the absolute value for formatting
    abs_value = abs(value)

    if abs_value >= 1000000:
        formatted_value = f"{abs_value/1000000:.1f}M"
    elif abs_value >= 1000:
        formatted_value = f"{abs_value/1000:.1f}K"
    else:
        formatted_value = str(abs_value)

    # Add the negative sign back if the original value was negative
    return f"-{formatted_value}" if is_negative else formatted_value

# Your columns with metrics
#st.markdown('### Metrics')
col1, col2, col3, col4, col5 = st.columns(5)
with col1.container():
    col1.metric("Sales", f"RM {format_value(total_sales)}", f"{format_value(compare_sales)}")
with col2.container():
    col2.metric("Profit", f"RM {format_value(total_profit)}", f"{format_value(compare_profit)}")
with col3.container():
    col3.metric("Total Item Sold", f"{total_qty} units", f"{format_value(compare_total_qty)}")
with col4.container():
    col4.metric("Average Margin per Order", f"{average_margin}", f"{format_value(compare_average_margin)}")
with col5.container():
    col5.metric("Profit Percentage", f"{profit_per}%", f"{format_value(compare_profit_per)}%")

####################################################################################################################################################################

colB1, colB2 = st.columns([2.2,3])

with colB1:
    mapfig = px.scatter_geo(df_state_sum,
                         lat='lat',
                         lon='lon',
                         size=state_option,
                         color='Shipping State',  # Set the color based on the state
                         hover_name='Shipping State',
                         projection='natural earth',
                         color_discrete_sequence=px.colors.qualitative.Plotly)

    # Ensure that the land and country borders are visible
    mapfig.update_geos(
        visible=True,  # This ensures that the geographic layout is visible
        showcountries=True,  # This shows country borders
        countrycolor="Black"  # You can customize the country border color
    )

    # Update the layout to focus on Malaysia
    mapfig.update_layout(
        geo=dict(
            scope='asia',  # Set the scope to 'asia'
            center={'lat': 4.2105, 'lon': 108.5},  # Center the map on Malaysia
            projection_scale=6,
            showland=True,  # Ensure the land is shown
            landcolor='rgb(217, 217, 217)',
            showocean=True,  # Set the land color
            oceancolor='rgb(0, 0, 0)',
            countrywidth=0.1  # Set the country border width
        )
    )

    mapfig.update_layout(width=630, height=340, title=f'{state_option} by State')
    st.plotly_chart(mapfig)

# Convert 'month' column to datetime to ensure proper sorting
df_ALL['month'] = pd.to_datetime(df_ALL['month'], format='%B')

df_ALL['Unit Total'] = pd.to_numeric(df_ALL['Unit Total'], errors='coerce')
df_ALL['Margin Per Item'] = pd.to_numeric(df_ALL['Margin Per Item'], errors='coerce')

# Now, group your data by 'month' and sum 'Unit Total' and 'Margin Per Item'
monthly_data = df_ALL.groupby(df_ALL['month'].dt.strftime('%B')).agg({'Unit Total':'sum', 'Margin Per Item':'sum'}).reset_index()

# Sort the DataFrame by 'month' in chronological order
monthly_data['month'] = pd.Categorical(monthly_data['month'], categories=[
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'],
    ordered=True)
monthly_data.sort_values('month', inplace=True)

# Create a bar chart for 'Unit Total'
bar_chart = go.Bar(
    x=monthly_data['month'],
    y=monthly_data['Unit Total'],
    name='Sales',
    marker=dict()
)

# Create a line chart for 'Margin Per Item'
line_chart = go.Scatter(
    x=monthly_data['month'],
    y=monthly_data['Margin Per Item'],
    name='Profit',
    mode='lines+markers',
    line=dict(color='red')
)

# Create a layout with titles and axis labels
layout = go.Layout(
    title='Monthly Sales and Profit',
    xaxis=dict(title='Month'),
    yaxis=dict(title='Sales'),
    yaxis2=dict(title='Profit', overlaying='y', side='right')
)

# Combine the bar and line charts
barline_fig = go.Figure(data=[bar_chart, line_chart], layout=layout)

with colB2:
    barline_fig.update_layout(width=900, height=340)
    st.plotly_chart(barline_fig)
####################################################################################################################################################################
colC1, colC2, colC3, colC4 = st.columns([2.2,1,1,1])
with colC1:
    df_manu_sum = df_state.groupby(['Manufacturer'], as_index=False).agg({
        'Unit Total': 'sum',
        'Margin Per Item': 'sum',
        'Quantity': 'sum'
    })
    df_manu_sum.rename(columns={
        'Unit Total': 'Sales',
        'Margin Per Item': 'Profit',
        'Quantity': 'Units'
    }, inplace=True)

    # Assuming 'df' is your DataFrame with the 'manufacturer' and 'sales' columns
    bar1_fig = px.bar(df_manu_sum, x='Manufacturer', y=f'{state_option}',
                 color='Manufacturer')  # This will assign a different color to each manufacturer

    # Show the figure
    bar1_fig.update_layout(width=600,height=320, title=f"{state_option} by Brand")
    st.plotly_chart(bar1_fig)

def group_small_slices(df, value_column, category_column, threshold=0.05):
    # Calculate the total sum of the values
    total = df[value_column].sum()

    # Find the slices that are smaller than the threshold
    small_slices = df[df[value_column] / total < threshold]

    # Sum the values of the small slices to create the 'Others' category
    others_sum = small_slices[value_column].sum()

    # Remove the small slices from the DataFrame
    df = df[df[value_column] / total >= threshold]

    # Add the 'Others' category row
    others_row = pd.DataFrame({category_column: ['Others'], value_column: [others_sum]})
    df = pd.concat([df, others_row], ignore_index=True)

    return df

df_marketplace = df['Order Source'].value_counts().reset_index()
df_payment = df['Payment Method'].value_counts().reset_index()
df_courier = df['Courier'].value_counts().reset_index()


with colC2:
    # Create the pie chart
    pie_fig1 = go.Figure(data=[go.Pie(labels=df_marketplace['Order Source'], values=df_marketplace['count'], hole=0.5)])

    pie_fig1.update_layout(width=300,height=320, title='Order by MarketPlace')
    st.plotly_chart(pie_fig1)
with colC3:
    df_pie2 = group_small_slices(df_payment, 'count', 'Payment Method')
    # Create the pie chart
    pie_fig2 = go.Figure(data=[go.Pie(labels=df_pie2['Payment Method'], values=df_pie2['count'], hole=0.7, textinfo="none", sort=False)])
    pie_fig2.update_layout(width=300,height=320, title='Preferred Payment Methods')
    st.plotly_chart(pie_fig2)
with colC4:
    df_pie3 = group_small_slices(df_courier, 'count', 'Courier')
    # Create the pie chart
    pie_fig3 = go.Figure(data=[go.Pie(labels=df_pie3['Courier'], values=df_pie3['count'], textinfo="label", sort=False)])
    pie_fig3.update_layout(width=300,height=320, showlegend=False, title='Shipping Courier')
    st.plotly_chart(pie_fig3)
####################################################################################################################################################################
colD1, colD2 = st.columns(2)
with colD1:
    df_category_sum = df_state.groupby(['Category'], as_index=False).agg({
        'Unit Total': 'sum',
        'Margin Per Item': 'sum',
        'Quantity': 'sum'
    })
    df_category_sum.rename(columns={
        'Unit Total': 'Sales',
        'Margin Per Item': 'Profit',
        'Quantity': 'Units'
    }, inplace=True)

    df_category_sum =  df_category_sum.sort_values([f'{state_option}'], ascending=True).tail(10)

    # Assuming 'df' is your DataFrame with the 'manufacturer' and 'sales' columns
    bar2_fig = px.bar(df_category_sum, x=f'{state_option}', y='Category',
                 color_discrete_sequence=['purple'])  # This will assign a different color to each manufacturer

    # Show the figure
    bar2_fig.update_layout(height=310, title=f"Top 10 Category by {state_option}")
    st.plotly_chart(bar2_fig)
with colD2:
    df_prod_sum = df_state.groupby(['Model'], as_index=False).agg({
        'Unit Total': 'sum',
        'Margin Per Item': 'sum',
        'Quantity': 'sum'
    })
    df_prod_sum.rename(columns={
        'Unit Total': 'Sales',
        'Margin Per Item': 'Profit',
        'Quantity': 'Units'
    }, inplace=True)

    df_prod_sum =  df_prod_sum.sort_values([f'{state_option}'], ascending=True).tail(10)

    # Assuming 'df' is your DataFrame with the 'manufacturer' and 'sales' columns
    bar3_fig = px.bar(df_prod_sum, x=f'{state_option}', y='Model',
                 color_discrete_sequence=['orange'])  # This will assign a different color to each manufacturer

    # Show the figure
    bar3_fig.update_layout(height=310, title=f"Top 10 Product by {state_option}")
    st.plotly_chart(bar3_fig)
