# Project-Automation-of-Daily-Excel-Tasks-for-Operations-Team

Being part of an operations data analytics team, we deal with the same kind of data every day. some times our workflow involved taking two or three Excel files, importing those files, creating pivot tables, and summarizing the data. This process used to take almost 1 hour daily.

📌 Project Overview

The objective was to automate this repetitive workflow, streamline data processing, and deliver insights efficiently.

📊 Key Highlights:

✔ Automated data merging, cleaning, and transformation.

✔ Analyzed sales trends, top-performing products, and customer behavior.

✔ Generated pivot tables for in-depth analysis.
💡 Technologies Used: Python, Pandas, Excel



👩‍💻 Impact:

Reduced manual effort and improved accuracy.

Delivered faster insights for better decision-making.

Demonstrated the power of automation in operations.

📈 Key Learnings:

This project emphasized the value of automation in streamlining daily workflows and highlighted how small optimizations can have a big impact on efficiency and productivity.


✔ Delivered a consolidated multi-sheet Excel report.

## Code Part ---> Explained and designed Step By step
## Project 2 - Automation on daily excel tasks for operations team.

### Necessary libraries for automating the process
import pandas as pd
import tkinter as tk
from tkinter import filedialog,messagebox,simpledialog
from tkinter import *
from tkinter.filedialog import askopenfilename
import os
import sys
import datetime

## initiating the tkinter window
root = tk.Tk()
root.withdraw()

## Loading the DataSets
sales_data_file = filedialog.askopenfilename(title = "Select the sales data file")
customers_data_file =filedialog.askopenfilename(title = "Select the customers data file")
output_path = filedialog.askdirectory(title="Choose the loacatio for saving the file (ouputpath)")

## Reading the files
sales_data = pd.read_csv(sales_data_file,encoding = 'latin')
customers_data = pd.read_csv(customers_data_file,encoding='latin')

## merging the both the files into one consolidated sheet using order_id common column
consolidated_file = pd.merge(sales_data,customers_data,on='Order ID',how='inner')

## Removing the duplicates from the combined file
removing_duplicates = consolidated_file.drop_duplicates()

## handling the missing values if any
consolidated_file['Sales'] = consolidated_file['Sales'].fillna(consolidated_file['Sales'].mean())

## filtering the data based on the specific region
filtered_data = consolidated_file[consolidated_file['Region'] == 'North']

## Grouping the data By catogory 
category_summary = filtered_data.groupby('Category').agg(
    Total_sales = ('Sales','sum'),
    order_count = ('Order ID','count')
).reset_index()

## Pivot table analysis 
sales_region_analysis = pd.pivot_table(consolidated_file,
                                 values='Sales',
                                 index='Category',
                                 columns='Region',
                                 aggfunc ='sum',
                                 fill_value = 0
                                 )

orders_state_analysis = pd.pivot_table(consolidated_file,
                                 values='Order ID',
                                 index='State',
                                 columns= 'Category',
                                 aggfunc = 'count',
                                 fill_value = 0
                                 )

## Monthly trend analysis
consolidated_file['Month'] = pd.to_datetime(consolidated_file['Date']).dt.to_period('M')
mothly_sales_trend = consolidated_file.groupby('Month').agg(Total_sales = ('Sales','sum')).reset_index()

## Top performeers
top_products = consolidated_file.groupby('Product').agg(Total_Sales=('Sales', 'sum')).nlargest(5, 'Total_Sales').reset_index()
top_customers = consolidated_file.groupby('Customer Name').agg(Total_Sales=('Sales', 'sum')).nlargest(5, 'Total_Sales').reset_index()

## Correlation analysis (simple computational example)
correlation_sales_orders = consolidated_file.groupby("Category").agg(
    Total_Sales = ('Sales','sum'),
    order_count = ('Order ID','count')
).corr()

## writing these results into excel sheet 
current_date = datetime.now().strftime('%Y-%m-%d')
file_path = f"{output_path}/Final Observation File.xlsx{current_date}"
with pd.ExcelWriter(file_path,engine='openpyxl') as writer:
    consolidated_file.to_excel(writer,sheet_name='Combined_clean_Data',index=False)
    category_summary.to_excel(writer,sheet_name='Category Symmary',index=False)
    sales_region_analysis.to_excel(writer,sheet_name='sales region analysis',index=False)
    orders_state_analysis.to_excel(writer,sheet_name = 'order sate analysis',index = False)
    mothly_sales_trend.to_excel(writer,sheet_name='Monthly Slaes Trend',index=False)
    top_products.to_excel(writer,sheet_name='Top Products',index=False)
    top_customers.to_excel(writer,sheet_name='Top Customers',index=False)
    correlation_sales_orders.to_excel(writer,sheet_name='correlation_sales_orders',index=False)
    
print(f"The Analysis File is Genrated at{output_path}")


###### Note:I have attached all the data Sets required for this Project please go through the following.


