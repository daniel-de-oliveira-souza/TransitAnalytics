# -*- coding: utf-8 -*-
"""
Python script that performs data analysis and visualization on logistics data. The code imports two classes, "DataProcessor" and "DataVisualizer",
from a module called "classes".

The main execution of the code begins by defining a directory path for Excel files. 
It then creates an instance of the "DataProcessor" class, which contains methods for processing Excel files and returning a pandas DataFrame.
"""
from logistics_classes import DataProcessor, DataVisualizer, JsonToMarkdownConverter
import pandas as pd

# Main execution
my_xlsx_dir = '/workspaces/transitanalytics/excel_files'
processor = DataProcessor(my_xlsx_dir)
df = processor.process_xlsx_files()

# Determine column names based on DataFrame structure
sys_stkout_column_exists = "Sys_StkOut" if 'Sys_StkOut' in df.columns else None
monthnumber_or_weeknumber = 'Month_Num' if 'Month_Num' in df.columns else 'WeekNumber'
wh_or_sr = 'WH' if 'WH' in df.columns else 'SR'
srdescription_or_whname = 'SR Description' if 'SR Description' in df.columns else 'WH Name'

try:
    # Set the new index for sorting and rearrange columns if needed
    df = df.set_index([wh_or_sr])

    # If Sys_StkOut column exists, rearrange it to the desired position
    if sys_stkout_column_exists:
        column_order = [col for col in df.columns if col != sys_stkout_column_exists]
        column_order.insert(3, sys_stkout_column_exists)  # Insert at the 4th position (index 3)
        df = df[column_order]

    # Sort the DataFrame once by the appropriate columns
    df = df.sort_values(by=[monthnumber_or_weeknumber, wh_or_sr])

    # Reset index and set multi-level index for grouped data
    grouped_df = df.reset_index()
    grouped_df = grouped_df.set_index([monthnumber_or_weeknumber, wh_or_sr])

except KeyError as e:
    print(f"Error: DataFrame does not have the expected columns. Missing column: {e}")

# Group the DataFrame by the specified column
grouped = df.groupby(monthnumber_or_weeknumber)

# Define the columns to sum
columns_to_sum = ["FM Count", "InStock", "StockOut"]
if sys_stkout_column_exists:  # Check if 'Sys_StkOut' column exists and add it if necessary
    columns_to_sum.append(sys_stkout_column_exists)

try:
    # Calculate the sum for the specified columns and round the results
    totals = grouped[columns_to_sum].sum().round(2)

    # Calculate the %StockOut, ensuring no division by zero occurs
    totals['%StockOut'] = (totals['StockOut'] / totals['FM Count']).fillna(0).round(3)

except KeyError as e:
    print(f"Error: DataFrame does not have one of the expected columns. Missing column: {e}")

totals['Avg_%StockOut_Change'] = totals['%StockOut'].pct_change().round(3)

# Check for necessary columns before proceeding with further operations
if 'Date' in totals.columns:
    # Check for 'WeekNumber' column
    if 'WeekNumber' in totals.columns:
        # Rearrange columns based on whether 'Sys_StkOut' column exists
        selected_columns = ['WeekNumber', 'Date', 'FM Count', 'InStock', 'StockOut', 'Sys StkOut', '%StockOut'] if sys_stkout_column_exists in df.columns else ['Date', 'FM Count', 'InStock', 'StockOut', '%StockOut']
        totals = totals[selected_columns]
        totals = totals.sort_values(by=['Date'])
    # Rename '%StockOut' to 'Avg_%StockOut'
    totals = totals.rename(columns={'%StockOut': 'Avg_%StockOut'})

    # Calculate and format 'Avg_%StockOut_Change'
    totals['Avg_%StockOut_Change'] = totals['Avg_%StockOut'].pct_change().round(3)
    totals['Avg_%StockOut'] = (totals['Avg_%StockOut'] * 100).round(2).astype(str) + '%'
    totals['Avg_%StockOut_Change'] = (totals['Avg_%StockOut_Change'] * 100).round(2).astype(str) + '%'
    totals['Avg_%StockOut_Change'].fillna(0, inplace=True)

    # Rearrange the totals DataFrame
    totals_df = totals[['Date'] + [col for col in totals.columns if col != 'Date']]
else:
    totals_df = totals

# Display the totals DataFrame (optional)
# print(totals_df)
    
# Reset the index to convert the multi-index into columns
df_reset = df.reset_index()
grouped_df_reset = grouped_df.reset_index()
totals_df_reset = totals_df.reset_index()   

# Save DataFrames as JSON
df_reset.to_json('/workspaces/transitanalytics/content/cumulative_data.json', orient='split')
grouped_df_reset.to_json('/workspaces/transitanalytics/content/grouped_data.json', orient='split')
totals_df_reset.to_json('/workspaces/transitanalytics/content/summary_stats.json', orient='split')

visualizer = DataVisualizer(df)

# Call to create_bar_plot and create_line_plot
plot1_path = visualizer.create_bar_plot(df, srdescription_or_whname, monthnumber_or_weeknumber, wh_or_sr)
plot2_path = visualizer.create_line_plot(totals_df_reset, monthnumber_or_weeknumber)

# JsonToMarkdown
json_dir = '/workspaces/transitanalytics/content'
markdown_dir = '/workspaces/transitanalytics/content'
converter = JsonToMarkdownConverter(json_dir, markdown_dir)
converter.convert()

