### Graph Plotting Automation on Excel ###

# Based on Excel file, and the output would also be on excel.
# The gragh will be presented on same column comparing against different sheets, and each sheet should have same format (same column)

import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd


### Select Input File (Only one)
root = tk.Tk()
file = fd.askopenfilenames()
data = pd. ExcelFile(file[0])
root.destroy()


### New variables for the progress
new_table = pd.DataFrame()
list_columns_removed = [                                        # Remove some columns, which would not be plotted
    "Column name 1",
    "Column name 2"
    ]
list_columns_plot = [                                           # Select which columns would be plotted
    "Column name 1",
    "Column name 2"
    ]
sheet_name = []

### Function here would be used in the progress
# Function: creating table based on column name
def making_table(index):
    new_table["Time"] = Calc_data[0]["Time"]                    # Assign same x-axis for all table, "Time" as an example
    n = 0
    for table in Calc_data:
        column_name = table["Strain_ID"][0]                     # Set Column name as individual ID 
        new_table[column_name] = table[index]                   # Copy and paste from original to new table a based on new ID
        n += 1
    globals()[index] = new_table
    return globals()[index]

# Function: change colume's name to meet any purpose, Excel name restriction
def change_columns_name(table):
    table.rename(columns = {
        "old name 1" : "new name 1",
        "ols name 2" : "new name 2"
    }, inplace = True)


### Extract data from file and been saved as dataframe in list
for sheet in data.sheet_names:
    if "Calc" in sheet:                                         # Select specific sheet from file, "Calc" as an example
        sheet_name.append(sheet)
sheet_name = sheet_name[1:]                                     # remove specific sheet if needed


### Data sorting and organizing for each table
Calc_data = []
for n in sheet_name:
    table = data.parse(n)
    table.drop(labels = list_columns_removed, axis = 1, inplace = True)
    change_columns_name(table)
    table.drop(table.index[13:], 0, inplace = True)
    Calc_data.append(table)

# Graph plotting automation
graph_data = pd.ExcelWriter('graph.xlsx', engine = 'xlsxwriter')

max_row = len(Calc_data[1])
for name in list_columns_plot:
    new_table = making_table(name)
    new_table.to_excel(graph_data, sheet_name = name)

    workbook = graph_data.book
    worksheet = graph_data.sheets[name]
    chart = workbook.add_chart({'type':'scatter',
                                'subtype':'straight_with_markers'})

    for n in range(len(sheet_name)):
        col = n + 2
        chart.add_series({
            'name':         [name, 0, col],
            'categories':   [name, 1, max_row, 1],
            'values':       [name, 1, col, max_row, col],
            'marker':       {'type':'circle', 'size':4},
        })

    worksheet.insert_chart('K2', chart)

graph_data.save()
