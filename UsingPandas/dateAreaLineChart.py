from datetime import date
from dateutil.relativedelta import relativedelta
import pandas as pd
import random

count = 0
dates = []

#Generate 12 months worth of dates and append them to a list for the 
#first column of data
while count < 12:
    count += 1
    iDate = date.today() + relativedelta(months=+count)
    #print(count)
    #print(iDate)
    dates.append(iDate)
#print(dates)

# Create some sample data to plot.
max_row     = 12
categories  = ['Node 1', 'Node 2', 'Node 3', 'Node 4']
multi_iter1 = {'Date': dates}

for category in categories:
    multi_iter1[category] = [random.randint(10, 100) for x in dates]

# Create a Pandas dataframe from the data.
index_2 = multi_iter1.pop('Date')
df      = pd.DataFrame(multi_iter1, index=index_2)
df      = df.reindex(columns=sorted(df.columns))

# Create a Pandas Excel writer using XlsxWriter as the engine.
sheet_name = 'Collections'
writer     = pd.ExcelWriter('DateArea_Chart.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name=sheet_name)

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook  = writer.book
worksheet = writer.sheets[sheet_name]

# Create a chart object.
chart = workbook.add_chart({'type': 'line'})

# Configure the series of the chart from the dataframe data.
for i in range(len(categories)):
    col = i + 1
    chart.add_series({
        'name':       ['Collections', 0, col],
        'categories': ['Collections', 1, 0,   max_row, 0],
        'values':     ['Collections', 1, col, max_row, col],
    })

# Configure the chart axes.
chart.set_x_axis({'name': 'Date'})
chart.set_y_axis({'name': 'Value', 'major_gridlines': {'visible': False}})

# Insert the chart into the worksheet.
worksheet.insert_chart('G2', chart)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

    