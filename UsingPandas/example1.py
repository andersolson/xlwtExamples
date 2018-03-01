import pandas as pd

# Create a Pandas dataframe from the data.
df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})

# Create a Pandas Excel writer using XlsxWriter as the engine. 
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
# Output excel file is written in the same location where the
# script is saved.
writer.save()

#-------------------- 2nd part, experiment on my own

# Changed field name to text
df = pd.DataFrame({'Names': ["Peter", "Paul", "Bob", "Teddy", "Sofia"]})

# Change output file name. 
writer = pd.ExcelWriter('Names.xlsx', engine='xlsxwriter')

# Change the sheet name.
df.to_excel(writer, sheet_name='Names')

# Close the Pandas Excel writer and output the Excel file.
# Output excel file is written in the same location where the
# script is saved.
writer.save()