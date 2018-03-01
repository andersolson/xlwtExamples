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

# Changed field name to text
df = pd.DataFrame({'Names': ["Peter","Ruth","Bob","Sarah"],
                   'BirthDays': ['12/3/2012','4/3/99','6/24/1985','9-12-87'],
                   'Job': ["Police Man","Bar tender","fluffer","lumber jack"],
                   'A Number': [1, 3.45, 5, 1000000]})

# Change output file name. 
writer = pd.ExcelWriter('Professions.xlsx', engine='xlsxwriter')

# Change the sheet name.
df.to_excel(writer, sheet_name='Profession')

# Close the Pandas Excel writer and output the Excel file.
# Output excel file is written in the same location where the
# script is saved.
writer.save()

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

# Making multiple sheets
df0 = pd.DataFrame({'Names': ["Peter","Ruth","Bob","Sarah"],
                   'BirthDays': ['12/3/2012','4/3/99','6/24/1985','9-12-87'],
                   'Job': ["Police Man","Bar tender","fluffer","lumber jack"],
                   'A Number': [1, 3.45, 5, 1000000]})

df1 = pd.DataFrame({'Names': ["Bill", "Anne", "Todd"],
                   'Age': [5, 17, 89]})

# Change output file name. 
writer = pd.ExcelWriter('DoubleOrNothing.xlsx', engine='xlsxwriter')

# Write 2 sheets
df0.to_excel(writer, sheet_name='Profession')
df1.to_excel(writer, sheet_name='Age')

# Close the Pandas Excel writer and output the Excel file.
# Output excel file is written in the same location where the
# script is saved.
writer.save()