import xlwt
import sys
import datetime
ezxf = xlwt.easyxf

def write_xls(file_name, sheet_name, headings, data, heading_xf, data_xfs):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)
    rowx = 8 #0
    
    # Write and format the headings for the sheet
    for colx, value in enumerate(headings):
        sheet.write(rowx, colx, value, heading_xf)
    #sheet.set_panes_frozen(True) # frozen headings instead of split panes
    #sheet.set_horz_split_pos(rowx+1) # in general, freeze after last heading row
    #sheet.set_remove_splits(True) # if user does unfreeze, don't leave a split there
    
    # Write and format the body/data for the sheet
    for row in data:
        # The row number is defined for the cell, and a new row is added on each time 
        rowx += 1
        # The column number is defined for each cell, and the value to be writen is pulled
        # out using the enumerate list function.
        for colx, value in enumerate(row):
            print ("Colx is:", colx)
            print ("value is:", value)
            # A new sheet is writen with the cell row number identified,
            # the cell column number identified, the value that will 
            # be written to the cell defined AND... Magically the 
            # cell formating is defined with the 'data_xfs[colx]'. This
            # is really important, but I don't know how it works.
            sheet.write(rowx, colx, value, data_xfs[colx])
    book.save(file_name)

# Output file location
excelOutput = "C:\\Users\\aolson\\Documents\\Working\\Plus_Metro\\PYTHON\\Reports\\"

# Datetime to make date values in the nested list
mkd = datetime.date

# The column header names
hdngs = ['Date', 'Stock Code', 'Quantity', 'Unit Price', 'Value', 'Message']

# The different format types/'kinds' for each column !!!ORDER MATTERS!!! ORDER IS
# SUPER FUCKING IMPORTANT, BUT I DON"T KNOW WHY"
kinds =  'date text int price money text'.split()

# The formating for the header using xlwt.easyxf
heading_xf = ezxf('font: bold on; align: wrap on, vert centre, horiz center')    

# Nested list of data for inserting into excel
data = [[mkd(2007, 7, 1), 'ABC', 1000, 1.234567, 1234.57, ''],
        [mkd(2007, 12, 31), 'XYZ', -100, 4.654321, -465.43, 'Goods returned'],
        ] + [[mkd(2008, 6, 30), 'PQRCD', 100, 2.345678, 234.57, ''],] * 100

# Dictionaty for storing/defining format values
kind_to_xf_map = {
    'text': ezxf(),
    'date': ezxf(num_format_str='yyyy-mm-dd'),
    'int': ezxf(num_format_str='#,##0'),
    'money': ezxf('font: italic on; pattern: pattern solid, fore-colour grey25',
                  num_format_str='$#,##0.00'),
    'price': ezxf(num_format_str='#0.000000'), 
}

# Dictionary is referenced with 'kind_to_xf_map[k]', which is all
# the key values in the dictionary. The dictionary is linked to the 
# different values in 'kinds'. This is really important, but I don't know
# how it works.
data_xfs = [kind_to_xf_map[k] for k in kinds]

# Call the function
write_xls(excelOutput + 'xlwt_easyxf_simple_demo.xls', 'Demo', hdngs, data, heading_xf, data_xfs)