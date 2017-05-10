excelOutput = "C:\\Users\\aolson\\Documents\\Working\\Plus_Metro\\PYTHON\\Reports\\"

#===============================#
# Make a reference table for font color
#===============================#

#from xlwt import *

#font0 = Font()
#font0.name = 'Times New Roman'
#font0.bold = True

#style0 = XFStyle()
#style0.font = font0

#wb = Workbook()
#ws0 = wb.add_sheet('data')

#ws0.write(1, 1, 'Test', style0)

#for i in range(0, 0x53):
   #fnt = Font()
   #fnt.name = 'Arial'
   #fnt.colour_index = i
   #fnt.outline = True

   #borders = Borders()
   #borders.left = i

   #style = XFStyle()
   #style.font = fnt
   #style.borders = borders

   #ws0.write(i, 2, 'colour', style)
   #ws0.write(i, 3, hex(i), style0)

#wb.save(excelOutput + 'colours.xls')

#===============================#
# Example for making and naming a sheet
#===============================#

#from xlwt import *
#
#w = Workbook()
#w.country_code = 61
#ws = w.add_sheet('AU')
#w.save(excelOutput + 'country.xls')

#===============================#
# Make an example table using easyxf
#===============================#

#import xlwt
#from datetime import datetime

#style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    #num_format_str='#,##0.00')
#style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

#wb = xlwt.Workbook()
#ws = wb.add_sheet('A Test Sheet')

#ws.write(0, 0, 1234.56, style0)
#ws.write(1, 0, datetime.now(), style1)
#ws.write(2, 0, 1)
#ws.write(2, 1, 1)
#ws.write(2, 2, xlwt.Formula("A3+B3"))

#wb.save(excelOutput + 'example.xls')

#===============================#
# Make a reference table of all the 
# number formats for a cell in excel
#===============================#

#from xlwt import *

#w = Workbook()
#ws = w.add_sheet('Hey, Dude')

#fmts = [
    #'general',
    #'0',
    #'0.00',
    #'#,##0',
    #'#,##0.00',
    #'"$"#,##0_);("$"#,##',
    #'"$"#,##0_);[Red]("$"#,##',
    #'"$"#,##0.00_);("$"#,##',
    #'"$"#,##0.00_);[Red]("$"#,##',
    #'0%',
    #'0.00%',
    #'0.00E+00',
    #'# ?/?',
    #'# ??/??',
    #'M/D/YY',
    #'D-MMM-YY',
    #'D-MMM',
    #'MMM-YY',
    #'h:mm AM/PM',
    #'h:mm:ss AM/PM',
    #'h:mm',
    #'h:mm:ss',
    #'M/D/YY h:mm',
    #'_(#,##0_);(#,##0)',
    #'_(#,##0_);[Red](#,##0)',
    #'_(#,##0.00_);(#,##0.00)',
    #'_(#,##0.00_);[Red](#,##0.00)',
    #'_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
    #'_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    #'_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    #'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    #'mm:ss',
    #'[h]:mm:ss',
    #'mm:ss.0',
    #'##0.0E+0',
    #'@'   
#]

#i = 0
#for fmt in fmts:
    #ws.write(i, 0, fmt)

    #style = XFStyle()
    #style.num_format_str = fmt

    #ws.write(i, 4, -1278.9078, style)

    #i += 1

#w.save(excelOutput + 'num_formats.xls')

#===============================#
# Make a huge excel table and 
# and track the time it takes to make
#===============================#

#from time import *
#from xlwt.Workbook import *
#from xlwt.Style import *

#style = XFStyle()

#wb = Workbook()
#ws0 = wb.add_sheet('0')

#colcount = 200 + 1
#rowcount = 6000 + 1

#t0 = time()
#print("\nstart: %s" % ctime(t0))

#print("Filling...")
#for col in xrange(colcount):
    ##print("[%d]" % col, end=' ') 
    #for row in xrange(rowcount):
        ##ws0.write(row, col, "BIG(%d, %d)" % (row, col))
        #ws0.write(row, col, "BIG")

#t1 = time() - t0
#print("\nsince starting elapsed %.2f s" % (t1))

#print("Storing...")
#wb.save(excelOutput + 'big-16Mb.xls')

#t2 = time() - t0
#print("since starting elapsed %.2f s" % (t2))

#===============================#
# Make a reference sheet for all 
# possible border styles.
#===============================#

#from xlwt import *

#font0 = Font()
#font0.name = 'Times New Roman'
#font0.struck_out = True
#font0.bold = True

#style0 = XFStyle()
#style0.font = font0


#wb = Workbook()
#ws0 = wb.add_sheet('0')

#ws0.write(1, 1, 'Test', style0)

#for i in range(0, 0x53):
    #borders = Borders()
    #borders.left = i
    #borders.right = i
    #borders.top = i
    #borders.bottom = i

    #style = XFStyle()
    #style.borders = borders

    #ws0.write(i, 2, '', style)
    #ws0.write(i, 3, hex(i), style0)

#ws0.write_merge(5, 8, 6, 10, "")

#wb.save(excelOutput + 'blanks.xls')

#===============================#
# Not sure what this does...
#===============================#

#from xlwt import *

#w = Workbook()
#ws = w.add_sheet('Hey, Dude')

#for i in range(6, 80):
   #fnt = Font()
   #fnt.height = i*20
   #style = XFStyle()
   #style.font = fnt
   #ws.write(1, i, 'Test')
   #ws.col(i).width = 0x0d00 + i
#w.save(excelOutput + 'col_width.xls')

#===============================#
# Make a reference table of all 
# possible date formats.
#===============================#

#from xlwt import *
#from datetime import datetime

#w = Workbook()
#ws = w.add_sheet('Hey, Dude')

#fmts = [
   #'M/D/YY',
   #'D-MMM-YY',
   #'D-MMM',
   #'MMM-YY',
   #'h:mm AM/PM',
   #'h:mm:ss AM/PM',
   #'h:mm',
   #'h:mm:ss',
   #'M/D/YY h:mm',
   #'mm:ss',
   #'[h]:mm:ss',
   #'mm:ss.0',
#]

#i = 0
#for fmt in fmts:
   #ws.write(i, 0, fmt)

   #style = XFStyle()
   #style.num_format_str = fmt

   #ws.write(i, 4, datetime.now(), style)

   #i += 1

#w.save(excelOutput + 'dates.xls')

#===============================#
# Make a table with examples of how to 
# create formulas in each cell
#===============================#

#from xlwt import *

#w = Workbook()
#ws = w.add_sheet('F')

#ws.write(0, 0, Formula("-(1+1)"))
#ws.write(1, 0, Formula("-(1+1)/(-2-2)"))
#ws.write(2, 0, Formula("-(134.8780789+1)"))
#ws.write(3, 0, Formula("-(134.8780789e-10+1)"))
#ws.write(4, 0, Formula("-1/(1+1)+9344"))

#ws.write(0, 1, Formula("-(1+1)"))
#ws.write(1, 1, Formula("-(1+1)/(-2-2)"))
#ws.write(2, 1, Formula("-(134.8780789+1)"))
#ws.write(3, 1, Formula("-(134.8780789e-10+1)"))
#ws.write(4, 1, Formula("-1/(1+1)+9344"))

#ws.write(0, 2, Formula("A1*B1"))
#ws.write(1, 2, Formula("A2*B2"))
#ws.write(2, 2, Formula("A3*B3"))
#ws.write(3, 2, Formula("A4*B4*sin(pi()/4)"))
#ws.write(4, 2, Formula("A5%*B5*pi()/1000"))

###############
### NOTE: parameters are separated by semicolon!!!
###############


#ws.write(5, 2, Formula("C1+C2+C3+C4+C5/(C1+C2+C3+C4/(C1+C2+C3+C4/(C1+C2+C3+C4)+C5)+C5)-20.3e-2"))
#ws.write(5, 3, Formula("C1^2"))
#ws.write(6, 2, Formula("SUM(C1;C2;;;;;C3;;;C4)"))
#ws.write(6, 3, Formula("SUM($A$1:$C$5)"))

#ws.write(7, 0, Formula('"lkjljllkllkl"'))
#ws.write(7, 1, Formula('"yuyiyiyiyi"'))
#ws.write(7, 2, Formula('A8 & B8 & A8'))
#ws.write(8, 2, Formula('now()'))

#ws.write(10, 2, Formula('TRUE'))
#ws.write(11, 2, Formula('FALSE'))
#ws.write(12, 3, Formula('IF(A1>A2;3;"hkjhjkhk")'))

#w.save(excelOutput + 'formulas.xls')

#===============================#
# Make a table that demonstrates how 
# to write a value to a specific cell using col/row numbers
#===============================#

#from xlwt import *

#w = Workbook()
#ws = w.add_sheet('Hey, Dude')

#ws.write(0, 0, 1)
#ws.write(1, 0, 1.23)
#ws.write(2, 0, 12345678)
#ws.write(3, 0, 123456.78)

#ws.write(0, 1, -1)
#ws.write(1, 1, -1.23)
#ws.write(2, 1, -12345678)
#ws.write(3, 1, -123456.78)

#ws.write(0, 2, -17867868678687.0)
#ws.write(1, 2, -1.23e-5)
#ws.write(2, 2, -12345678.90780980)
#ws.write(3, 2, -123456.78)

#w.save(excelOutput + 'numbers.xls')

#===============================#
# Demonstrates how to freeze different panes.
#===============================#

#from xlwt import *

#w = Workbook()
#ws1 = w.add_sheet('sheet 1')
#ws2 = w.add_sheet('sheet 2')
#ws3 = w.add_sheet('sheet 3')
#ws4 = w.add_sheet('sheet 4')
#ws5 = w.add_sheet('sheet 5')
#ws6 = w.add_sheet('sheet 6')

#for i in range(0x100):
   #ws1.write(i//0x10, i%0x10, i)

#for i in range(0x100):
   #ws2.write(i//0x10, i%0x10, i)

#for i in range(0x100):
   #ws3.write(i//0x10, i%0x10, i)

#for i in range(0x100):
   #ws4.write(i//0x10, i%0x10, i)

#for i in range(0x100):
   #ws5.write(i//0x10, i%0x10, i)

#for i in range(0x100):
   #ws6.write(i//0x10, i%0x10, i)

#ws1.panes_frozen = True
#ws1.horz_split_pos = 2

#ws2.panes_frozen = True
#ws2.vert_split_pos = 2

#ws3.panes_frozen = True
#ws3.horz_split_pos = 1
#ws3.vert_split_pos = 1

#ws4.panes_frozen = False
#ws4.horz_split_pos = 12
#ws4.horz_split_first_visible = 2

#ws5.panes_frozen = False
#ws5.vert_split_pos = 40
#ws4.vert_split_first_visible = 2

#ws6.panes_frozen = False
#ws6.horz_split_pos = 12
#ws4.horz_split_first_visible = 2
#ws6.vert_split_pos = 40
#ws4.vert_split_first_visible = 2

#w.save(excelOutput + 'panes.xls')

#===============================#
# Demonstrates how to change the height of rows.
#===============================#

#from xlwt import *

#w = Workbook()
#ws = w.add_sheet('Hey, Dude')

#for i in range(6, 80):
   #fnt = Font()
   #fnt.height = i*20
   #style = XFStyle()
   #style.font = fnt
   #ws.write(i, 1, 'Test')
   #ws.row(i).set_style(style)
#w.save(excelOutput + 'row_styles.xls')

#===============================#
# Write an XLS file with a single worksheet, containing
# a heading row and some rows of data.
#===============================#

import xlwt
import datetime
ezxf = xlwt.easyxf

def write_xls(file_name, sheet_name, headings, data, heading_xf, data_xfs):
   book = xlwt.Workbook()
   sheet = book.add_sheet(sheet_name)
   rowx = 0
   for colx, value in enumerate(headings):
      sheet.write(rowx, colx, value, heading_xf)
   sheet.set_panes_frozen(True) # frozen headings instead of split panes
   sheet.set_horz_split_pos(rowx+1) # in general, freeze after last heading row
   sheet.set_remove_splits(True) # if user does unfreeze, don't leave a split there
   for row in data:
      rowx += 1
      for colx, value in enumerate(row):
         sheet.write(rowx, colx, value, data_xfs[colx])
   book.save(file_name)

if __name__ == '__main__':
   import sys
   mkd = datetime.date
   hdngs = ['Date', 'Stock Code', 'Quantity', 'Unit Price', 'Value', 'Message']
   kinds =  'date    text          int         price         money    text'.split()
   data = [
      [mkd(2007, 7, 1), 'ABC', 1000, 1.234567, 1234.57, ''],
      [mkd(2007, 12, 31), 'XYZ', -100, 4.654321, -465.43, 'Goods returned'],
      ] + [
         [mkd(2008, 6, 30), 'PQRCD', 100, 2.345678, 234.57, ''],
         ] * 100

   heading_xf = ezxf('font: bold on; align: wrap on, vert centre, horiz center')
   kind_to_xf_map = {
      'date': ezxf(num_format_str='m/d/yyyy'),
      #'date': ezxf(num_format_str='mm/dd/yyyy'),
      #'date': ezxf(num_format_str='mm-dd-yyyy'),
      #'date': ezxf(num_format_str='yyyy-mm-dd'),
      'int': ezxf(num_format_str='#,##0'),
      'money': ezxf('font: italic on; pattern: pattern solid, fore-colour grey25',
                    num_format_str='$#,##0.00'),
      'price': ezxf(num_format_str='#0.000000'),
      'text': ezxf(),
   }
   data_xfs = [kind_to_xf_map[k] for k in kinds]
   write_xls(excelOutput + 'xlwt_easyxf_simple_demo.xls', 'Demo', hdngs, data, heading_xf, data_xfs)
