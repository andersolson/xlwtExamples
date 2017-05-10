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
    hdngs = ['City', 'Country', 'Directory_Name', 'FTP_Folder', 'GSD', 'Quality_Tier',
             'Production_Version', 'Same_as_PlusVivid_HVA', 'Oldest_Date_of_Image',
             'Newest_Date_of_Image', 'KM2', 'Date_Delivered', 'FTP_RollOff']
    kinds =  'text   text   text   text   int0   text   int1   text   date   date   int2   date   date'.split()
    data = [
        ['Madrid', 'Spain', 'c://','c://', .5, 'Premium', 1, 'NO DATA', mkd(2007, 12, 31), mkd(2007, 12, 31), 345.567, mkd(2007, 12, 31), '03/26/2016'],
        ['Granada', 'Spain', 'c://','c://', .5, 'Standard', 3, 'NO DATA', mkd(2005, 11, 8), mkd(2005, 11, 8), 35.567, mkd(2005, 11, 8), mkd(2005, 11, 8)],
        ['Barcelona', 'Spain', 'c://','c://', .5, 'Standard', 5, 'NO DATA', mkd(2007, 11, 8), mkd(2007, 11, 8), 3500.567, mkd(2007, 11, 8), mkd(2005, 7, 8)]
        ] 

    heading_xf = ezxf('font: bold on; align: wrap off, vert centre, horiz center')
    kind_to_xf_map = {
        'date': ezxf(num_format_str='M/D/YY'),
        'int0': ezxf(num_format_str='0.0'),
        'int1' : ezxf(num_format_str='0'),
        'int2' : ezxf(num_format_str='0.00'),
        'text': ezxf(),
        }
    data_xfs = [kind_to_xf_map[k] for k in kinds]
    write_xls('C:\\Users\\aolson\\Documents\\Working\\Plus_Metro\\PYTHON\\Reports\\PlusMetro_xlwt_demo.xls', 'Demo', hdngs, data, heading_xf, data_xfs)