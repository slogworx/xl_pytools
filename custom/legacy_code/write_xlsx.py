""" write_xlsx() - Outputs a formatted row in Excel workbook format

:row_data: A list of cell data for the row
:cell_loc: Must be a dict with 'row' and 'col' as keys, e.g. {'row': 0, 'col': 0}
:workbook: <class 'xlsxwriter.workbook.Workbook'> via xlsxwriter.Workbook(output filename)
:worksheet: <class 'xlsxwriter.worksheet.Worksheet'> via workbook.add_worksheet(worksheet name)
:header: True if the row is a header, False if it is not

"""

import xlsxwriter
import datetime

preferred_font = 'Arial'
border_thickness = 1

def write_xlsx(row_data, cell_loc, workbook, worksheet, header = False):

    if header:
        xl_format = workbook.add_format({'bold': True})
        xl_format.set_bg_color('#C5D9F1')
        xl_format.set_border(border_thickness)
        xl_format.set_font_name(preferred_font)
    else:
        # Formatting data values
        xl_format = workbook.add_format()
        xl_format.set_font_name(preferred_font)
        xl_format.set_border(border_thickness)

        # Formatting datetime only
        dt_format = workbook.add_format()
        dt_format.set_font_name(preferred_font)
        dt_format.set_border(border_thickness)
        dt_format.set_num_format('mm/dd/yyyy hh:mm AM/PM')

    for cell in row_data:
            # Check for datetime.datetime, and write out the row
            if isinstance(cell, datetime.datetime):
                worksheet.write(cell_loc['row'], cell_loc['col'], cell, dt_format)
            else:
                worksheet.write(
                    cell_loc['row'], cell_loc['col'], cell, xl_format)
            
            cell_loc['col'] +=1

    # Always reset the col for the next row ()
    worksheet.set_column(0, cell_loc['col'], 20) 
    cell_loc['col'] = 0 
