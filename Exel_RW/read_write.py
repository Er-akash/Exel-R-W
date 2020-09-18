import xlrd
import xlsxwriter
import os.path

EXCEL_FILES_FOLDER = '/home/pi/Desktop/Akash/Exel_RW/'

excel_file_path = EXCEL_FILES_FOLDER+'Acc_2.xlsx'
try:
        with open(excel_file_path, 'r') as file:
            loc = (excel_file_path)
            wb = xlrd.open_workbook(loc)
            sheetAk = wb.sheet_by_name('My sheet')

            rowsCount = sheetAk.nrows
            colsCount = sheetAk.ncols

            ##print("rowsCount: ",rowsCount," colsCount: ",colsCount)
            for i in range(sheetAk.nrows ): 
                print(sheetAk.cell_value(i,0))
            ##    print("h",i)
            print(sheetAk.row_values(4)) 
            for i in range(0,rowsCount):
                print('\nrow: ',i+1)
                for j in range(colsCount):
                    col_name = sheetAk.cell_value(0, j)
                    cell_val = sheetAk.cell_value(i, j)
                    print(col_name,': ',cell_val)

        
except FileNotFoundError:
    print("OSError File Not found Please Create New file")
    workbook = xlsxwriter.Workbook('Acc_2.xlsx')
    worksheet = workbook.add_worksheet("My sheet")
    scores = ( 
        ['Work', 'Status'], 
        ['RTC Installation', 'Yes'], 
        ['Any Dest cmd', 'Yes'], 
        ['Exel', 'No'],
        ['LSDP Testing', 'Yes'],
        ['browser ui', 'Yes'],
        )
    row = 0
    col = 0

    for name, score in (scores): 
            worksheet.write(row, col, name) 
            worksheet.write(row, col + 1, score) 
            row += 1

    workbook.close()
    print("New Exel Sheet is Created")
except IndexError:
        exit1=0
        print(" Index Error")

