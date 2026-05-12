import xlwings as xw

app = xw.App(visible=True) 

app.screen_updating = False
app.calculation = 'manual'
app.display_alerts = False

try:
    wb = app.books.open('Nil UL & NS Analysis.xlsx') 
    sheet = wb.sheets['Analysis-Aug\'25']

    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    
    target_range = f'H2:ZZ{last_row}'

    sheet.range(target_range).clear_contents()
    print("success!")

    wb.save()

finally:
    app.screen_updating = True
    app.calculation = 'automatic'
    app.display_alerts = True