


import xlwings as xw

def letter_to_num(letter):
    num = 0
    for c in letter.upper():
        num = num * 26 + (ord(c) - ord('A')) + 1
    return num

def process_excel(filepath, month, year, logger=None):
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)

    app = xw.App(visible=True) 
    app.screen_updating = False
    app.calculation = 'manual'
    app.display_alerts = False

    try:
        log(f"Opening workbook: {filepath}")
        wb = app.books.open(filepath) 
        sheet_name = f"Analysis-{month}'{year}"
        log(f"Accessing sheet: {sheet_name}")
        sheet = wb.sheets[sheet_name]
        
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        ul_map = [
            ('A', 'B'), ('A', 'C'), ('A', 'D'), ('A', 'E'), ('A', 'F'),
            ('A', 'G'), ('A', 'H'), ('A', 'I'), ('A', 'J'), ('A', 'K'),
            ('A', 'L'), ('A', 'M'), ('A', 'N'), ('A', 'O'), ('A', 'P'),
            ('A', 'Q'), ('A', 'R'), ('A', 'S'), ('T', 'U'), ('V', 'W'),
            ('T', 'U'), ('V', 'W'), ('V', 'X'), ('V', 'Y'), ('Z', 'AA'),
            ('AB', 'AC'), ('AB', 'AD'), ('AE', 'AF'), ('AG', 'AH'), ('AG', 'AI'), ('AG', 'AJ')
        ]

        for day in range(1, 32):
            log(f"Writing formulas for Day {day}...")
            nil_col = xw.utils.col_name(8 + ((day - 1) * 3))
            ul_col = xw.utils.col_name(9 + ((day - 1) * 3))
            ns_col = xw.utils.col_name(10 + ((day - 1) * 3))

            ul_ro_col, ul_data_col = ul_map[day-1]
            ul_idx = letter_to_num(ul_data_col) - letter_to_num(ul_ro_col) + 1
            sheet.range(f'{ul_col}2:{ul_col}{last_row}').formula = \
                f'=IFNA(VLOOKUP($A2, Upliftment!${ul_ro_col}:${ul_data_col}, {ul_idx}, 0), 0)'

            ns_start_num = 1 + ((day - 1) * 4)
            ns_ro_letter = xw.utils.col_name(ns_start_num)
            ns_data_letter = xw.utils.col_name(ns_start_num + 3)
            sheet.range(f'{ns_col}2:{ns_col}{last_row}').formula = \
                f'=IFERROR(VLOOKUP($A2, \'Nozzle Sales\'!${ns_ro_letter}:${ns_data_letter}, 4, 0), 0)'

            sheet.range(f'{nil_col}2:{nil_col}{last_row}').formula = f'=AND({ul_col}2=0, {ns_col}2=0)'

        log("Calculating and freezing...")
        app.calculate()
        
        sheet.range(f'H2:CV{last_row}').value = sheet.range(f'H2:CV{last_row}').value

        log("Success! Formulas injected and values frozen.")
        wb.save()

    finally:
        app.screen_updating = True
        app.calculation = 'automatic'
        app.display_alerts = True