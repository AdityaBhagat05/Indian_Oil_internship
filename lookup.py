import xlwings as xw

# Helper function to convert Excel column letters to numbers
def letter_to_num(letter):
    num = 0
    for c in letter.upper():
        num = num * 26 + (ord(c) - ord('A')) + 1
    return num

app = xw.App(visible=True) 

try:
    wb = app.books.open('Nil UL & NS Analysis.xlsx') 
    sheet = wb.sheets['Analysis-Aug\'25']

    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

    print("Injecting 31-day formulas LIVE (no speed hacks)...")

    # The Mapping List
    ul_map = [
        ('A', 'B'),  ('A', 'C'),  ('A', 'D'),  ('A', 'E'),  ('A', 'F'),  
        ('A', 'G'),  ('A', 'H'),  ('A', 'I'),  ('A', 'J'),  ('A', 'K'),  
        ('A', 'L'),  ('A', 'M'),  ('A', 'N'),  ('A', 'O'),  ('A', 'P'),  
        ('A', 'Q'),  ('A', 'R'),  ('A', 'S'),  ('T', 'U'),  ('V', 'W'),  
        ('V', 'X'),  ('V', 'Y'),  ('Z', 'AA'), ('AB', 'AC'),('AB', 'AD'),
        ('AE', 'AF'),('AG', 'AH'),('AG', 'AI'),('AG', 'AJ'),('AG', 'AK'),('AG', 'AL') 
    ]

    for day in range(1, 32):
        print(f"Writing formulas for Day {day}...")
        
        # Target Columns
        nil_col_letter = xw.utils.col_name(8 + ((day - 1) * 3))
        ul_col_letter = xw.utils.col_name(9 + ((day - 1) * 3))
        ns_col_letter = xw.utils.col_name(10 + ((day - 1) * 3))

        # Upliftment
        ul_ro_col = ul_map[day-1][0]
        ul_data_col = ul_map[day-1][1]
        
        # Using the new helper function for the math
        ul_index = letter_to_num(ul_data_col) - letter_to_num(ul_ro_col) + 1
        
        sheet.range(f'{ul_col_letter}2:{ul_col_letter}{last_row}').formula = \
            f'=IFNA(VLOOKUP($A2, Upliftment!${ul_ro_col}:${ul_data_col}, {ul_index}, 0), 0)'

        # Nozzle Sales
        ns_start_col_num = 1 + ((day - 1) * 4) 
        ns_ro_col = xw.utils.col_name(ns_start_col_num)
        ns_data_col = xw.utils.col_name(ns_start_col_num + 3) 
        
        sheet.range(f'{ns_col_letter}2:{ns_col_letter}{last_row}').formula = \
            f'=IFERROR(VLOOKUP($A2, \'Nozzle Sales\'!${ns_ro_col}:${ns_data_col}, 4, 0), 0)'

        # Nil Check
        sheet.range(f'{nil_col_letter}2:{nil_col_letter}{last_row}').formula = \
            f'=AND({ul_col_letter}2=0, {ns_col_letter}2=0)'

    print("\nSuccess! All formulas injected. Check your sheet.")
    wb.save()

except Exception as e:
    print(f"\nCRASHED! Here is the error: {e}")

finally:
    pass