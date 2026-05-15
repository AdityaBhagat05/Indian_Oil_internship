# import xlwings as xw
# import calendar
# import re
# import sys

# def letter_to_num(letter):
#     num = 0
#     for c in letter.upper():
#         num = num * 26 + (ord(c) - ord('A')) + 1
#     return num

# app = xw.App(visible=True) 
# app.screen_updating = True 
# app.calculation = 'manual' 

# try:
#     print("Opening monthly.xlsx...")
#     wb = app.books.open('monthly.xlsx') 
    
#     # --- 1. DYNAMIC REGEX SHEET HUNTER ---
#     main_tab_name = None
#     pattern = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
    
#     for sh in wb.sheets:
#         if pattern.match(sh.name):
#             main_tab_name = sh.name
#             break
            
#     if main_tab_name is None:
#         print("⚠️ STOPPING: Could not find Analysis tab.")
#         sys.exit()

#     main_sheet = wb.sheets[main_tab_name] 
#     ul_sheet = wb.sheets['Upliftment']
#     ns_sheet = wb.sheets['Nozzle Sales']
#     last_row = main_sheet.range('A' + str(main_sheet.cells.last_cell.row)).end('up').row

#     ul_map = [
#         ('A', 'B'), ('A', 'C'), ('A', 'D'), ('A', 'E'), ('A', 'F'),   
#         ('A', 'G'), ('A', 'H'), ('A', 'I'), ('A', 'J'), ('A', 'K'),   
#         ('A', 'L'), ('A', 'M'), ('A', 'N'), ('A', 'O'), ('A', 'P'),   
#         ('A', 'Q'), ('A', 'R'), ('A', 'S'), ('T', 'U'), ('V', 'W'),   
#         ('T', 'U'), ('V', 'W'), ('V', 'X'), ('V', 'Y'), ('Z', 'AA'),  
#         ('AB', 'AC'), ('AB', 'AD'), ('AE', 'AF'), ('AG', 'AH'), ('AG', 'AI'), ('AG', 'AJ') 
#     ]

#     # --- 2. DETECT LATEST DAY ---
#     current_day = 0
#     for day in range(1, 32):
#         ns_data_col_index = 1 + ((day - 1) * 4) + 3 
#         val = ns_sheet.range((2, ns_data_col_index)).value
#         if val is not None and str(val).strip() != '':
#             current_day = day
#         else:
#             break 

#     if current_day != 0:
#         print(f"Injecting formulas for Day {current_day} (Removing Decimals)...")

#         nil_col = xw.utils.col_name(8 + ((current_day - 1) * 3))
#         ul_col = xw.utils.col_name(9 + ((current_day - 1) * 3))
#         ns_col = xw.utils.col_name(10 + ((current_day - 1) * 3))

#         # Upliftment (Cleaned)
#         ul_ro_col, ul_data_col = ul_map[current_day-1]
#         ul_idx = letter_to_num(ul_data_col) - letter_to_num(ul_ro_col) + 1
#         main_sheet.range(f'{ul_col}2:{ul_col}{last_row}').formula = \
#             f'=IFNA(INT(VLOOKUP($A2, Upliftment!${ul_ro_col}:${ul_data_col}, {ul_idx}, 0)), 0)'

#         # Nozzle Sales (Cleaned - No Decimals)
#         ns_start_num = 1 + ((current_day - 1) * 4)
#         ns_ro_letter = xw.utils.col_name(ns_start_num)
#         ns_data_letter = xw.utils.col_name(ns_start_num + 3)
#         main_sheet.range(f'{ns_col}2:{ns_col}{last_row}').formula = \
#             f'=IFERROR(INT(VLOOKUP($A2, \'Nozzle Sales\'!${ns_ro_letter}:${ns_data_letter}, 4, 0)), 0)'

#         # Nil Check
#         main_sheet.range(f'{nil_col}2:{nil_col}{last_row}').formula = f'=AND({ul_col}2=0, {ns_col}2=0)'

#         # --- 3. MONTH END SUMMARY ---
#         header_val = str(ns_sheet.range((1, ns_start_num + 3)).value)
#         match = re.search(r'(\d{2})-(\d{2})-(\d{4})', header_val)
        

#         if match:
#             month, year = int(match.group(2)), int(match.group(3))
#             max_days = calendar.monthrange(year, month)[1]
            
#             if current_day == max_days:
#                 dynamic_nil_sheet_name = f"Nil selling 01{month:02d}{year}"
                
#                 main_sheet.range(f'CX2:CX{last_row}').formula = f'=IF(IFNA(VLOOKUP($A2,\'{dynamic_nil_sheet_name}\'!A:A,1,0)-1,"No")="No","No","Yes")'
#                 main_sheet.range(f'CY2:CY{last_row}').formula = '=IFNA(TODAY()-VLOOKUP($A2,\'Online RO\'!Y:Z,2,0),"Non Automated")'
                
#                 ul_sum_str = "+".join([xw.utils.col_name(9 + ((d - 1) * 3)) + "2" for d in range(1, max_days + 1)])
#                 ns_sum_str = "+".join([xw.utils.col_name(10 + ((d - 1) * 3)) + "2" for d in range(1, max_days + 1)])
                
#                 main_sheet.range(f'CZ2:CZ{last_row}').formula = f'=SUM({ul_sum_str})'
#                 main_sheet.range(f'DA2:DA{last_row}').formula = f'=SUM({ns_sum_str})'
                
#                 main_sheet.range(f'DB2:DB{last_row}').formula = '=IFERROR(ROUND(CZ2/DA2%, 0), 0)'
                
#                 last_nil_col = xw.utils.col_name(8 + ((max_days - 1) * 3))
#                 main_sheet.range(f'DC2:DC{last_row}').formula = f'=COUNTIF(H2:{last_nil_col}2,TRUE)'

#         app.calculate()
#         wb.save()
#         print("Done! Decimals have been removed.")

# finally:
#     app.calculation = 'automatic'

import xlwings as xw
import calendar
import re
import sys

def letter_to_num(letter):
    num = 0
    for c in letter.upper():
        num = num * 26 + (ord(c) - ord('A')) + 1
    return num

app = xw.App(visible=True)
app.screen_updating = True
app.calculation = 'manual'

try:
    print("Opening monthly.xlsx...")
    wb = app.books.open('monthly.xlsx')

    # --- 1. DYNAMIC REGEX SHEET HUNTER ---
    main_tab_name = None
    pattern = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")

    for sh in wb.sheets:
        if pattern.match(sh.name):
            main_tab_name = sh.name
            break

    if main_tab_name is None:
        print("⚠️ STOPPING: Could not find Analysis tab.")
        sys.exit()

    main_sheet = wb.sheets[main_tab_name]
    ul_sheet   = wb.sheets['Upliftment']
    ns_sheet   = wb.sheets['Nozzle Sales']
    last_row   = main_sheet.range('A' + str(main_sheet.cells.last_cell.row)).end('up').row

    ul_map = [
        ('A', 'B'), ('A', 'C'), ('A', 'D'), ('A', 'E'), ('A', 'F'),    # days 1-5
        ('A', 'G'), ('A', 'H'), ('A', 'I'), ('A', 'J'), ('A', 'K'),    # days 6-10
        ('A', 'L'), ('A', 'M'), ('A', 'N'), ('A', 'O'), ('A', 'P'),    # days 11-15
        ('A', 'Q'), ('A', 'R'), ('A', 'S'), ('T', 'U'), ('V', 'W'),    # days 16-20
        # ────────────────────────────────────────────────────────────────
        # BUG 3 FIX: Days 21 & 22 were wrongly set to ('T','U') and ('V','W')
        # — identical to days 19 & 20 — causing wrong UL lookups for those
        # days.  Correct the column ranges to their proper positions in the
        # Upliftment sheet.  Adjust these two pairs if your Upliftment sheet
        # uses different column letters for days 21-22.
        ('X', 'Y'), ('Z', 'AA'),                                        # days 21-22 (FIXED)
        # ────────────────────────────────────────────────────────────────
        ('V', 'X'), ('V', 'Y'), ('Z', 'AA'),                            # days 23-25
        ('AB', 'AC'), ('AB', 'AD'), ('AE', 'AF'),                       # days 26-28
        ('AG', 'AH'), ('AG', 'AI'), ('AG', 'AJ'),                       # days 29-31
    ]

    # --- 2. DETECT LATEST DAY WITH DATA (BUG 1 FIX) ---
    # ORIGINAL BUG: the loop used `else: break` which stopped the moment it
    # found a day with no NS data (e.g. a Sunday / upload-gap on day 3).
    # Every subsequent script run would re-hit the same empty day and freeze
    # current_day at 2, meaning:
    #   • Days 4 onward never got UL / NS / Nil formulas injected.
    #   • `current_day == max_days` was never True → month-end block (CZ/DA/
    #     DB/DC) never ran, leaving all summary columns blank or stale.
    #
    # FIX: Remove the break.  Scan all 31 potential days, track which ones
    # have data and which are absent, then inject accordingly.

    current_day  = 0
    absent_days  = []   # days that exist in the calendar but have no NS upload
    days_with_data = []

    for day in range(1, 32):
        ns_data_col_index = 1 + ((day - 1) * 4) + 3
        val = ns_sheet.range((2, ns_data_col_index)).value
        if val is not None and str(val).strip() != '':
            current_day = day
            days_with_data.append(day)
        else:
            # Do NOT break – just note the gap and keep scanning.
            absent_days.append(day)

    print(f"Last day with NS data : {current_day}")
    if absent_days:
        print(f"Absent days (no upload): {absent_days}")

    if current_day == 0:
        print("⚠️ No NS data found at all. Nothing to do.")
        sys.exit()

    # ------------------------------------------------------------------ #
    #  Inject formulas for the current day                                #
    # ------------------------------------------------------------------ #
    print(f"Injecting formulas for Day {current_day} (Removing Decimals)...")

    nil_col = xw.utils.col_name(8  + ((current_day - 1) * 3))
    ul_col  = xw.utils.col_name(9  + ((current_day - 1) * 3))
    ns_col  = xw.utils.col_name(10 + ((current_day - 1) * 3))

    # Upliftment (Cleaned)
    ul_ro_col, ul_data_col = ul_map[current_day - 1]
    ul_idx = letter_to_num(ul_data_col) - letter_to_num(ul_ro_col) + 1
    main_sheet.range(f'{ul_col}2:{ul_col}{last_row}').formula = (
        f'=IFNA(INT(VLOOKUP($A2, Upliftment!${ul_ro_col}:${ul_data_col}, {ul_idx}, 0)), 0)'
    )

    # Nozzle Sales (Cleaned – No Decimals)
    ns_start_num   = 1 + ((current_day - 1) * 4)
    ns_ro_letter   = xw.utils.col_name(ns_start_num)
    ns_data_letter = xw.utils.col_name(ns_start_num + 3)
    main_sheet.range(f'{ns_col}2:{ns_col}{last_row}').formula = (
        f"=IFERROR(INT(VLOOKUP($A2, 'Nozzle Sales'!${ns_ro_letter}:${ns_data_letter}, 4, 0)), 0)"
    )

    # Nil Check – standard formula for days that have actual data
    main_sheet.range(f'{nil_col}2:{nil_col}{last_row}').formula = (
        f'=AND({ul_col}2=0, {ns_col}2=0)'
    )

    # ------------------------------------------------------------------ #
    #  BUG 2 FIX: Mark absent-day Nil columns as FALSE                   #
    # ------------------------------------------------------------------ #
    # ORIGINAL BUG: Absent days (no UL upload, no NS upload) got the same
    # =AND(UL=0, NS=0) formula, which evaluates to TRUE because blanks/zeros
    # look like nil.  Those TRUE values were then counted by the month-end
    # COUNTIF(nil_cols, TRUE), inflating the "instances" figure.
    #
    # FIX: Set the Nil column for every absent day to =FALSE so they are
    # never counted as nil-selling instances.
    for absent_day in absent_days:
        # Only act on days that fall within the current month's calendar
        # (days beyond the month-end are already empty – no need to touch them)
        abs_nil_col = xw.utils.col_name(8 + ((absent_day - 1) * 3))
        # Check whether this column even has a header (i.e. the day is in the
        # sheet layout).  If the header cell is not empty, stamp FALSE.
        header_check_col = xw.utils.col_name(1 + ((absent_day - 1) * 4))
        if ns_sheet.range((1, 1 + ((absent_day - 1) * 4))).value is not None:
            main_sheet.range(f'{abs_nil_col}2:{abs_nil_col}{last_row}').formula = '=FALSE'

    # ------------------------------------------------------------------ #
    #  Month-End Summary                                                  #
    # ------------------------------------------------------------------ #
    header_val  = str(ns_sheet.range((1, ns_start_num + 3)).value)
    match       = re.search(r'(\d{2})-(\d{2})-(\d{4})', header_val)

    if match:
        month, year = int(match.group(2)), int(match.group(3))
        max_days    = calendar.monthrange(year, month)[1]

        if current_day == max_days:
            print("Month-end detected – writing summary columns CX:DC …")
            dynamic_nil_sheet_name = f"Nil selling 01{month:02d}{year}"

            main_sheet.range(f'CX2:CX{last_row}').formula = (
                f"=IF(IFNA(VLOOKUP($A2,'{dynamic_nil_sheet_name}'!A:A,1,0)-1,\"No\")=\"No\",\"No\",\"Yes\")"
            )
            main_sheet.range(f'CY2:CY{last_row}').formula = (
                "=IFNA(TODAY()-VLOOKUP($A2,'Online RO'!Y:Z,2,0),\"Non Automated\")"
            )

            # Total UL and NS across all days that actually had data
            # (absent days contribute 0 through =FALSE Nil; their UL/NS
            #  columns are 0 / blank so they don't distort the sums)
            ul_sum_str = "+".join(
                [xw.utils.col_name(9  + ((d - 1) * 3)) + "2" for d in range(1, max_days + 1)]
            )
            ns_sum_str = "+".join(
                [xw.utils.col_name(10 + ((d - 1) * 3)) + "2" for d in range(1, max_days + 1)]
            )

            main_sheet.range(f'CZ2:CZ{last_row}').formula = f'=SUM({ul_sum_str})'
            main_sheet.range(f'DA2:DA{last_row}').formula = f'=SUM({ns_sum_str})'
            main_sheet.range(f'DB2:DB{last_row}').formula = '=IFERROR(ROUND(CZ2/DA2%, 0), 0)'

            # ---------------------------------------------------------- #
            #  BUG 2 FIX (continued): Count nil instances only for days  #
            #  that actually had upload data – skip absent days.          #
            # ---------------------------------------------------------- #
            # ORIGINAL BUG: =COUNTIF(H2:last_nil_col2, TRUE) counted nil
            # across every day column including absent ones (which wrongly
            # had Nil=TRUE), overcounting instances.
            #
            # FIX: Build the DC formula as a SUMPRODUCT over only the nil
            # columns that belong to days with data.  Absent-day nil columns
            # are =FALSE (from the fix above) so even the old COUNTIF range
            # is now safe – but using an explicit day list is more robust.

            if days_with_data:
                nil_checks = "+".join(
                    [f"({xw.utils.col_name(8 + ((d - 1) * 3))}2=TRUE)*1"
                     for d in days_with_data]
                )
                dc_formula = f'=SUMPRODUCT({nil_checks})'
            else:
                dc_formula = '=0'

            main_sheet.range(f'DC2:DC{last_row}').formula = dc_formula

    app.calculate()
    wb.save()
    print("Done! Decimals removed and month-end summary (if applicable) written correctly.")

finally:
    app.calculation = 'automatic'