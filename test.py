import xlwings as xw

print("Booting up the Lightning Auditor...")
# Keeping it invisible so it runs at maximum speed
app = xw.App(visible=False) 

try:
    # 1. Open both workbooks
    # Make sure both of these files are in the same folder as this Python script!
    print("Opening workbooks...")
    wb_automated = app.books.open('Nil UL & NS Analysis.xlsx')
    wb_manual = app.books.open('main.xlsx')

    # Target the analysis sheets (Assuming they have the same name in both files)
    sheet_auto = wb_automated.sheets['Analysis-Aug\'25']
    sheet_man = wb_manual.sheets['Analysis-Aug\'25']

    # --- SET THE DAY YOU WANT TO AUDIT HERE ---
    day_to_audit = 21 
    # ------------------------------------------
    print(f"\n--- AUDITING DAY {day_to_audit} ---")

    # 2. Calculate the target columns for this specific day
    nil_col = xw.utils.col_name(8 + ((day_to_audit - 1) * 3))
    ul_col = xw.utils.col_name(9 + ((day_to_audit - 1) * 3))
    ns_col = xw.utils.col_name(10 + ((day_to_audit - 1) * 3))

    # Find the last rows (they should be identical, but we check both just in case)
    last_row_auto = sheet_auto.range('A' + str(sheet_auto.cells.last_cell.row)).end('up').row
    last_row_man = sheet_man.range('A' + str(sheet_man.cells.last_cell.row)).end('up').row

    # 3. THE IN-MEMORY BULK READ
    # Pulling all the data into Python's RAM instantly
    print("Reading Manual Data (The Truth)...")
    man_ros = sheet_man.range(f'A2:A{last_row_man}').value
    man_ul = sheet_man.range(f'{ul_col}2:{ul_col}{last_row_man}').value
    man_ns = sheet_man.range(f'{ns_col}2:{ns_col}{last_row_man}').value
    
    print("Reading Automated Data (The Test)...")
    auto_ros = sheet_auto.range(f'A2:A{last_row_auto}').value
    auto_ul = sheet_auto.range(f'{ul_col}2:{ul_col}{last_row_auto}').value
    auto_ns = sheet_auto.range(f'{ns_col}2:{ns_col}{last_row_auto}').value

    # 4. Build a dictionary for the manual data so we can look up by RO Code instantly
    # We treat 'None' (blank cells) as 0 to prevent false-positive errors
    manual_truth = {}
    for i, ro in enumerate(man_ros):
        if ro is not None:
            manual_truth[ro] = {
                'ul': man_ul[i] if man_ul[i] is not None else 0,
                'ns': man_ns[i] if man_ns[i] is not None else 0
            }

    # 5. THE COMPARISON LOOP
    print("Comparing files...\n")
    errors_found = 0

    for i, ro in enumerate(auto_ros):
        if ro is None: continue # Skip blank rows
        
        # Check if this RO even exists in your manual file
        if ro not in manual_truth:
            print(f"WARNING: RO Code {ro} exists in the Automated file but NOT in the Manual file.")
            continue

        # Grab the values we need to compare
        expected_ul = manual_truth[ro]['ul']
        expected_ns = manual_truth[ro]['ns']
        
        actual_ul = auto_ul[i] if auto_ul[i] is not None else 0
        actual_ns = auto_ns[i] if auto_ns[i] is not None else 0

        # Check Upliftment
        if expected_ul != actual_ul:
            print(f"UL ERROR on RO {ro}: Manual says {expected_ul} | Auto says {actual_ul}")
            errors_found += 1
            
        # Check Nozzle Sales
        if expected_ns != actual_ns:
            print(f"NS ERROR on RO {ro}: Manual says {expected_ns} | Auto says {actual_ns}")
            errors_found += 1

    # 6. Final Report
    if errors_found == 0:
        print(f"\n✅ 100% MATCH! Day {day_to_audit} is completely identical across both files.")
    else:
        print(f"\n❌ Found {errors_found} mismatches on Day {day_to_audit}.")

finally:
    # Safely close everything down in the background
    wb_automated.close()
    wb_manual.close()
    app.quit()
    print("Excel processes closed.")