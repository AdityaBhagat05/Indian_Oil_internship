

import xlwings as xw
import calendar
import re
from datetime import datetime, date as date_type, timedelta

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4,  "may": 5,  "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


# ──────────────────────────────────────────────────────────────────────────────
def col_letter(n: int) -> str:
    """1-based column number → Excel column letter(s).  1→'A', 27→'AA', …"""
    return xw.utils.col_name(n)


def parse_header_date(hdr) -> date_type | None:
    if isinstance(hdr, datetime):
        return hdr.date()
    if isinstance(hdr, (int, float)) and 35000 < hdr < 60000:
        return (datetime(1899, 12, 30) + timedelta(days=int(hdr))).date()
    if isinstance(hdr, str):
        for fmt in ('%d.%m.%Y', '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
            try:
                return datetime.strptime(hdr.strip(), fmt).date()
            except ValueError:
                pass
    return None


def build_upliftment_map(ul_sheet, year: int, month: int, max_days: int):
    end_col = ul_sheet.range('A1').end('right').column
    raw = ul_sheet.range((1, 1), (1, end_col)).value
    if not isinstance(raw, (list, tuple)):
        raw = [raw]

    stp_cols: list[int] = []
    date_to_col: dict[date_type, int] = {}

    for idx, hdr in enumerate(raw):
        col_num = idx + 1
        if hdr is None:
            continue
        d = parse_header_date(hdr)
        if d is not None:
            date_to_col[d] = col_num
        elif isinstance(hdr, str) and 'party' in hdr.lower():
            stp_cols.append(col_num)

    day_map: list[tuple[int, int] | None] = []
    for day in range(1, max_days + 1):
        target = date_type(year, month, day)
        if target in date_to_col:
            data_col = date_to_col[target]
            preceding = [s for s in stp_cols if s <= data_col]
            key_col = max(preceding) if preceding else 1
            day_map.append((key_col, data_col))
        else:
            day_map.append(None)

    return day_map, stp_cols, len(date_to_col)


# ══════════════════════════════════════════════════════════════════════════════
def process_excel(filepath: str, month_str: str, year_str: str, logger=print):
    """
    Entry point called by app.py.

    Parameters
    ----------
    filepath  : full path to the Excel workbook
    month_str : 3-letter month abbreviation, e.g. "Aug"
    year_str  : 2-digit year string, e.g. "25"
    logger    : callable for status messages (defaults to print)
    """
    # ── Resolve month / year from UI inputs ───────────────────────────────────
    month_key = month_str.strip().lower()[:3]
    if month_key not in MONTH_MAP:
        logger(f"⚠  Unrecognised month '{month_str}'. Use Jan–Dec.")
        return
    month = MONTH_MAP[month_key]

    try:
        year_2d = int(year_str.strip())
    except ValueError:
        logger(f"⚠  Invalid year '{year_str}'. Enter a 2-digit number, e.g. 25.")
        return
    year = 2000 + year_2d if year_2d < 100 else year_2d

    max_days = calendar.monthrange(year, month)[1]
    logger(f"Month    : {month:02d}/{year}  ({max_days} calendar days)")

    # ── Open workbook ─────────────────────────────────────────────────────────
    app = xw.App(visible=True, add_book=False)
    app.screen_updating      = False
    app.display_alerts       = False
    app.api.AskToUpdateLinks = False   # suppress "Update links?" dialog

    try:
        wb = app.books.open(filepath, update_links=False)
        app.calculation = 'manual'     # must be set after a workbook is loaded
        logger(f"Opened   : {wb.name}")

        # ── 1. Find Analysis sheet ─────────────────────────────────────────────
        pattern = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
        sheet   = next((wb.sheets[s.name] for s in wb.sheets
                        if pattern.match(s.name)), None)
        if sheet is None:
            logger("⚠  No 'Analysis-Mon\\'YY' tab found. Check tab names and retry.")
            return

        ul_sheet = wb.sheets['Upliftment']
        ns_sheet = wb.sheets['Nozzle Sales']

        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        if last_row < 2:
            logger("⚠  Column A is empty – paste RO codes and retry.")
            return

        logger(f"Main tab : {sheet.name}")
        logger(f"Last row : {last_row}")

        # ── 2. Build Upliftment day-map ────────────────────────────────────────
        ul_day_map, _, _ = build_upliftment_map(ul_sheet, year, month, max_days)

        # ── 3. Detect latest day with data in Nozzle Sales ────────────────────
        current_day = 0
        for day in range(1, max_days + 1):
            ns_data_col_index = 1 + ((day - 1) * 4) + 3
            val = ns_sheet.range((2, ns_data_col_index)).value
            if val is not None and str(val).strip() != '':
                current_day = day
            else:
                break

        if current_day == 0:
            logger("⚠  No Nozzle Sales data found – nothing to inject.")
            return

        logger(f"Latest day with data : Day {current_day}")
        logger(f"Injecting formulas for Day {current_day} …")

        nil_col = col_letter(8  + (current_day - 1) * 3)
        ul_col  = col_letter(9  + (current_day - 1) * 3)
        ns_col  = col_letter(10 + (current_day - 1) * 3)

        # Upliftment
        entry = ul_day_map[current_day - 1]
        if entry is None:
            sheet.range(f'{ul_col}2:{ul_col}{last_row}').value = 0
            logger(f"  UL Day {current_day} : holiday/missing → wrote 0")
        else:
            key_c, dat_c = entry
            k_ltr = col_letter(key_c)
            d_ltr = col_letter(dat_c)
            idx   = dat_c - key_c + 1
            sheet.range(f'{ul_col}2:{ul_col}{last_row}').formula = (
                f'=IFNA(VLOOKUP($A2,Upliftment!${k_ltr}:${d_ltr},{idx},0),0)'
            )
            logger(f"  UL Day {current_day} : VLOOKUP({k_ltr}:{d_ltr}, idx={idx})")

        # Nozzle Sales
        ns_start_num = 1 + (current_day - 1) * 4
        ns_key   = col_letter(ns_start_num)
        ns_data  = col_letter(ns_start_num + 3)
        sheet.range(f'{ns_col}2:{ns_col}{last_row}').formula = (
            f"=IFERROR(VLOOKUP($A2,'Nozzle Sales'!${ns_key}:${ns_data},4,0),0)"
        )
        logger(f"  NS Day {current_day} : VLOOKUP({ns_key}:{ns_data}, idx=4)")

        # Nil check
        sheet.range(f'{nil_col}2:{nil_col}{last_row}').formula = (
            f'=AND({ul_col}2=0,{ns_col}2=0)'
        )
        logger(f"  Nil col {nil_col} written.")

        # ── 4. Month-end summary (only on the last calendar day) ───────────────
        if current_day == max_days:
            logger("Month-end detected — injecting summary formulas (CW–DB) …")
            nil_sht = f"Nil selling 01{month:02d}{year}"

            sheet.range(f'CW2:CW{last_row}').formula = (
                f"=IF(IFNA(VLOOKUP($A2,'{nil_sht}'!A:A,1,0)-1,\"No\")=\"No\",\"No\",\"Yes\")"
            )
            sheet.range(f'CX2:CX{last_row}').formula = (
                "=IFNA(TODAY()-VLOOKUP($A2,'Online RO'!Y:Z,2,0),\"Non Automated\")"
            )

            ul_sum = "+".join(
                col_letter(9  + (d - 1) * 3) + "2" for d in range(1, max_days + 1)
            )
            sheet.range(f'CY2:CY{last_row}').formula = f'=SUM({ul_sum})'

            ns_sum = "+".join(
                col_letter(10 + (d - 1) * 3) + "2" for d in range(1, max_days + 1)
            )
            sheet.range(f'CZ2:CZ{last_row}').formula = f'=SUM({ns_sum})'

            sheet.range(f'DA2:DA{last_row}').formula = '=IFERROR(CY2/CZ2%,0)'

            last_nil_col = col_letter(8 + (max_days - 1) * 3)
            sheet.range(f'DB2:DB{last_row}').formula = (
                f'=COUNTIF(H2:{last_nil_col}2,TRUE)'
            )
            logger("Summary formulas injected.")

        # ── 5. Calculate, save ────────────────────────────────────────────────
        app.calculate()
        wb.save()
        logger("✅  Done. Workbook saved successfully.")

    finally:
        app.screen_updating = True
        app.calculation     = 'automatic'
        app.display_alerts  = True