# """
# unstructured_lookup.py
# ──────────────────────
# Core daily-data processing logic, exposed as process_excel() so that
# app.py (Unstructured mode) can call it without launching a second GUI.

# Expected call from app.py:
#     import unstructured_lookup as lookup
#     lookup.process_excel(
#         filepath   = "path/to/main.xlsx",
#         date_str   = "DD-MM-YYYY",          # from the date entry
#         online_ro  = "path/to/ro.xlsx",     # optional – None if not selected
#         upliftment = "path/to/ul.xlsx",     # optional – None if not selected
#         nozzle     = "path/to/ns.xlsx",     # optional – None if not selected
#         logger     = gui_log,
#     )
# """

# import calendar
# import re
# from datetime import datetime, date as date_type, timedelta

# import xlwings as xw


# # ──────────────────────────────────────────────────────────────────────────────
# # Utility helpers
# # ──────────────────────────────────────────────────────────────────────────────

# MONTH_MAP = {
#     "jan": 1, "feb": 2, "mar": 3, "apr": 4,  "may": 5,  "jun": 6,
#     "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
# }


# def col_letter(n: int) -> str:
#     return xw.utils.col_name(n)


# def parse_header_date(hdr) -> date_type | None:
#     """Try to read *hdr* (datetime / Excel serial / string) as a date."""
#     if isinstance(hdr, datetime):
#         return hdr.date()
#     if isinstance(hdr, (int, float)) and 35000 < hdr < 60000:
#         return (datetime(1899, 12, 30) + timedelta(days=int(hdr))).date()
#     if isinstance(hdr, str):
#         for fmt in ('%d.%m.%Y', '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d',
#                     '%d.%m.%y', '%d-%m-%y'):
#             try:
#                 return datetime.strptime(hdr.strip(), fmt).date()
#             except ValueError:
#                 pass
#     return None


# def safe_round(val):
#     """Round numeric values to the nearest integer; treat blanks as 0."""
#     if val is None or str(val).strip() == '':
#         return 0
#     try:
#         return int(round(float(val)))
#     except (ValueError, TypeError):
#         return val


# # ──────────────────────────────────────────────────────────────────────────────
# # Sub-sheet readers / writers
# # ──────────────────────────────────────────────────────────────────────────────

# def build_upliftment_map(ul_sheet, year: int, month: int, max_days: int):
#     end_col = ul_sheet.range((1, 16384)).end('left').column
#     raw = ul_sheet.range((1, 1), (1, end_col)).value
#     if not isinstance(raw, (list, tuple)):
#         raw = [raw]

#     stp_cols: list[int] = []
#     date_to_col: dict[date_type, int] = {}

#     for idx, hdr in enumerate(raw):
#         col_num = idx + 1
#         if hdr is None:
#             continue
#         d = parse_header_date(hdr)
#         if d is not None:
#             date_to_col[d] = col_num
#         elif isinstance(hdr, str) and 'party' in hdr.lower():
#             stp_cols.append(col_num)

#     day_map: list[tuple[int, int] | None] = []
#     for day in range(1, max_days + 1):
#         target = date_type(year, month, day)
#         if target in date_to_col:
#             data_col = date_to_col[target]
#             preceding = [s for s in stp_cols if s <= data_col]
#             key_col = max(preceding) if preceding else 1
#             day_map.append((key_col, data_col))
#         else:
#             day_map.append(None)

#     return day_map, stp_cols, len(date_to_col)


# def _find_ro_sheet_cols(ro_sheet, date_label: str, logger=print):
#     """
#     Scan row 1 of the Online RO sub-sheet for a column header matching
#     *date_label*.  Returns (ro_col_letter, date_col_letter) or (None, None).

#     The sheet is structured in pairs:
#         [RO Code] [As on DD-MM-YYYY]  [RO Code] [As on DD-MM-YYYY] …
#     so the date column is always one to the right of its paired RO column.
#     """
#     end_col = ro_sheet.range((1, 16384)).end('left').column
#     headers = ro_sheet.range((1, 1), (1, end_col)).value
#     if not isinstance(headers, list):
#         headers = [headers]

#     for idx, hdr in enumerate(headers):
#         if hdr and isinstance(hdr, str) and date_label.lower() in hdr.lower():
#             data_col = idx + 1
#             key_col  = max(1, data_col - 1)
#             logger(f"  [RO] Found '{date_label}' at col "
#                    f"{col_letter(data_col)} (key col {col_letter(key_col)})")
#             return col_letter(key_col), col_letter(data_col)

#     logger(f"  [RO] ⚠ Header '{date_label}' not found in Online RO sheet.")
#     return None, None


# # ── Upliftment ────────────────────────────────────────────────────────────────

# def write_upliftment_to_subsheet(ext_sheet, ul_sheet, target_date_str: str,
#                                   logger=print):
#     last_a = ext_sheet.range('A' + str(ext_sheet.cells.last_cell.row)).end('up').row
#     last_b = ext_sheet.range('B' + str(ext_sheet.cells.last_cell.row)).end('up').row
#     ext_last = max(last_a, last_b)
#     logger(f"  [UL] External rows 2 → {ext_last}")

#     raw = ext_sheet.range(f'A2:B{ext_last}').value
#     if raw is None:
#         logger("  [UL] ⚠ No data found in external sheet.")
#         return None
#     if not isinstance(raw[0], list):
#         raw = [raw]

#     ro_col  = [[row[0]] for row in raw]
#     val_col = [[safe_round(row[1])] for row in raw]
#     data_rows = len(ro_col)

#     if ul_sheet.range('A1').value is None:
#         ul_sheet.range('A1').value = 'Ship-To Party'
#         ul_sheet.range(f'A2:A{data_rows + 1}').value = ro_col
#         logger("  [UL] Wrote Ship-To Party column (A)")
#     else:
#         logger("  [UL] Ship-To Party column already present — skipped")

#     ul_end_col = ul_sheet.range((1, 16384)).end('left').column
#     new_col = 'B' if (ul_end_col == 1 and ul_sheet.range('B1').value is None) \
#               else col_letter(ul_end_col + 1)

#     ul_sheet.range(f'{new_col}1').value = f"'{target_date_str}"
#     ul_sheet.range(f'{new_col}2:{new_col}{data_rows + 1}').value = val_col
#     logger(f"  [UL] Written to column {new_col} with header '{target_date_str}'")
#     return new_col


# # ── Nozzle Sales ──────────────────────────────────────────────────────────────

# def write_nozzle_to_subsheet(ext_sheet, ns_sheet, target_day: int,
#                               month: int, year: int, logger=print):
#     ext_last_row = ext_sheet.range('A' + str(ext_sheet.cells.last_cell.row)).end('up').row

#     ext_last_col = 1
#     for c in range(2, 9):
#         if ext_sheet.range((2, c)).value is not None:
#             ext_last_col = c
#         else:
#             break

#     logger(f"  [NS] External data range: A2 → {col_letter(ext_last_col)}{ext_last_row}")

#     data = ext_sheet.range((2, 1), (ext_last_row, ext_last_col)).value
#     if not isinstance(data, list):
#         data = [[data]]
#     elif not isinstance(data[0], list):
#         data = [[v] for v in data]

#     for r in range(len(data)):
#         for c in range(1, len(data[r])):
#             data[r][c] = safe_round(data[r][c])

#     date_label = f"{target_day:02d}-{month:02d}-{year}"
#     all_headers = [
#         "RO Code",
#         f"MS NS L {date_label}",
#         f"HSD NS L {date_label}",
#         f"Total KL {date_label}",
#     ]
#     col_headers = all_headers[:ext_last_col]

#     ns_r1_last = ns_sheet.range((1, 16384)).end('left').column
#     paste_col  = 1 if (ns_r1_last == 1 and ns_sheet.range('A1').value is None) \
#                  else ns_r1_last + 1

#     for i, hdr in enumerate(col_headers):
#         ns_sheet.range((1, paste_col + i)).value = hdr

#     ns_sheet.range((2, paste_col)).value = data
#     logger(f"  [NS] Written to col {col_letter(paste_col)} "
#            f"with header '{date_label}' ✅")
#     return paste_col


# # ── Online RO ─────────────────────────────────────────────────────────────────

# def write_online_ro_to_subsheet(ext_sheet, ro_sheet, target_day: int,
#                                  month: int, year: int, logger=print):
#     last_row = ext_sheet.range('A' + str(ext_sheet.cells.last_cell.row)).end('up').row
#     logger(f"  [RO] External data range: A2 → B{last_row}")

#     data = ext_sheet.range(f'A2:B{last_row}').value
#     if not isinstance(data, list):
#         data = [[data]]
#     elif not isinstance(data[0], list):
#         data = [[v] for v in data]

#     for r in range(len(data)):
#         if len(data[r]) > 1:
#             val = data[r][1]
#             if isinstance(val, str):
#                 d = parse_header_date(val)
#                 if d:
#                     data[r][1] = d

#     date_label = f"As on {target_day:02d}-{month:02d}-{year}"

#     ro_r1_last = ro_sheet.range((1, 16384)).end('left').column
#     paste_col  = 1 if (ro_r1_last == 1 and ro_sheet.range('A1').value is None) \
#                  else ro_r1_last + 1

#     c1 = col_letter(paste_col)
#     c2 = col_letter(paste_col + 1)

#     logger(f"  [RO] Writing headers to row 1 at cols {c1}:{c2}")
#     ro_sheet.range(f'{c1}1').value = "RO Code"
#     ro_sheet.range(f'{c2}1').value = date_label
#     ro_sheet.range(f'{c1}2').value = data
#     ro_sheet.range(f'{c2}2:{c2}{len(data) + 1}').number_format = 'dd-mm-yyyy'

#     logger(f"  [RO] Written '{date_label}' to cols {c1}:{c2} ✅")
#     return paste_col


# # ──────────────────────────────────────────────────────────────────────────────
# # Main-sheet mapping
# # ──────────────────────────────────────────────────────────────────────────────

# def map_daily_data_to_main(wb, target_day, month, year, max_days, mode,
#                             logger=print):
#     logger(f"Mapping to Main Sheet (mode: {mode.upper()})…")
#     pattern     = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
#     analysis_sh = next(
#         (wb.sheets[s.name] for s in wb.sheets if pattern.match(s.name)), None
#     )
#     if analysis_sh is None:
#         logger("⚠  No Analysis tab found!")
#         return

#     ul_sheet = wb.sheets['Upliftment']
#     last_row = analysis_sh.range('A' + str(analysis_sh.cells.last_cell.row)).end('up').row

#     nil_col = col_letter(8  + (target_day - 1) * 3)
#     ul_col  = col_letter(9  + (target_day - 1) * 3)
#     ns_col  = col_letter(10 + (target_day - 1) * 3)

#     if mode in ("both", "ul"):
#         ul_day_map, _, _ = build_upliftment_map(ul_sheet, year, month, max_days)
#         entry = ul_day_map[target_day - 1]
#         if entry is None:
#             analysis_sh.range(f'{ul_col}2:{ul_col}{last_row}').value = 0
#             logger(f"  [UL] No date column found → wrote 0 in {ul_col}")
#         else:
#             key_c, dat_c = entry
#             k_ltr = col_letter(key_c)
#             d_ltr = col_letter(dat_c)
#             idx   = dat_c - key_c + 1
#             analysis_sh.range(f'{ul_col}2:{ul_col}{last_row}').formula = (
#                 f'=IFNA(VLOOKUP($A2,Upliftment!${k_ltr}:${d_ltr},{idx},0),0)'
#             )
#             logger(f"  [UL] Mapped to col {ul_col}")

#     if mode in ("both", "ns"):
#         ns_start = 1 + (target_day - 1) * 4
#         ns_key   = col_letter(ns_start)
#         ns_data  = col_letter(ns_start + 3)
#         analysis_sh.range(f'{ns_col}2:{ns_col}{last_row}').formula = (
#             f"=IFERROR(VLOOKUP($A2,'Nozzle Sales'!${ns_key}:${ns_data},4,0),0)"
#         )
#         logger(f"  [NS] Mapped to col {ns_col}")

#     if mode in ("both", "ul", "ns"):
#         analysis_sh.range(f'{nil_col}2:{nil_col}{last_row}').formula = (
#             f'=AND({ul_col}2=0,{ns_col}2=0)'
#         )
#         logger(f"  [NIL] Nil condition set in col {nil_col}")


# def map_online_ro_to_main(wb, target_day, month, year, max_days, logger=print):
#     logger("Mapping Online RO to Main Sheet…")
#     pattern     = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
#     analysis_sh = next(
#         (wb.sheets[s.name] for s in wb.sheets if pattern.match(s.name)), None
#     )
#     if not analysis_sh:
#         logger("⚠  No Analysis tab found!")
#         return

#     ro_sheet   = wb.sheets['Online RO']
#     last_row   = analysis_sh.range('A' + str(analysis_sh.cells.last_cell.row)).end('up').row
#     date_label = f"As on {target_day:02d}-{month:02d}-{year}"

#     ro_col_ltr, date_col_ltr = _find_ro_sheet_cols(ro_sheet, date_label, logger)
#     if not ro_col_ltr:
#         logger("  [RO] ⚠ Cannot map — run the sub-sheet writer first.")
#         return

#     formula = (
#         f"=IFNA(TODAY()-VLOOKUP($A2,'Online RO'!"
#         f"${ro_col_ltr}:${date_col_ltr},2,0),\"Non Automated\")"
#     )
#     analysis_sh.range(f'CX2:CX{last_row}').formula = formula
#     logger(f"  [RO] CX mapped — TODAY()-VLOOKUP using "
#            f"'Online RO'!${ro_col_ltr}:${date_col_ltr}")


# def map_month_end_summary(wb, month, year, max_days, logger=print):
#     """Inject final summary formulas into columns CW through DB."""
#     logger("Month-end detected — injecting summary formulas (CW–DB)…")
#     pattern     = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
#     analysis_sh = next(
#         (wb.sheets[s.name] for s in wb.sheets if pattern.match(s.name)), None
#     )
#     if not analysis_sh:
#         logger("⚠  No Analysis tab found!")
#         return

#     last_row = analysis_sh.range('A' + str(analysis_sh.cells.last_cell.row)).end('up').row

#     # CW — Nil-selling check
#     nil_sht = f"Nil selling 01{month:02d}{year}"
#     analysis_sh.range(f'CW2:CW{last_row}').formula = (
#         f'=IF(IFNA(VLOOKUP($A2,\'{nil_sht}\'!A:A,1,0)-1,"No")="No","No","Yes")'
#     )

#     # CX — Automation health vs last day of month RO date
#     ro_sheet   = wb.sheets['Online RO']
#     date_label = f"As on {max_days:02d}-{month:02d}-{year}"
#     ro_col_ltr, date_col_ltr = _find_ro_sheet_cols(ro_sheet, date_label, logger)

#     if ro_col_ltr:
#         ro_range = f"'Online RO'!${ro_col_ltr}:${date_col_ltr}"
#     else:
#         ro_range = "'Online RO'!$Y:$Z"
#         logger("  [SUMMARY] ⚠ Falling back to Y:Z for Online RO range.")

#     analysis_sh.range(f'CX2:CX{last_row}').formula = (
#         f'=IFNA(TODAY()-VLOOKUP($A2,{ro_range},2,0),"Non Automated")'
#     )

#     # CY — Sum of Upliftment
#     ul_sum = "+".join(
#         col_letter(9  + (d - 1) * 3) + "2" for d in range(1, max_days + 1)
#     )
#     analysis_sh.range(f'CY2:CY{last_row}').formula = f'=SUM({ul_sum})'

#     # CZ — Sum of Nozzle Sales
#     ns_sum = "+".join(
#         col_letter(10 + (d - 1) * 3) + "2" for d in range(1, max_days + 1)
#     )
#     analysis_sh.range(f'CZ2:CZ{last_row}').formula = f'=SUM({ns_sum})'

#     # DA — Percentage
#     analysis_sh.range(f'DA2:DA{last_row}').formula = '=IFERROR(ROUND(CY2/CZ2%, 0), 0)'

#     # DB — Count of Nil days
#     last_nil_col = col_letter(8 + (max_days - 1) * 3)
#     analysis_sh.range(f'DB2:DB{last_row}').formula = (
#         f'=COUNTIF(H2:{last_nil_col}2,TRUE)'
#     )
#     logger("  [SUMMARY] Month-end formulas injected ✅")


# # ──────────────────────────────────────────────────────────────────────────────
# # Public entry point  ← called by app.py
# # ──────────────────────────────────────────────────────────────────────────────

# def process_excel(
#     filepath: str,
#     date_str: str,
#     online_ro:  str | None = None,
#     upliftment: str | None = None,
#     nozzle:     str | None = None,
#     logger=print,
# ):
#     """
#     Parameters
#     ----------
#     filepath   : Path to the main Analysis workbook.
#     date_str   : Date string in DD-MM-YYYY format (from app.py's date entry).
#     online_ro  : Path to the external Online RO file, or None to skip.
#     upliftment : Path to the external Upliftment file, or None to skip.
#     nozzle     : Path to the external Nozzle Sales file, or None to skip.
#     logger     : Callable used for progress messages (defaults to print).
#     """
#     # ── Parse date ────────────────────────────────────────────────────────────
#     try:
#         dt = datetime.strptime(date_str.strip(), "%d-%m-%Y")
#     except ValueError:
#         raise ValueError(
#             f"Invalid date format '{date_str}'. Expected DD-MM-YYYY."
#         )

#     target_day      = dt.day
#     month           = dt.month
#     year            = dt.year
#     year_2d         = year % 100
#     max_days        = calendar.monthrange(year, month)[1]
#     target_date_str = f"{target_day:02d}.{month:02d}.{year_2d:02d}"

#     # ── Determine which sub-sheets to update ─────────────────────────────────
#     has_ul = bool(upliftment)
#     has_ns = bool(nozzle)
#     has_ro = bool(online_ro)

#     if has_ul and has_ns:
#         daily_mode = "both"
#     elif has_ul:
#         daily_mode = "ul"
#     elif has_ns:
#         daily_mode = "ns"
#     else:
#         daily_mode = None   # only RO — skip daily-data mapping

#     logger(f"Month    : {month:02d}/{year}  ({max_days} days)")
#     logger(f"Target   : Day {target_day}  →  {target_date_str}")
#     logger(f"Files    : UL={'yes' if has_ul else 'skip'}  "
#            f"NS={'yes' if has_ns else 'skip'}  "
#            f"RO={'yes' if has_ro else 'skip'}")

#     # ── Open workbooks ────────────────────────────────────────────────────────
#     xl_app = xw.App(visible=True, add_book=False)
#     xl_app.screen_updating      = False
#     xl_app.display_alerts       = False
#     xl_app.api.AskToUpdateLinks = False

#     try:
#         logger("Opening Main Analysis File…")
#         wb          = xl_app.books.open(filepath, update_links=False)
#         xl_app.calculation = 'manual'

#         # ── Write to sub-sheets ───────────────────────────────────────────────
#         if has_ul:
#             logger("Opening External Upliftment File…")
#             wb_ext = xl_app.books.open(upliftment, update_links=False)
#             write_upliftment_to_subsheet(
#                 wb_ext.sheets[0], wb.sheets['Upliftment'],
#                 target_date_str, logger,
#             )
#             wb_ext.close()

#         if has_ns:
#             logger("Opening External Nozzle Sales File…")
#             wb_ext = xl_app.books.open(nozzle, update_links=False)
#             write_nozzle_to_subsheet(
#                 wb_ext.sheets[0], wb.sheets['Nozzle Sales'],
#                 target_day, month, year, logger,
#             )
#             wb_ext.close()

#         if has_ro:
#             logger("Opening External Online RO File…")
#             wb_ext = xl_app.books.open(online_ro, update_links=False)
#             write_online_ro_to_subsheet(
#                 wb_ext.sheets[0], wb.sheets['Online RO'],
#                 target_day, month, year, logger,
#             )
#             wb_ext.close()

#         # ── Map to Analysis sheet ─────────────────────────────────────────────
#         if daily_mode:
#             map_daily_data_to_main(
#                 wb, target_day, month, year, max_days, daily_mode, logger
#             )

#         if has_ro:
#             map_online_ro_to_main(wb, target_day, month, year, max_days, logger)

#         # ── Month-end summary (only when processing the last day) ─────────────
#         if target_day == max_days:
#             map_month_end_summary(wb, month, year, max_days, logger)

#         xl_app.calculate()
#         wb.save()
#         logger("✅  Done — workbook saved.")

#     finally:
#         xl_app.screen_updating = True
#         xl_app.calculation     = 'automatic'
#         xl_app.display_alerts  = True

"""
unstructured_lookup.py
──────────────────────
Core daily-data processing logic, exposed as process_excel() so that
app.py (Unstructured mode) can call it without launching a second GUI.

Expected call from app.py:
    import unstructured_lookup as lookup
    lookup.process_excel(
        filepath   = "path/to/main.xlsx",
        date_str   = "DD-MM-YYYY",          # from the date entry
        online_ro  = "path/to/ro.xlsx",     # optional – None if not selected
        upliftment = "path/to/ul.xlsx",     # optional – None if not selected
        nozzle     = "path/to/ns.xlsx",     # optional – None if not selected
        logger     = gui_log,
    )
"""

import calendar
import re
from datetime import datetime, date as date_type, timedelta

import xlwings as xw


# ──────────────────────────────────────────────────────────────────────────────
# Utility helpers
# ──────────────────────────────────────────────────────────────────────────────

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4,  "may": 5,  "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def col_letter(n: int) -> str:
    return xw.utils.col_name(n)


def parse_header_date(hdr) -> date_type | None:
    """Try to read *hdr* (datetime / Excel serial / string) as a date."""
    if isinstance(hdr, datetime):
        return hdr.date()
    if isinstance(hdr, (int, float)) and 35000 < hdr < 60000:
        return (datetime(1899, 12, 30) + timedelta(days=int(hdr))).date()
    if isinstance(hdr, str):
        for fmt in ('%d.%m.%Y', '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d',
                    '%d.%m.%y', '%d-%m-%y'):
            try:
                return datetime.strptime(hdr.strip(), fmt).date()
            except ValueError:
                pass
    return None


def safe_round(val):
    """Round numeric values to the nearest integer; treat blanks as 0."""
    if val is None or str(val).strip() == '':
        return 0
    try:
        return int(round(float(val)))
    except (ValueError, TypeError):
        return val


def clean_ro(val) -> str | None:
    """
    Normalises RO codes so int/float/string all compare identically:
      133530.0 → '133530',  133530 → '133530',  '133530' → '133530'
    Prevents positional misalignment when the external file starts from a
    different RO than the Upliftment sheet — the root cause of random values.
    """
    if val is None:
        return None
    try:
        return str(int(float(val)))
    except (ValueError, TypeError):
        s = str(val).strip()
        return s if s else None


# ──────────────────────────────────────────────────────────────────────────────
# Sub-sheet readers / writers
# ──────────────────────────────────────────────────────────────────────────────

def build_upliftment_map(ul_sheet, year: int, month: int, max_days: int):
    end_col = ul_sheet.range((1, 16384)).end('left').column
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


def _find_ro_sheet_cols(ro_sheet, date_label: str, logger=print):
    """
    Scan row 1 of the Online RO sub-sheet for a column header matching
    *date_label*.  Returns (ro_col_letter, date_col_letter) or (None, None).

    The sheet is structured in pairs:
        [RO Code] [As on DD-MM-YYYY]  [RO Code] [As on DD-MM-YYYY] …
    so the date column is always one to the right of its paired RO column.
    """
    end_col = ro_sheet.range((1, 16384)).end('left').column
    headers = ro_sheet.range((1, 1), (1, end_col)).value
    if not isinstance(headers, list):
        headers = [headers]

    for idx, hdr in enumerate(headers):
        if hdr and isinstance(hdr, str) and date_label.lower() in hdr.lower():
            data_col = idx + 1
            key_col  = max(1, data_col - 1)
            logger(f"  [RO] Found '{date_label}' at col "
                   f"{col_letter(data_col)} (key col {col_letter(key_col)})")
            return col_letter(key_col), col_letter(data_col)

    logger(f"  [RO] ⚠ Header '{date_label}' not found in Online RO sheet.")
    return None, None


# ── Upliftment ────────────────────────────────────────────────────────────────

def write_upliftment_to_subsheet(ext_sheet, ul_sheet, target_date_str: str,
                                  logger=print):
    """
    FIX: Old code pasted values positionally (row2=first value, row3=second…).
    If the external file starts from a different RO than the Upliftment sheet's
    column A, the positional paste misaligns values → wrong ROs get non-zero
    values (e.g. a random 12 appearing where 0 is expected).

    New approach:
      1. Read external file → build a dict {clean_ro: value}
      2. On first run, write column A (Ship-To Party) from the external file.
      3. On every run, read the EXISTING column A from the Upliftment sheet,
         then for each RO look it up in the dict → write matched value or 0.
    This guarantees missing ROs always get 0, never a neighbouring row's value.
    """
    # ── Read external file into dict ─────────────────────────────────────────
    last_a = ext_sheet.range('A' + str(ext_sheet.cells.last_cell.row)).end('up').row
    last_b = ext_sheet.range('B' + str(ext_sheet.cells.last_cell.row)).end('up').row
    ext_last = max(last_a, last_b)
    logger(f"  [UL] External rows 2 → {ext_last}")

    raw = ext_sheet.range(f'A2:B{ext_last}').value
    if raw is None:
        logger("  [UL] ⚠ No data found in external sheet.")
        return None
    if not isinstance(raw[0], list):
        raw = [raw]

    # Build lookup dict  {normalised_ro_string: rounded_value}
    ext_dict: dict[str, int] = {}
    for row in raw:
        key = clean_ro(row[0])
        val = safe_round(row[1]) if row[1] is not None else 0
        if key:
            ext_dict[key] = val

    logger(f"  [UL] External dict: {len(ext_dict)} entries, "
           f"{sum(1 for v in ext_dict.values() if v != 0)} non-zero")

    # ── Write Ship-To Party column A on the very first day ───────────────────
    if ul_sheet.range('A1').value is None:
        ul_sheet.range('A1').value = 'Ship-To Party'
        ro_col = [[row[0]] for row in raw]
        ul_sheet.range(f'A2:A{len(ro_col) + 1}').value = ro_col
        logger("  [UL] Wrote Ship-To Party column (A)")
    else:
        logger("  [UL] Ship-To Party column already present — skipped")

    # ── Read EXISTING column A from the Upliftment sheet ─────────────────────
    ul_last_row = ul_sheet.range(
        'A' + str(ul_sheet.cells.last_cell.row)
    ).end('up').row

    ul_ro_raw = ul_sheet.range(f'A2:A{ul_last_row}').value
    if ul_ro_raw is None:
        logger("  [UL] ⚠ Upliftment column A is empty.")
        return None
    if not isinstance(ul_ro_raw, list):
        ul_ro_raw = [ul_ro_raw]

    # ── Match each RO in column A against the external dict ──────────────────
    # Any RO missing from the external file gets 0 — never a neighbouring value.
    matched = 0
    val_col = []
    for ro_raw in ul_ro_raw:
        key = clean_ro(ro_raw)
        val = ext_dict.get(key, 0)
        if val != 0:
            matched += 1
        val_col.append([val])

    logger(f"  [UL] Matched {matched}/{len(val_col)} ROs "
           f"({len(val_col) - matched} will be 0)")

    # ── Find next empty column and write ─────────────────────────────────────
    ul_end_col = ul_sheet.range((1, 16384)).end('left').column
    new_col = 'B' if (ul_end_col == 1 and ul_sheet.range('B1').value is None) \
              else col_letter(ul_end_col + 1)

    ul_sheet.range(f'{new_col}1').value = f"'{target_date_str}"
    ul_sheet.range(f'{new_col}2:{new_col}{ul_last_row}').value = val_col
    logger(f"  [UL] Written to column {new_col} with header '{target_date_str}'")
    return new_col


# ── Nozzle Sales ──────────────────────────────────────────────────────────────

def write_nozzle_to_subsheet(ext_sheet, ns_sheet, target_day: int,
                              month: int, year: int, logger=print):
    ext_last_row = ext_sheet.range('A' + str(ext_sheet.cells.last_cell.row)).end('up').row

    ext_last_col = 1
    for c in range(2, 9):
        if ext_sheet.range((2, c)).value is not None:
            ext_last_col = c
        else:
            break

    logger(f"  [NS] External data range: A2 → {col_letter(ext_last_col)}{ext_last_row}")

    data = ext_sheet.range((2, 1), (ext_last_row, ext_last_col)).value
    if not isinstance(data, list):
        data = [[data]]
    elif not isinstance(data[0], list):
        data = [[v] for v in data]

    for r in range(len(data)):
        for c in range(1, len(data[r])):
            data[r][c] = safe_round(data[r][c])

    date_label = f"{target_day:02d}-{month:02d}-{year}"
    all_headers = [
        "RO Code",
        f"MS NS L {date_label}",
        f"HSD NS L {date_label}",
        f"Total KL {date_label}",
    ]
    col_headers = all_headers[:ext_last_col]

    ns_r1_last = ns_sheet.range((1, 16384)).end('left').column
    paste_col  = 1 if (ns_r1_last == 1 and ns_sheet.range('A1').value is None) \
                 else ns_r1_last + 1

    for i, hdr in enumerate(col_headers):
        ns_sheet.range((1, paste_col + i)).value = hdr

    ns_sheet.range((2, paste_col)).value = data
    logger(f"  [NS] Written to col {col_letter(paste_col)} "
           f"with header '{date_label}' ✅")
    return paste_col


# ── Online RO ─────────────────────────────────────────────────────────────────

def write_online_ro_to_subsheet(ext_sheet, ro_sheet, target_day: int,
                                 month: int, year: int, logger=print):
    last_row = ext_sheet.range('A' + str(ext_sheet.cells.last_cell.row)).end('up').row
    logger(f"  [RO] External data range: A2 → B{last_row}")

    data = ext_sheet.range(f'A2:B{last_row}').value
    if not isinstance(data, list):
        data = [[data]]
    elif not isinstance(data[0], list):
        data = [[v] for v in data]

    for r in range(len(data)):
        if len(data[r]) > 1:
            val = data[r][1]
            if isinstance(val, str):
                d = parse_header_date(val)
                if d:
                    data[r][1] = d

    date_label = f"As on {target_day:02d}-{month:02d}-{year}"

    ro_r1_last = ro_sheet.range((1, 16384)).end('left').column
    paste_col  = 1 if (ro_r1_last == 1 and ro_sheet.range('A1').value is None) \
                 else ro_r1_last + 1

    c1 = col_letter(paste_col)
    c2 = col_letter(paste_col + 1)

    logger(f"  [RO] Writing headers to row 1 at cols {c1}:{c2}")
    ro_sheet.range(f'{c1}1').value = "RO Code"
    ro_sheet.range(f'{c2}1').value = date_label
    ro_sheet.range(f'{c1}2').value = data
    ro_sheet.range(f'{c2}2:{c2}{len(data) + 1}').number_format = 'dd-mm-yyyy'

    logger(f"  [RO] Written '{date_label}' to cols {c1}:{c2} ✅")
    return paste_col


# ──────────────────────────────────────────────────────────────────────────────
# Main-sheet mapping
# ──────────────────────────────────────────────────────────────────────────────

def map_daily_data_to_main(wb, target_day, month, year, max_days, mode,
                            logger=print):
    logger(f"Mapping to Main Sheet (mode: {mode.upper()})…")
    pattern     = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
    analysis_sh = next(
        (wb.sheets[s.name] for s in wb.sheets if pattern.match(s.name)), None
    )
    if analysis_sh is None:
        logger("⚠  No Analysis tab found!")
        return

    ul_sheet = wb.sheets['Upliftment']
    last_row = analysis_sh.range('A' + str(analysis_sh.cells.last_cell.row)).end('up').row

    nil_col = col_letter(8  + (target_day - 1) * 3)
    ul_col  = col_letter(9  + (target_day - 1) * 3)
    ns_col  = col_letter(10 + (target_day - 1) * 3)

    if mode in ("both", "ul"):
        ul_day_map, _, _ = build_upliftment_map(ul_sheet, year, month, max_days)
        entry = ul_day_map[target_day - 1]
        if entry is None:
            analysis_sh.range(f'{ul_col}2:{ul_col}{last_row}').value = 0
            logger(f"  [UL] No date column found → wrote 0 in {ul_col}")
        else:
            key_c, dat_c = entry
            k_ltr = col_letter(key_c)
            d_ltr = col_letter(dat_c)
            idx   = dat_c - key_c + 1
            analysis_sh.range(f'{ul_col}2:{ul_col}{last_row}').formula = (
                f'=IFNA(VLOOKUP($A2,Upliftment!${k_ltr}:${d_ltr},{idx},0),0)'
            )
            logger(f"  [UL] Mapped to col {ul_col}")

    if mode in ("both", "ns"):
        ns_start = 1 + (target_day - 1) * 4
        ns_key   = col_letter(ns_start)
        ns_data  = col_letter(ns_start + 3)
        analysis_sh.range(f'{ns_col}2:{ns_col}{last_row}').formula = (
            f"=IFERROR(VLOOKUP($A2,'Nozzle Sales'!${ns_key}:${ns_data},4,0),0)"
        )
        logger(f"  [NS] Mapped to col {ns_col}")

    if mode in ("both", "ul", "ns"):
        analysis_sh.range(f'{nil_col}2:{nil_col}{last_row}').formula = (
            f'=AND({ul_col}2=0,{ns_col}2=0)'
        )
        logger(f"  [NIL] Nil condition set in col {nil_col}")


def map_online_ro_to_main(wb, target_day, month, year, max_days, logger=print):
    logger("Mapping Online RO to Main Sheet…")
    pattern     = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
    analysis_sh = next(
        (wb.sheets[s.name] for s in wb.sheets if pattern.match(s.name)), None
    )
    if not analysis_sh:
        logger("⚠  No Analysis tab found!")
        return

    ro_sheet   = wb.sheets['Online RO']
    last_row   = analysis_sh.range('A' + str(analysis_sh.cells.last_cell.row)).end('up').row
    date_label = f"As on {target_day:02d}-{month:02d}-{year}"

    ro_col_ltr, date_col_ltr = _find_ro_sheet_cols(ro_sheet, date_label, logger)
    if not ro_col_ltr:
        logger("  [RO] ⚠ Cannot map — run the sub-sheet writer first.")
        return

    formula = (
        f"=IFNA(TODAY()-VLOOKUP($A2,'Online RO'!"
        f"${ro_col_ltr}:${date_col_ltr},2,0),\"Non Automated\")"
    )
    analysis_sh.range(f'CX2:CX{last_row}').formula = formula
    logger(f"  [RO] CX mapped — TODAY()-VLOOKUP using "
           f"'Online RO'!${ro_col_ltr}:${date_col_ltr}")


def map_month_end_summary(wb, month, year, max_days, logger=print):
    """Inject final summary formulas into columns CW through DB."""
    logger("Month-end detected — injecting summary formulas (CW–DB)…")
    pattern     = re.compile(r"^Analysis-[A-Za-z]{3}'\d{2}$")
    analysis_sh = next(
        (wb.sheets[s.name] for s in wb.sheets if pattern.match(s.name)), None
    )
    if not analysis_sh:
        logger("⚠  No Analysis tab found!")
        return

    last_row = analysis_sh.range('A' + str(analysis_sh.cells.last_cell.row)).end('up').row

    # CW — Nil-selling check
    nil_sht = f"Nil selling 01{month:02d}{year}"
    analysis_sh.range(f'CW2:CW{last_row}').formula = (
        f'=IF(IFNA(VLOOKUP($A2,\'{nil_sht}\'!A:A,1,0)-1,"No")="No","No","Yes")'
    )

    # CX — Automation health vs last day of month RO date
    ro_sheet   = wb.sheets['Online RO']
    date_label = f"As on {max_days:02d}-{month:02d}-{year}"
    ro_col_ltr, date_col_ltr = _find_ro_sheet_cols(ro_sheet, date_label, logger)

    if ro_col_ltr:
        ro_range = f"'Online RO'!${ro_col_ltr}:${date_col_ltr}"
    else:
        ro_range = "'Online RO'!$Y:$Z"
        logger("  [SUMMARY] ⚠ Falling back to Y:Z for Online RO range.")

    analysis_sh.range(f'CX2:CX{last_row}').formula = (
        f'=IFNA(TODAY()-VLOOKUP($A2,{ro_range},2,0),"Non Automated")'
    )

    # CY — Sum of Upliftment
    ul_sum = "+".join(
        col_letter(9  + (d - 1) * 3) + "2" for d in range(1, max_days + 1)
    )
    analysis_sh.range(f'CY2:CY{last_row}').formula = f'=SUM({ul_sum})'

    # CZ — Sum of Nozzle Sales
    ns_sum = "+".join(
        col_letter(10 + (d - 1) * 3) + "2" for d in range(1, max_days + 1)
    )
    analysis_sh.range(f'CZ2:CZ{last_row}').formula = f'=SUM({ns_sum})'

    # DA — Percentage
    analysis_sh.range(f'DA2:DA{last_row}').formula = '=IFERROR(ROUND(CY2/CZ2%, 0), 0)'

    # DB — Count of Nil days
    last_nil_col = col_letter(8 + (max_days - 1) * 3)
    analysis_sh.range(f'DB2:DB{last_row}').formula = (
        f'=COUNTIF(H2:{last_nil_col}2,TRUE)'
    )
    logger("  [SUMMARY] Month-end formulas injected ✅")


# ──────────────────────────────────────────────────────────────────────────────
# Public entry point  ← called by app.py
# ──────────────────────────────────────────────────────────────────────────────

def process_excel(
    filepath: str,
    date_str: str,
    online_ro:  str | None = None,
    upliftment: str | None = None,
    nozzle:     str | None = None,
    logger=print,
):
    """
    Parameters
    ----------
    filepath   : Path to the main Analysis workbook.
    date_str   : Date string in DD-MM-YYYY format (from app.py's date entry).
    online_ro  : Path to the external Online RO file, or None to skip.
    upliftment : Path to the external Upliftment file, or None to skip.
    nozzle     : Path to the external Nozzle Sales file, or None to skip.
    logger     : Callable used for progress messages (defaults to print).
    """
    # ── Parse date ────────────────────────────────────────────────────────────
    try:
        dt = datetime.strptime(date_str.strip(), "%d-%m-%Y")
    except ValueError:
        raise ValueError(
            f"Invalid date format '{date_str}'. Expected DD-MM-YYYY."
        )

    target_day      = dt.day
    month           = dt.month
    year            = dt.year
    year_2d         = year % 100
    max_days        = calendar.monthrange(year, month)[1]
    target_date_str = f"{target_day:02d}.{month:02d}.{year_2d:02d}"

    # ── Determine which sub-sheets to update ─────────────────────────────────
    has_ul = bool(upliftment)
    has_ns = bool(nozzle)
    has_ro = bool(online_ro)

    if has_ul and has_ns:
        daily_mode = "both"
    elif has_ul:
        daily_mode = "ul"
    elif has_ns:
        daily_mode = "ns"
    else:
        daily_mode = None   # only RO — skip daily-data mapping

    logger(f"Month    : {month:02d}/{year}  ({max_days} days)")
    logger(f"Target   : Day {target_day}  →  {target_date_str}")
    logger(f"Files    : UL={'yes' if has_ul else 'skip'}  "
           f"NS={'yes' if has_ns else 'skip'}  "
           f"RO={'yes' if has_ro else 'skip'}")

    # ── Open workbooks ────────────────────────────────────────────────────────
    xl_app = xw.App(visible=True, add_book=False)
    xl_app.screen_updating      = False
    xl_app.display_alerts       = False
    xl_app.api.AskToUpdateLinks = False

    try:
        logger("Opening Main Analysis File…")
        wb          = xl_app.books.open(filepath, update_links=False)
        xl_app.calculation = 'manual'

        # ── Write to sub-sheets ───────────────────────────────────────────────
        if has_ul:
            logger("Opening External Upliftment File…")
            wb_ext = xl_app.books.open(upliftment, update_links=False)
            write_upliftment_to_subsheet(
                wb_ext.sheets[0], wb.sheets['Upliftment'],
                target_date_str, logger,
            )
            wb_ext.close()

        if has_ns:
            logger("Opening External Nozzle Sales File…")
            wb_ext = xl_app.books.open(nozzle, update_links=False)
            write_nozzle_to_subsheet(
                wb_ext.sheets[0], wb.sheets['Nozzle Sales'],
                target_day, month, year, logger,
            )
            wb_ext.close()

        if has_ro:
            logger("Opening External Online RO File…")
            wb_ext = xl_app.books.open(online_ro, update_links=False)
            write_online_ro_to_subsheet(
                wb_ext.sheets[0], wb.sheets['Online RO'],
                target_day, month, year, logger,
            )
            wb_ext.close()

        # ── Map to Analysis sheet ─────────────────────────────────────────────
        if daily_mode:
            map_daily_data_to_main(
                wb, target_day, month, year, max_days, daily_mode, logger
            )

        if has_ro:
            map_online_ro_to_main(wb, target_day, month, year, max_days, logger)

        # ── Month-end summary (only when processing the last day) ─────────────
        if target_day == max_days:
            map_month_end_summary(wb, month, year, max_days, logger)

        xl_app.calculate()
        wb.save()
        logger("✅  Done — workbook saved.")

    finally:
        xl_app.screen_updating = True
        xl_app.calculation     = 'automatic'
        xl_app.display_alerts  = True