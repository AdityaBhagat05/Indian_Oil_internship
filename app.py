


import os
import threading
import customtkinter as ctk
from tkinter import filedialog
import traceback

NAVY       = "#0D1B2A"
NAVY_HOV   = "#162436"
ORANGE     = "#F47C1C"
ORANGE_HOV = "#D96A0E"
SUCCESS    = "#1F7244"
ERROR_RED  = "#DC2626"
WARN_AMBER = "#D97706"

C_PAGE_BG = ("#EEF2F7", "#141414")
C_CARD    = ("#FFFFFF", "#1E1E1E")
C_INPUT   = ("#F9FAFB", "#2A2A2A")
C_BORDER  = ("#E5E7EB", "#333333")
C_TITLE   = ("#0D1B2A", "#F0F4F8")
C_BODY    = ("#374151", "#C9D1D9")
C_MUTED   = ("#9CA3AF", "#6B7280")

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# ── Dialog ─────────────────────────────────────────────────────────────────────

class _AppDialog(ctk.CTkToplevel):
    """Styled modal dialog for errors, warnings, and info messages."""

    _KINDS = {
        "error":   (ERROR_RED,  "✕  Error"),
        "warning": (WARN_AMBER, "⚠  Warning"),
        "info":    (ORANGE,     "ℹ  Information"),
    }

    def __init__(self, parent, kind: str, title: str, message: str):
        super().__init__(parent)
        accent, header = self._KINDS.get(kind, self._KINDS["info"])

        self.title(title)
        self.resizable(False, False)
        self.configure(fg_color=C_CARD)
        self.grab_set()
        self.lift()
        self.focus_force()

        bar = ctk.CTkFrame(self, fg_color=accent, corner_radius=0, height=52)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        ctk.CTkLabel(
            bar, text=header,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#FFFFFF",
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            self, text=message,
            font=ctk.CTkFont(size=13),
            text_color=C_BODY,
            wraplength=360,
            justify="left",
        ).pack(padx=28, pady=(22, 18), anchor="w")

        ctk.CTkButton(
            self,
            text="OK",
            width=110, height=38,
            fg_color=accent, hover_color=NAVY,
            text_color="#FFFFFF",
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=8,
            command=self.destroy,
        ).pack(pady=(0, 22))

        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        dw = self.winfo_reqwidth()
        dh = self.winfo_reqheight()
        x  = px + (pw - dw) // 2
        y  = py + (ph - dh) // 2
        self.geometry(f"+{x}+{y}")

        self.bind("<Return>", lambda _e: self.destroy())
        self.bind("<Escape>", lambda _e: self.destroy())


# ── Helpers ───────────────────────────────────────────────────────────────────

def _badge(parent, number: str) -> ctk.CTkLabel:
    return ctk.CTkLabel(
        parent, text=str(number),
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color="#FFFFFF", fg_color=ORANGE,
        corner_radius=12, width=24, height=24,
    )


def _card(parent, **kw) -> ctk.CTkFrame:
    kw.setdefault("fg_color",      C_CARD)
    kw.setdefault("corner_radius", 14)
    kw.setdefault("border_width",  1)
    kw.setdefault("border_color",  C_BORDER)
    return ctk.CTkFrame(parent, **kw)


def _section_title(parent, text: str, row: int, pad_top: int = 20):
    ctk.CTkLabel(
        parent, text=text,
        font=ctk.CTkFont(size=15, weight="bold"),
        text_color=C_TITLE, anchor="w",
    ).grid(row=row, column=0, columnspan=2, sticky="w",
           padx=28, pady=(pad_top, 6))


# ── App ───────────────────────────────────────────────────────────────────────

class FormulaInjectorApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Excel Formula Injector")
        self.geometry("960x700")
        self.minsize(860, 620)
        self.configure(fg_color=NAVY)

        self.target_file       = None
        self._online_ro_file   = None
        self._upliftment_file  = None
        self._nozzle_file      = None
        self._has_log          = False
        self._dark_mode        = False
        self._mode             = "Monthly"
        self._data_type        = "Structured"

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_content_host()
        self._switch_page("dashboard")

    # ── Dialog helpers ────────────────────────────────────────────────────────

    def _show_error(self, title: str, message: str):
        _AppDialog(self, "error", title, message)

    def _show_warning(self, title: str, message: str):
        _AppDialog(self, "warning", title, message)

    def _show_info(self, title: str, message: str):
        _AppDialog(self, "info", title, message)

    def _thread_error(self, title: str, message: str):
        self.after(0, self._show_error, title, message)

    def _thread_warning(self, title: str, message: str):
        self.after(0, self._show_warning, title, message)

    # ── Sidebar ───────────────────────────────────────────────────────────────

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, fg_color=NAVY, corner_radius=0, width=195)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_columnconfigure(0, weight=1)
        sb.grid_rowconfigure(8, weight=1)

        logo = ctk.CTkFrame(sb, fg_color=ORANGE, corner_radius=35, width=70, height=70)
        logo.grid(row=0, column=0, pady=(32, 28), padx=20)
        logo.grid_propagate(False)
        ctk.CTkLabel(logo, text="EFI", font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="#FFFFFF").place(relx=.5, rely=.5, anchor="center")

        self._nav_btns = {}
        for idx, (key, label) in enumerate([
            ("dashboard", "Dashboard"),
            ("settings",  "Settings"),
            ("help",      "Help"),
            ("about",     "About"),
        ], start=1):
            btn = ctk.CTkButton(
                sb, text=f"  {label}", anchor="w",
                font=ctk.CTkFont(size=14),
                fg_color="transparent", hover_color=NAVY_HOV,
                text_color="#FFFFFF", corner_radius=10,
                height=46, width=165,
                command=lambda k=key: self._switch_page(k),
            )
            btn.grid(row=idx, column=0, padx=15, pady=3)
            self._nav_btns[key] = btn

    def _set_active_nav(self, key: str):
        for k, btn in self._nav_btns.items():
            if k == key:
                btn.configure(fg_color=ORANGE,        hover_color=ORANGE_HOV)
            else:
                btn.configure(fg_color="transparent", hover_color=NAVY_HOV)

    # ── Content host ──────────────────────────────────────────────────────────

    def _build_content_host(self):
        host = ctk.CTkFrame(self, fg_color=C_PAGE_BG, corner_radius=0)
        host.grid(row=0, column=1, sticky="nsew")
        host.grid_columnconfigure(0, weight=1)
        host.grid_rowconfigure(0, weight=1)
        self._host = host

        self._pages = {
            "dashboard": self._build_dashboard(host),
            "settings":  self._build_settings(host),
            "help":      self._build_help(host),
            "about":     self._build_about(host),
        }

    def _switch_page(self, key: str):
        self._current_page = key
        for k, page in self._pages.items():
            if k == key:
                page.grid(row=0, column=0, sticky="nsew")
                page.tkraise()
            else:
                page.grid_forget()
        self._set_active_nav(key)

    # =========================================================================
    #  DASHBOARD
    # =========================================================================

    def _build_dashboard(self, parent) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(
            parent, fg_color=C_PAGE_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_MUTED,
        )
        page.grid_columnconfigure(0, weight=1)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(page, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=32, pady=(30, 0))
        hdr.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(hdr, text="Excel Formula Injector",
                     font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=C_TITLE, anchor="w",
                     ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(hdr, text="Automate. Inject. Simplify.",
                     font=ctk.CTkFont(size=13), text_color=C_MUTED, anchor="w",
                     ).grid(row=1, column=0, sticky="w", pady=(2, 0))
        ctk.CTkFrame(hdr, fg_color=ORANGE, height=3, width=44, corner_radius=2,
                     ).grid(row=2, column=0, sticky="w", pady=(8, 0))

        # ── Card 1 — Data type selector ───────────────────────────────────────
        c_dtype = _card(page)
        c_dtype.grid(row=1, column=0, sticky="ew", padx=32, pady=(24, 0))
        c_dtype.grid_columnconfigure(0, weight=1)

        dh = ctk.CTkFrame(c_dtype, fg_color="transparent")
        dh.grid(row=0, column=0, sticky="w", padx=22, pady=(18, 12))
        _badge(dh, "1").grid(row=0, column=0, padx=(0, 10))
        ctk.CTkLabel(dh, text="Select Data Type",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=1)

        self._dtype_desc = ctk.CTkLabel(
            c_dtype,
            text="For data already structured in the master file.",
            font=ctk.CTkFont(size=12), text_color=C_BODY,
            anchor="w", wraplength=700, justify="left",
        )
        self._dtype_desc.grid(row=1, column=0, sticky="w", padx=22, pady=(0, 10))

        self._dtype_seg = ctk.CTkSegmentedButton(
            c_dtype,
            values=["Structured", "Unstructured"],
            selected_color=ORANGE,
            selected_hover_color=ORANGE_HOV,
            unselected_color=C_INPUT,
            unselected_hover_color=C_BORDER,
            text_color=C_TITLE,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=46,
            width=340,
            corner_radius=10,
            command=self._on_dtype_change,
        )
        self._dtype_seg.set("Structured")
        self._dtype_seg.grid(row=2, column=0, sticky="w", padx=22, pady=(0, 18))

        # ── Card 2 — File picker ──────────────────────────────────────────────
        c1 = _card(page)
        c1.grid(row=2, column=0, sticky="ew", padx=32, pady=(14, 0))
        c1.grid_columnconfigure(0, weight=1)

        sh = ctk.CTkFrame(c1, fg_color="transparent")
        sh.grid(row=0, column=0, sticky="w", padx=22, pady=(18, 10))
        _badge(sh, "2").grid(row=0, column=0, padx=(0, 10))
        ctk.CTkLabel(sh, text="Select Target Excel File",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=1)

        drop = ctk.CTkFrame(c1, fg_color=C_INPUT, corner_radius=10,
                            border_width=1, border_color=C_BORDER)
        drop.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 18))
        drop.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(drop, text=" X ", font=ctk.CTkFont(size=14, weight="bold"),
                     fg_color=SUCCESS, text_color="#FFFFFF",
                     corner_radius=5, width=34, height=34,
                     ).grid(row=0, column=0, padx=(16, 14), pady=14)

        ftxt = ctk.CTkFrame(drop, fg_color="transparent")
        ftxt.grid(row=0, column=1, sticky="w")
        self._file_main = ctk.CTkLabel(ftxt, text="Choose an Excel file",
                                       font=ctk.CTkFont(size=13, weight="bold"),
                                       text_color=C_TITLE, anchor="w")
        self._file_main.grid(row=0, column=0, sticky="w")
        self._file_sub = ctk.CTkLabel(ftxt, text="Click to browse or drag & drop",
                                      font=ctk.CTkFont(size=11),
                                      text_color=C_MUTED, anchor="w")
        self._file_sub.grid(row=1, column=0, sticky="w")

        ctk.CTkButton(drop, text="Browse", fg_color=NAVY, hover_color=NAVY_HOV,
                      font=ctk.CTkFont(size=13, weight="bold"), text_color="#FFFFFF",
                      corner_radius=8, width=95, height=38,
                      command=self.pick_file,
                      ).grid(row=0, column=2, padx=16, pady=14)

        # ── Card 3a — Injection mode (Structured only) ────────────────────────
        self._c_mode = _card(page)
        self._c_mode.grid(row=3, column=0, sticky="ew", padx=32, pady=(14, 0))
        self._c_mode.grid_columnconfigure(0, weight=1)

        mh = ctk.CTkFrame(self._c_mode, fg_color="transparent")
        mh.grid(row=0, column=0, sticky="w", padx=22, pady=(18, 12))
        _badge(mh, "3").grid(row=0, column=0, padx=(0, 10))
        ctk.CTkLabel(mh, text="Select Injection Mode",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=1)

        self._mode_seg = ctk.CTkSegmentedButton(
            self._c_mode,
            values=["Daily", "Monthly"],
            selected_color=ORANGE,
            selected_hover_color=ORANGE_HOV,
            unselected_color=C_INPUT,
            unselected_hover_color=C_BORDER,
            text_color=C_TITLE,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=46,
            width=320,
            corner_radius=10,
            command=self._on_mode_change,
        )
        self._mode_seg.set("Monthly")
        self._mode_seg.grid(row=1, column=0, sticky="w", padx=22, pady=(0, 18))

        # ── Card 3b — Month + Year (Structured only) ──────────────────────────
        self._c_monthyear = _card(page)
        self._c_monthyear.grid(row=4, column=0, sticky="ew", padx=32, pady=(14, 0))
        self._c_monthyear.grid_columnconfigure(0, weight=1)
        self._c_monthyear.grid_columnconfigure(1, weight=1)

        mf = ctk.CTkFrame(self._c_monthyear, fg_color="transparent")
        mf.grid(row=0, column=0, sticky="ew", padx=22, pady=20)
        mhdr2 = ctk.CTkFrame(mf, fg_color="transparent")
        mhdr2.grid(row=0, column=0, sticky="w", pady=(0, 8))
        _badge(mhdr2, "4").grid(row=0, column=0, padx=(0, 8))
        ctk.CTkLabel(mhdr2, text="Month",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=1)

        months = ["Jan","Feb","Mar","Apr","May","Jun",
                  "Jul","Aug","Sep","Oct","Nov","Dec"]
        self.month_var = ctk.StringVar(value="Aug")
        self.month_combo = ctk.CTkComboBox(
            mf, values=months, variable=self.month_var,
            width=200, height=46, corner_radius=8,
            border_color=ORANGE, button_color=ORANGE,
            button_hover_color=ORANGE_HOV,
            fg_color=C_INPUT, text_color=C_TITLE,
            font=ctk.CTkFont(size=14),
        )
        self.month_combo.grid(row=1, column=0, sticky="w")

        yf = ctk.CTkFrame(self._c_monthyear, fg_color="transparent")
        yf.grid(row=0, column=1, sticky="ew", padx=22, pady=20)
        yh = ctk.CTkFrame(yf, fg_color="transparent")
        yh.grid(row=0, column=0, sticky="w", pady=(0, 8))
        _badge(yh, "5").grid(row=0, column=0, padx=(0, 8))
        ctk.CTkLabel(yh, text="Year (YY)",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=1)

        self.year_entry = ctk.CTkEntry(
            yf, placeholder_text="25", width=200, height=46,
            corner_radius=8, border_color=ORANGE,
            fg_color=C_INPUT, text_color=C_TITLE,
            font=ctk.CTkFont(size=14),
        )
        self.year_entry.insert(0, "25")
        self.year_entry.grid(row=1, column=0, sticky="w")

        # ── Card 3c — Date entry (Unstructured only) ──────────────────────────
        self._c_date = _card(page)
        self._c_date.grid_columnconfigure(0, weight=1)

        dthdr = ctk.CTkFrame(self._c_date, fg_color="transparent")
        dthdr.grid(row=0, column=0, sticky="w", padx=22, pady=(18, 8))
        _badge(dthdr, "3").grid(row=0, column=0, padx=(0, 10))
        ctk.CTkLabel(dthdr, text="Date",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=1)

        ctk.CTkLabel(self._c_date,
                     text="Enter the date for which formulas should be injected.",
                     font=ctk.CTkFont(size=12), text_color=C_BODY,
                     anchor="w", wraplength=700, justify="left",
                     ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 6))

        self.date_entry = ctk.CTkEntry(
            self._c_date, placeholder_text="DD-MM-YYYY",
            width=260, height=46,
            corner_radius=8, border_color=ORANGE,
            fg_color=C_INPUT, text_color=C_TITLE,
            font=ctk.CTkFont(size=14),
        )
        self.date_entry.grid(row=2, column=0, sticky="w", padx=22, pady=(0, 18))

        ctk.CTkLabel(self._c_date,
                     text="Calls unstructured_lookup.py with the selected file, date, "
                          "and any source files you choose below.",
                     font=ctk.CTkFont(size=11), text_color=C_MUTED,
                     anchor="w", wraplength=700, justify="left",
                     ).grid(row=3, column=0, sticky="w", padx=22, pady=(0, 16))

        # ── Card 3d — Source files (Unstructured only) ────────────────────────
        # FIX: files are now optional; at least one must be selected.
        self._c_source_files = _card(page)
        self._c_source_files.grid_columnconfigure(0, weight=1)

        sfhdr = ctk.CTkFrame(self._c_source_files, fg_color="transparent")
        sfhdr.grid(row=0, column=0, sticky="w", padx=22, pady=(18, 4))
        _badge(sfhdr, "4").grid(row=0, column=0, padx=(0, 10))
        ctk.CTkLabel(sfhdr, text="Source Data Files",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=1)

        ctk.CTkLabel(self._c_source_files,
                     # Updated to reflect that files are optional
                     text="Select one or more source files to merge into the master. "
                          "You can skip files you don't need today — "
                          "at least one must be selected.",
                     font=ctk.CTkFont(size=12), text_color=C_BODY,
                     anchor="w", wraplength=700, justify="left",
                     ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 12))

        (self._online_ro_main, self._online_ro_sub) = self._source_row(
            self._c_source_files, row=2, label="Online RO File",
            attr="_online_ro_file",
        )
        (self._upliftment_main, self._upliftment_sub) = self._source_row(
            self._c_source_files, row=3, label="Upliftment File",
            attr="_upliftment_file",
        )
        (self._nozzle_main, self._nozzle_sub) = self._source_row(
            self._c_source_files, row=4, label="Nozzle Sales File",
            attr="_nozzle_file",
        )

        ctk.CTkFrame(self._c_source_files, fg_color="transparent", height=6,
                     ).grid(row=5, column=0)

        # ── Inject button ─────────────────────────────────────────────────────
        self.run_btn = ctk.CTkButton(
            page, text="Inject Formulas",
            fg_color=ORANGE, hover_color=ORANGE_HOV, text_color="#FFFFFF",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=54, corner_radius=12,
            command=self.start_process,
        )
        self.run_btn.grid(row=6, column=0, sticky="ew", padx=32, pady=(16, 0))

        # ── Activity log card ─────────────────────────────────────────────────
        c3 = _card(page)
        c3.grid(row=7, column=0, sticky="ew", padx=32, pady=(16, 32))
        c3.grid_columnconfigure(0, weight=1)

        lh = ctk.CTkFrame(c3, fg_color="transparent")
        lh.grid(row=0, column=0, sticky="w", padx=22, pady=(16, 8))
        ctk.CTkLabel(lh, text="Activity Log",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=0)

        self._empty_frame = ctk.CTkFrame(c3, fg_color="transparent")
        self._empty_frame.grid(row=1, column=0, pady=(8, 24))
        ctk.CTkLabel(self._empty_frame, text="No activity yet",
                     font=ctk.CTkFont(size=13),
                     text_color=C_MUTED).grid(row=0, column=0, pady=(4, 0))

        self.log_output = ctk.CTkTextbox(
            c3, height=170, corner_radius=8,
            fg_color=C_INPUT, border_width=1, border_color=C_BORDER,
            text_color=C_BODY, font=ctk.CTkFont(size=12), state="disabled",
        )
        self.log_output.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 18))
        self.log_output.grid_remove()

        return page

    # ── Data-type selector callback ───────────────────────────────────────────

    def _on_dtype_change(self, value: str):
        self._data_type = value
        if value == "Structured":
            self._dtype_desc.configure(
                text="For data already structured in the master file."
            )
            self._c_mode.grid(row=3, column=0, sticky="ew", padx=32, pady=(14, 0))
            self._c_monthyear.grid(row=4, column=0, sticky="ew", padx=32, pady=(14, 0))
            self._c_date.grid_forget()
            self._c_source_files.grid_forget()
        else:
            self._dtype_desc.configure(
                text="When data is in separate files and/or without headers."
            )
            self._c_mode.grid_forget()
            self._c_monthyear.grid_forget()
            self._c_date.grid(row=3, column=0, sticky="ew", padx=32, pady=(14, 0))
            self._c_source_files.grid(row=4, column=0, sticky="ew",
                                      padx=32, pady=(14, 0))

    # ── Mode selector callback ────────────────────────────────────────────────

    def _on_mode_change(self, value: str):
        self._mode = value

    # =========================================================================
    #  SETTINGS
    # =========================================================================

    def _build_settings(self, parent) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(
            parent, fg_color=C_PAGE_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_MUTED,
        )
        page.grid_columnconfigure(0, weight=1)
        self._page_header(page, "Settings", "Configure app appearance")

        _section_title(page, "Dark Mode", row=1, pad_top=28)
        dm = _card(page)
        dm.grid(row=2, column=0, sticky="ew", padx=32, pady=(0, 4))
        dm.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(dm, text="Switch between light and dark theme",
                     font=ctk.CTkFont(size=13), text_color=C_BODY, anchor="w",
                     ).grid(row=0, column=0, sticky="w", padx=22, pady=18)

        self._dark_switch = ctk.CTkSwitch(
            dm, text="",
            progress_color=ORANGE, button_color="#FFFFFF",
            button_hover_color=C_BORDER,
            command=self._toggle_dark_mode,
        )
        self._dark_switch.grid(row=0, column=1, sticky="e", padx=22, pady=18)
        self._dark_switch.deselect()

        _section_title(page, "Interface Scale", row=3)
        sc = _card(page)
        sc.grid(row=4, column=0, sticky="ew", padx=32, pady=(0, 32))
        sc.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(sc, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 4))
        top.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(top, text="Zoom - scales all widgets and text",
                     font=ctk.CTkFont(size=13), text_color=C_BODY, anchor="w",
                     ).grid(row=0, column=0, sticky="w")
        self._scale_label = ctk.CTkLabel(top, text="100 %",
                                         font=ctk.CTkFont(size=13, weight="bold"),
                                         text_color=ORANGE)
        self._scale_label.grid(row=0, column=1, sticky="e")

        self._scale_slider = ctk.CTkSlider(
            sc, from_=0.8, to=1.4, number_of_steps=6,
            button_color=ORANGE, button_hover_color=ORANGE_HOV,
            progress_color=ORANGE, fg_color=C_BORDER,
            command=self._on_scale_change,
        )
        self._scale_slider.set(1.0)
        self._scale_slider.grid(row=1, column=0, sticky="ew", padx=22, pady=(4, 6))

        ticks = ctk.CTkFrame(sc, fg_color="transparent")
        ticks.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 16))
        for col, pct in enumerate(["80%","90%","100%","110%","120%","130%","140%"]):
            ticks.grid_columnconfigure(col, weight=1)
            ctk.CTkLabel(ticks, text=pct,
                         font=ctk.CTkFont(size=10),
                         text_color=C_MUTED).grid(row=0, column=col)

        return page

    def _toggle_dark_mode(self):
        self._dark_mode = not self._dark_mode
        ctk.set_appearance_mode("dark" if self._dark_mode else "light")

    def _on_scale_change(self, value: float):
        self._scale_label.configure(text=f"{int(round(value * 100))} %")
        ctk.set_widget_scaling(value)
        ctk.set_window_scaling(value)
        self._switch_page(self._current_page)

    # =========================================================================
    #  HELP
    # =========================================================================

    def _build_help(self, parent) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(
            parent, fg_color=C_PAGE_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_MUTED,
        )
        page.grid_columnconfigure(0, weight=1)
        self._page_header(page, "Help", "Everything you need to get started")

        faqs = [
            ("What file types are supported?",
             "The injector supports .xlsx, .xls, and .xlsm Excel files. "
             "Make sure the file is not open in Excel when injecting."),
            ("What does 'Month' and 'Year' refer to?",
             "These values build dynamic lookup references inside your formulas "
             "(e.g. Aug25). They must match the naming convention in your workbook exactly."),
            ("What is the difference between Daily and Monthly mode?",
             "Daily mode calls lookup_daily.py and injects daily formulas. "
             "Monthly mode calls lookup_monthly.py and injects monthly formulas. "
             "Select the appropriate mode before clicking Inject Formulas."),
            ("What is Unstructured mode?",
             "Unstructured mode accepts a single date in DD-MM-YYYY format and calls "
             "unstructured_lookup.py. Select only the source files you need today — "
             "Upliftment, Nozzle Sales, and/or Online RO. All three are optional but "
             "at least one must be provided."),
            ("Can I run Unstructured mode with just one source file?",
             "Yes. Select only the file(s) relevant for today's update. For example, "
             "select only the Online RO file to update just the RO sub-sheet and mapping."),
            ("The process finished but I see no changes - why?",
             "Check that the sheet names in your workbook match what the lookup module "
             "expects (Upliftment, Nozzle Sales, Online RO, Analysis-MonYY). "
             "Enable verbose logging in Settings for a detailed trace."),
            ("Can I run this on a password-protected file?",
             "No. Remove the password protection first, then re-apply it after injection."),
            ("Where are the injected files saved?",
             "The original Excel file is directly modified and saved in its current location."),
            ("How do I report a bug?",
             "Open an issue on the project repository and attach the full activity log."),
        ]

        for idx, (q, a) in enumerate(faqs):
            c = _card(page)
            c.grid(row=idx + 1, column=0, sticky="ew",
                   padx=32, pady=(12 if idx == 0 else 6, 0))
            c.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(c, text=q,
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=C_TITLE, anchor="w",
                         wraplength=680, justify="left",
                         ).grid(row=0, column=0, sticky="w", padx=22, pady=(16, 4))
            ctk.CTkLabel(c, text=a,
                         font=ctk.CTkFont(size=12), text_color=C_BODY,
                         anchor="w", wraplength=680, justify="left",
                         ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 16))

        cc = _card(page,
                   fg_color=("#FFF7ED", "#2A1A00"),
                   border_color=("#FED7AA", "#7C3A00"))
        cc.grid(row=len(faqs) + 1, column=0, sticky="ew", padx=32, pady=(14, 32))
        cc.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(cc,
                     text="Still stuck? Reach out at "
                          "https://github.com/AdityaBhagat05/Indian_Oil_internship",
                     font=ctk.CTkFont(size=13), text_color=ORANGE,
                     ).grid(row=0, column=0, padx=22, pady=16, sticky="w")

        return page

    # =========================================================================
    #  ABOUT
    # =========================================================================

    def _build_about(self, parent) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(
            parent, fg_color=C_PAGE_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_MUTED,
        )
        page.grid_columnconfigure(0, weight=1)
        self._page_header(page, "About", "Excel Formula Injector - version 2.4.1")

        hero = ctk.CTkFrame(page, fg_color=NAVY, corner_radius=14, height=120)
        hero.grid(row=1, column=0, sticky="ew", padx=32, pady=(24, 0))
        hero.grid_propagate(False)
        ctk.CTkLabel(hero, text="Excel Formula Injector",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color="#FFFFFF").place(relx=.5, rely=.4, anchor="center")
        ctk.CTkLabel(hero, text="Built for industrial automation workflows",
                     font=ctk.CTkFont(size=12),
                     text_color="#7A9BBF").place(relx=.5, rely=.68, anchor="center")

        info = _card(page)
        info.grid(row=2, column=0, sticky="ew", padx=32, pady=(16, 0))
        info.grid_columnconfigure(1, weight=1)

        for i, (k, v) in enumerate([
            ("Version",    "1.0.0"),
            ("Build",      "2026-05-13"),
            ("Platform",   "Windows"),
            ("Python",     "3.10+"),
            ("Framework",  "CustomTkinter 5.x"),
            ("License",    "MIT - free to use and modify"),
            ("Repository", "https://github.com/AdityaBhagat05/Indian_Oil_internship"),
        ]):
            ctk.CTkLabel(info, text=k,
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=C_MUTED, anchor="w",
                         ).grid(row=i, column=0, sticky="w", padx=22, pady=10)
            ctk.CTkLabel(info, text=v,
                         font=ctk.CTkFont(size=13), text_color=C_TITLE, anchor="w",
                         ).grid(row=i, column=1, sticky="w", padx=22, pady=10)

        _section_title(page, "Release Notes", row=3)
        cl = _card(page)
        cl.grid(row=4, column=0, sticky="ew", padx=32, pady=(0, 32))
        cl.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(cl, text="",
                     font=ctk.CTkFont(size=12), text_color=C_BODY,
                     anchor="w", justify="left", wraplength=700,
                     ).grid(row=0, column=0, sticky="w", padx=22, pady=18)

        return page

    # ── Shared page header ────────────────────────────────────────────────────

    def _page_header(self, parent, title: str, subtitle: str):
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=32, pady=(30, 0))
        hdr.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(hdr, text=title,
                     font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=C_TITLE, anchor="w",
                     ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(hdr, text=subtitle,
                     font=ctk.CTkFont(size=13), text_color=C_MUTED, anchor="w",
                     ).grid(row=1, column=0, sticky="w", pady=(2, 0))
        ctk.CTkFrame(hdr, fg_color=ORANGE, height=3, width=44, corner_radius=2,
                     ).grid(row=2, column=0, sticky="w", pady=(8, 0))

    # =========================================================================
    #  Business logic
    # =========================================================================

    def log_message(self, message: str):
        if not self._has_log:
            self._has_log = True
            self._empty_frame.grid_remove()
            self.log_output.grid()
        self.log_output.configure(state="normal")
        self.log_output.insert("end", message + "\n")
        self.log_output.see("end")
        self.log_output.configure(state="disabled")

    def pick_file(self):
        try:
            filepath = filedialog.askopenfilename(
                filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
            )
            if filepath:
                if not os.path.isfile(filepath):
                    self._show_error(
                        "File Not Found",
                        f"The selected file no longer exists:\n{filepath}"
                    )
                    return
                self.target_file = filepath
                name = os.path.basename(filepath)
                display = name if len(name) <= 32 else "..." + name[-29:]
                self._file_main.configure(text=display, text_color=SUCCESS)
                self._file_sub.configure(text="File selected - ready to inject",
                                         text_color=SUCCESS)
                self.log_message(f"Selected: {name}")
        except Exception as e:
            self._show_error("File Selection Error",
                             f"An unexpected error occurred while selecting the file:\n{e}")

    def start_process(self):
        # ── Validate: target file ─────────────────────────────────────────────
        if not self.target_file:
            self._show_error(
                "No File Selected",
                "Please select a target Excel file before injecting formulas."
            )
            return

        if not os.path.isfile(self.target_file):
            self._show_error(
                "File Not Found",
                f"The target file could not be found. It may have been moved or deleted:\n"
                f"{self.target_file}"
            )
            self.target_file = None
            self._file_main.configure(text="Choose an Excel file", text_color=C_TITLE)
            self._file_sub.configure(text="Click to browse or drag & drop",
                                     text_color=C_MUTED)
            return

        if self._data_type == "Structured":
            # ── Validate: month / year ────────────────────────────────────────
            month = self.month_combo.get().strip()
            year  = self.year_entry.get().strip()

            if not month:
                self._show_error("Missing Month", "Please select a month before injecting.")
                return

            if not year:
                self._show_error("Missing Year",
                                 "Please enter a two-digit year (e.g. 25) before injecting.")
                return

            if not year.isdigit() or len(year) != 2:
                self._show_error(
                    "Invalid Year",
                    f"'{year}' is not a valid two-digit year.\n"
                    "Please enter exactly two digits (e.g. 25 for 2025)."
                )
                return

            mode_label = self._mode
            self.run_btn.configure(
                state="disabled",
                text=f"Injecting ({mode_label})...",
                fg_color="#9CA3AF", hover_color="#9CA3AF",
            )
            threading.Thread(
                target=self._worker_structured,
                args=(month, year, mode_label), daemon=True
            ).start()

        else:  # Unstructured
            # ── Validate: date ────────────────────────────────────────────────
            date_str = self.date_entry.get().strip()
            if not date_str:
                self._show_error(
                    "Missing Date",
                    "Please enter a date in DD-MM-YYYY format before injecting."
                )
                return

            parts = date_str.split("-")
            if (len(parts) != 3
                    or not all(p.isdigit() for p in parts)
                    or len(parts[0]) != 2
                    or len(parts[1]) != 2
                    or len(parts[2]) != 4):
                self._show_error(
                    "Invalid Date Format",
                    f"'{date_str}' is not a valid date.\n"
                    "Expected format: DD-MM-YYYY (e.g. 15-08-2025)."
                )
                return

            # ── FIX: require at least one source file (not all three) ─────────
            selected = [
                (label, path) for label, path in [
                    ("Online RO File",    self._online_ro_file),
                    ("Upliftment File",   self._upliftment_file),
                    ("Nozzle Sales File", self._nozzle_file),
                ] if path  # None means not chosen
            ]
            if not selected:
                self._show_error(
                    "No Source Files Selected",
                    "Please select at least one source file:\n\n"
                    "  •  Online RO File\n"
                    "  •  Upliftment File\n"
                    "  •  Nozzle Sales File\n\n"
                    "You do not need all three — pick only today's updates."
                )
                return

            # Check that the chosen files still exist on disk
            stale = [
                (label, path) for label, path in selected
                if not os.path.isfile(path)
            ]
            if stale:
                self._show_error(
                    "Source File(s) Not Found",
                    "The following source file(s) could not be found on disk "
                    "(they may have been moved or deleted):\n\n"
                    + "\n".join(f"  •  {lbl}: {os.path.basename(p)}"
                                for lbl, p in stale)
                )
                return

            self.run_btn.configure(
                state="disabled", text="Injecting (Unstructured)...",
                fg_color="#9CA3AF", hover_color="#9CA3AF",
            )
            threading.Thread(
                target=self._worker_unstructured,
                args=(date_str,
                      self._online_ro_file,    # may be None — lookup handles it
                      self._upliftment_file,
                      self._nozzle_file),
                daemon=True,
            ).start()

    # ── Structured worker ─────────────────────────────────────────────────────

    def _worker_structured(self, month: str, year: str, mode: str):
        def gui_log(msg):
            self.after(0, self.log_message, msg)

        gui_log(f"Data type: Structured")
        gui_log(f"Mode: {mode}")

        try:
            if mode == "Daily":
                import lookup_daily as lookup
            else:
                import lookup_monthly as lookup

            lookup.process_excel(self.target_file, month, year, logger=gui_log)
            gui_log("Done.")

        except ImportError:
            module = "lookup_daily" if mode == "Daily" else "lookup_monthly"
            self._thread_warning(
                "Module Not Found",
                f"{module}.py was not found. Running in demo mode.\n\n"
                f"Would inject {mode.lower()} formulas for {month}{year} "
                f"into {os.path.basename(self.target_file)}."
            )
            gui_log(f"[Demo] {module}.py not found — no changes made.")

        except PermissionError:
            self._thread_error(
                "Permission Denied",
                f"Cannot write to the target file. It may be open in Excel or "
                f"another application:\n\n{os.path.basename(self.target_file)}\n\n"
                "Close the file and try again."
            )
            gui_log("Error: Permission denied — file may be open elsewhere.")

        except FileNotFoundError as e:
            self._thread_error(
                "File Not Found",
                f"A required file could not be found during processing:\n\n{e}"
            )
            gui_log(f"Error: File not found — {e}")

        except Exception as e:
            detail = traceback.format_exc()
            self._thread_error(
                "Unexpected Error",
                f"An unexpected error occurred during injection:\n\n"
                f"{type(e).__name__}: {e}\n\n"
                "Check the Activity Log for details."
            )
            gui_log(f"Error: {type(e).__name__}: {e}")
            gui_log(detail)

        finally:
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="Inject Formulas",
                fg_color=ORANGE, hover_color=ORANGE_HOV,
            ))

    # ── Unstructured worker ───────────────────────────────────────────────────
    # FIX: passes None for any file not selected; unstructured_lookup.process_excel
    # skips whichever source is None, mirroring the original script's ul/ns/ro modes.

    def _worker_unstructured(self, date_str: str,
                              online_ro: str | None,
                              upliftment: str | None,
                              nozzle: str | None):
        def gui_log(msg):
            self.after(0, self.log_message, msg)

        gui_log("Data type: Unstructured")
        gui_log(f"Date: {date_str}")
        if online_ro:
            gui_log(f"Online RO:    {os.path.basename(online_ro)}")
        if upliftment:
            gui_log(f"Upliftment:   {os.path.basename(upliftment)}")
        if nozzle:
            gui_log(f"Nozzle Sales: {os.path.basename(nozzle)}")

        try:
            import unstructured_lookup as lookup
            lookup.process_excel(
                self.target_file,
                date_str,
                online_ro=online_ro,
                upliftment=upliftment,
                nozzle=nozzle,
                logger=gui_log,
            )
            gui_log("Done.")

        except ImportError:
            self._thread_warning(
                "Module Not Found",
                "unstructured_lookup.py was not found in the same folder as app.py.\n\n"
                f"Would inject unstructured formulas for {date_str} "
                f"into {os.path.basename(self.target_file)}."
            )
            gui_log("[Demo] unstructured_lookup.py not found — no changes made.")

        except ValueError as e:
            self._thread_error("Invalid Input", str(e))
            gui_log(f"Error: {e}")

        except PermissionError:
            self._thread_error(
                "Permission Denied",
                f"Cannot write to the target file. It may be open in Excel or "
                f"another application:\n\n{os.path.basename(self.target_file)}\n\n"
                "Close the file and try again."
            )
            gui_log("Error: Permission denied — file may be open elsewhere.")

        except FileNotFoundError as e:
            self._thread_error(
                "File Not Found",
                f"A required file could not be found during processing:\n\n{e}"
            )
            gui_log(f"Error: File not found — {e}")

        except Exception as e:
            detail = traceback.format_exc()
            self._thread_error(
                "Unexpected Error",
                f"An unexpected error occurred during injection:\n\n"
                f"{type(e).__name__}: {e}\n\n"
                "Check the Activity Log for details."
            )
            gui_log(f"Error: {type(e).__name__}: {e}")
            gui_log(detail)

        finally:
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="Inject Formulas",
                fg_color=ORANGE, hover_color=ORANGE_HOV,
            ))

    # ── Source-file picker helper ─────────────────────────────────────────────

    def _source_row(self, parent, row: int, label: str,
                    attr: str) -> tuple:
        row_frame = ctk.CTkFrame(parent, fg_color=C_INPUT, corner_radius=10,
                                 border_width=1, border_color=C_BORDER)
        row_frame.grid(row=row, column=0, sticky="ew", padx=22,
                       pady=(0, 10) if row < 4 else (0, 0))
        row_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(row_frame, text=" X ",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     fg_color=NAVY, text_color="#FFFFFF",
                     corner_radius=5, width=30, height=30,
                     ).grid(row=0, column=0, padx=(14, 12), pady=12)

        txt_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
        txt_frame.grid(row=0, column=1, sticky="w")

        # Show "(optional)" in the label so users know they can skip
        main_lbl = ctk.CTkLabel(txt_frame,
                                text=f"{label}  (optional)",
                                font=ctk.CTkFont(size=12, weight="bold"),
                                text_color=C_TITLE, anchor="w")
        main_lbl.grid(row=0, column=0, sticky="w")

        sub_lbl = ctk.CTkLabel(txt_frame, text="No file selected — will be skipped",
                               font=ctk.CTkFont(size=10),
                               text_color=C_MUTED, anchor="w")
        sub_lbl.grid(row=1, column=0, sticky="w")

        ctk.CTkButton(
            row_frame, text="Browse",
            fg_color=NAVY, hover_color=NAVY_HOV,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#FFFFFF", corner_radius=8,
            width=80, height=32,
            command=lambda a=attr, m=main_lbl, s=sub_lbl, lbl=label: (
                self.pick_source_file(a, m, s, lbl)
            ),
        ).grid(row=0, column=2, padx=14, pady=12)

        return main_lbl, sub_lbl

    def pick_source_file(self, attr: str, main_lbl, sub_lbl, label: str):
        try:
            filepath = filedialog.askopenfilename(
                filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
            )
            if filepath:
                if not os.path.isfile(filepath):
                    self._show_error(
                        "File Not Found",
                        f"The selected file could not be found:\n{filepath}"
                    )
                    return
                setattr(self, attr, filepath)
                name = os.path.basename(filepath)
                display = name if len(name) <= 36 else "..." + name[-33:]
                main_lbl.configure(text=display, text_color=SUCCESS)
                sub_lbl.configure(text="Selected ✓", text_color=SUCCESS)
                self.log_message(f"{label}: {name}")
        except Exception as e:
            self._show_error(
                "File Selection Error",
                f"An unexpected error occurred while selecting {label}:\n{e}"
            )


if __name__ == "__main__":
    try:
        app = FormulaInjectorApp()
        app.mainloop()
    except Exception as e:
        import sys
        print(f"Fatal error: {type(e).__name__}: {e}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)