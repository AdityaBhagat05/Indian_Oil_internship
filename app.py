
import os
import threading
import customtkinter as ctk
from tkinter import filedialog
NAVY       = "#0D1B2A"
NAVY_HOV   = "#162436"
ORANGE     = "#F47C1C"
ORANGE_HOV = "#D96A0E"
SUCCESS    = "#1F7244"

C_PAGE_BG = ("#EEF2F7", "#141414")   # scrollable-frame / host background
C_CARD    = ("#FFFFFF", "#1E1E1E")   # card surface
C_INPUT   = ("#F9FAFB", "#2A2A2A")   # entry / textbox background
C_BORDER  = ("#E5E7EB", "#333333")   # card / entry border
C_TITLE   = ("#0D1B2A", "#F0F4F8")   # large heading text
C_BODY    = ("#374151", "#C9D1D9")   # body / label text
C_MUTED   = ("#9CA3AF", "#6B7280")   # placeholder / subtitle text
# ─────────────────────────────────────────────────────────────────────────────

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


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

        self.target_file = None
        self._has_log    = False
        self._dark_mode  = False          # track state ourselves

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_content_host()
        self._switch_page("dashboard")

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
        ctk.CTkLabel(logo, text="🏭", font=ctk.CTkFont(size=26),
                     text_color="#FFFFFF").place(relx=.5, rely=.5, anchor="center")

        self._nav_btns = {}
        for idx, (key, icon, label) in enumerate([
            ("dashboard", "🏠", "Dashboard"),
            ("settings",  "⚙",  "Settings"),
            ("help",      "❓", "Help"),
            ("about",     "ℹ",  "About"),
        ], start=1):
            btn = ctk.CTkButton(
                sb, text=f"  {icon}  {label}", anchor="w",
                font=ctk.CTkFont(size=14),
                fg_color="transparent", hover_color=NAVY_HOV,
                text_color="#FFFFFF", corner_radius=10,
                height=46, width=165,
                command=lambda k=key: self._switch_page(k),
            )
            btn.grid(row=idx, column=0, padx=15, pady=3)
            self._nav_btns[key] = btn

        ctk.CTkLabel(sb, text="⚙  ⚡  🏗", font=ctk.CTkFont(size=18),
                     text_color="#1A3050").grid(row=8, column=0, pady=(0, 24), sticky="s")

    def _set_active_nav(self, key: str):
        for k, btn in self._nav_btns.items():
            if k == key:
                btn.configure(fg_color=ORANGE,       hover_color=ORANGE_HOV)
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

    # ═══════════════════════════════════════════════════════════════════════════
    #  DASHBOARD
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_dashboard(self, parent) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(
            parent, fg_color=C_PAGE_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_MUTED,
        )
        page.grid_columnconfigure(0, weight=1)

        # Header
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

        # Card 1 — file picker
        c1 = _card(page)
        c1.grid(row=1, column=0, sticky="ew", padx=32, pady=(24, 0))
        c1.grid_columnconfigure(0, weight=1)

        sh = ctk.CTkFrame(c1, fg_color="transparent")
        sh.grid(row=0, column=0, sticky="w", padx=22, pady=(18, 10))
        _badge(sh, "1").grid(row=0, column=0, padx=(0, 10))
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

        # Card 2 — Month + Year
        c2 = _card(page)
        c2.grid(row=2, column=0, sticky="ew", padx=32, pady=(14, 0))
        c2.grid_columnconfigure(0, weight=1)
        c2.grid_columnconfigure(1, weight=1)

        # Month
        mf = ctk.CTkFrame(c2, fg_color="transparent")
        mf.grid(row=0, column=0, sticky="ew", padx=22, pady=20)
        mh = ctk.CTkFrame(mf, fg_color="transparent")
        mh.grid(row=0, column=0, sticky="w", pady=(0, 8))
        _badge(mh, "2").grid(row=0, column=0, padx=(0, 8))
        ctk.CTkLabel(mh, text="Month",
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

        # Year
        yf = ctk.CTkFrame(c2, fg_color="transparent")
        yf.grid(row=0, column=1, sticky="ew", padx=22, pady=20)
        yh = ctk.CTkFrame(yf, fg_color="transparent")
        yh.grid(row=0, column=0, sticky="w", pady=(0, 8))
        _badge(yh, "3").grid(row=0, column=0, padx=(0, 8))
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

        # Inject button
        self.run_btn = ctk.CTkButton(
            page, text="🚀   Inject Formulas",
            fg_color=ORANGE, hover_color=ORANGE_HOV, text_color="#FFFFFF",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=54, corner_radius=12,
            command=self.start_process,
        )
        self.run_btn.grid(row=3, column=0, sticky="ew", padx=32, pady=(16, 0))

        # Activity log card
        c3 = _card(page)
        c3.grid(row=4, column=0, sticky="ew", padx=32, pady=(16, 32))
        c3.grid_columnconfigure(0, weight=1)

        lh = ctk.CTkFrame(c3, fg_color="transparent")
        lh.grid(row=0, column=0, sticky="w", padx=22, pady=(16, 8))
        ctk.CTkLabel(lh, text="📋  Activity Log",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=C_TITLE).grid(row=0, column=0)

        # Empty state
        self._empty_frame = ctk.CTkFrame(c3, fg_color="transparent")
        self._empty_frame.grid(row=1, column=0, pady=(8, 24))
        ctk.CTkLabel(self._empty_frame, text="📭",
                     font=ctk.CTkFont(size=36),
                     text_color=C_BORDER).grid(row=0, column=0)
        ctk.CTkLabel(self._empty_frame, text="Your activity will appear here",
                     font=ctk.CTkFont(size=13),
                     text_color=C_MUTED).grid(row=1, column=0, pady=(4, 0))

        self.log_output = ctk.CTkTextbox(
            c3, height=170, corner_radius=8,
            fg_color=C_INPUT, border_width=1, border_color=C_BORDER,
            text_color=C_BODY, font=ctk.CTkFont(size=12), state="disabled",
        )
        self.log_output.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 18))
        self.log_output.grid_remove()

        return page

    # ═══════════════════════════════════════════════════════════════════════════
    #  SETTINGS
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_settings(self, parent) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(
            parent, fg_color=C_PAGE_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_MUTED,
        )
        page.grid_columnconfigure(0, weight=1)
        self._page_header(page, "Settings", "Configure app appearance")

        # Dark mode card
        _section_title(page, "🌙  Dark Mode", row=1, pad_top=28)
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

        # Scale card
        _section_title(page, "🔤  Interface Scale", row=3)
        sc = _card(page)
        sc.grid(row=4, column=0, sticky="ew", padx=32, pady=(0, 32))
        sc.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(sc, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 4))
        top.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(top, text="Zoom — scales all widgets and text",
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
        # Use our own bool flag — never rely on switch.get() returning a string
        self._dark_mode = not self._dark_mode
        ctk.set_appearance_mode("dark" if self._dark_mode else "light")


    def _on_scale_change(self, value: float):
        self._scale_label.configure(text=f"{int(round(value * 100))} %")
        ctk.set_widget_scaling(value)
        ctk.set_window_scaling(value)
        self._switch_page(self._current_page)
    # ═══════════════════════════════════════════════════════════════════════════
    #  HELP
    # ═══════════════════════════════════════════════════════════════════════════

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
            ("The process finished but I see no changes — why?",
             "Check that the sheet names in your workbook match what the lookup module "
             "expects. Enable verbose logging in Settings for a detailed trace."),
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
        ctk.CTkLabel(cc, text="💬  Still stuck? Reach out to us at https://github.com/AdityaBhagat05/Indian_Oil_internship",
                     font=ctk.CTkFont(size=13), text_color=ORANGE,
                     ).grid(row=0, column=0, padx=22, pady=16, sticky="w")

        return page

    # ═══════════════════════════════════════════════════════════════════════════
    #  ABOUT
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_about(self, parent) -> ctk.CTkScrollableFrame:
        page = ctk.CTkScrollableFrame(
            parent, fg_color=C_PAGE_BG, corner_radius=0,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_MUTED,
        )
        page.grid_columnconfigure(0, weight=1)
        self._page_header(page, "About", "Excel Formula Injector — version 2.4.1")

        hero = ctk.CTkFrame(page, fg_color=NAVY, corner_radius=14, height=120)
        hero.grid(row=1, column=0, sticky="ew", padx=32, pady=(24, 0))
        hero.grid_propagate(False)
        ctk.CTkLabel(hero, text="🏭  Excel Formula Injector",
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
            ("License",    "MIT — free to use and modify"),
            ("Repository", "https://github.com/AdityaBhagat05/Indian_Oil_internship"),
        ]):
            ctk.CTkLabel(info, text=k,
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=C_MUTED, anchor="w",
                         ).grid(row=i, column=0, sticky="w", padx=22, pady=10)
            ctk.CTkLabel(info, text=v,
                         font=ctk.CTkFont(size=13), text_color=C_TITLE, anchor="w",
                         ).grid(row=i, column=1, sticky="w", padx=22, pady=10)

        _section_title(page, "📝  Release Notes", row=3)
        cl = _card(page)
        cl.grid(row=4, column=0, sticky="ew", padx=32, pady=(0, 32))
        cl.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(cl,
                     text=(""),
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

    # ═══════════════════════════════════════════════════════════════════════════
    #  Business logic
    # ═══════════════════════════════════════════════════════════════════════════

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
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")]
        )
        if filepath:
            self.target_file = filepath
            name = os.path.basename(filepath)
            display = name if len(name) <= 32 else "…" + name[-29:]
            self._file_main.configure(text=display, text_color=SUCCESS)
            self._file_sub.configure(text="✔  File selected — ready to inject",
                                     text_color=SUCCESS)
            self.log_message(f"📂  Selected: {name}")

    def start_process(self):
        if not self.target_file:
            self.log_message("⚠  Please select a target Excel file first.")
            return
        month = self.month_combo.get().strip()
        year  = self.year_entry.get().strip()
        if not month or not year:
            self.log_message("⚠  Month and Year are required.")
            return
        self.run_btn.configure(state="disabled", text="⏳   Injecting…",
                               fg_color="#9CA3AF", hover_color="#9CA3AF")
        threading.Thread(target=self._worker, args=(month, year), daemon=True).start()

    def _worker(self, month: str, year: str):
        def gui_log(msg):
            self.after(0, self.log_message, msg)
        try:
            import lookup
            lookup.process_excel(self.target_file, month, year, logger=gui_log)
        except ImportError:
            gui_log("⚠  lookup.py not found — running in demo mode.")
            gui_log(f"✅  Would inject formulas for {month}{year} "
                    f"into {os.path.basename(self.target_file)}")
        except Exception as e:
            gui_log(f"❌  Error: {e}")
        finally:
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="🚀   Inject Formulas",
                fg_color=ORANGE, hover_color=ORANGE_HOV,
            ))


if __name__ == "__main__":
    app = FormulaInjectorApp()
    app.mainloop()