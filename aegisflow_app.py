"""
AegisFlow Desktop Application
==============================
Unified GUI launcher for all AegisFlow tools.

Requirements: pip install customtkinter pandas openpyxl
Build EXE  : run build_exe.bat
"""

import sys
import os
import threading
import io
from tkinter import filedialog
from datetime import datetime

# ── Fix working directory when running as PyInstaller EXE ──────────────────
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))

import customtkinter as ctk

# ── Import tool modules ────────────────────────────────────────────────────
try:
    import mosaic_convert
    import fill_tlf_template as ftt
    import fill_tlf_status as fts
    import extract_programs as ep
    import generate_batch_xml as gbx
    TOOLS_LOADED = True
    IMPORT_ERROR = ""
except ImportError as exc:
    TOOLS_LOADED = False
    IMPORT_ERROR = str(exc)

# ── Appearance defaults ────────────────────────────────────────────────────
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

FONT_FAMILY = "Segoe UI"   # clean Windows system font
FONT_MONO   = "Consolas"   # monospace for log output

TOOL_DEFS = [
    {"id": "mosaic",   "icon": "🔄", "name": "MOSAIC Convert",     "desc": "Convert TiFo CSV → formatted Excel"},
    {"id": "template", "icon": "📊", "name": "Fill TLF Template",  "desc": "Merge MOSAIC output with People Management"},
    {"id": "status",   "icon": "✅", "name": "Fill TLF Status",    "desc": "Merge QC Status from TFL Status file"},
    {"id": "extract",  "icon": "📋", "name": "Extract Programs",   "desc": "Extract SAS %runpgm list from MOSAIC Excel"},
    {"id": "xml",      "icon": "📦", "name": "Generate Batch XML", "desc": "Generate Adobe PDF Builder batch list XML"},
]

# ═══════════════════════════════════════════════════════════════════════════
#  Stdout Redirector
# ═══════════════════════════════════════════════════════════════════════════

class StdoutRedirector:
    """Thread-safe redirect of stdout to a CTkTextbox."""

    def __init__(self, textbox):
        self._box = textbox
        self._orig = sys.__stdout__
        self._buf = []

    def write(self, s: str):
        if s:
            self._buf.append(s)
            self._box.after(0, self._flush_buf)

    def _flush_buf(self):
        if not self._buf:
            return
        text = "".join(self._buf)
        self._buf.clear()
        self._box.configure(state="normal")
        # Colour-code lines containing ✓/❌/⚠
        for line in text.splitlines(keepends=True):
            tag = None
            if any(c in line for c in ("✓", "OK:")):
                tag = "ok"
            elif any(c in line for c in ("❌", "ERROR:", "✗")):
                tag = "err"
            elif any(c in line for c in ("⚠", "WARN:", "Warning")):
                tag = "warn"
            self._box.insert("end", line, tag)
        self._box.see("end")
        self._box.configure(state="disabled")

    def flush(self):
        self._flush_buf()


# ═══════════════════════════════════════════════════════════════════════════
#  Reusable Widgets
# ═══════════════════════════════════════════════════════════════════════════

class FileInputRow(ctk.CTkFrame):
    """Label + path entry + browse/save button in one row."""

    def __init__(self, parent, label: str, placeholder: str = "",
                 save_mode: bool = False,
                 filetypes=None, default_ext: str = ".xlsx",
                 default_name: str = "", **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.save_mode = save_mode
        self._filetypes = filetypes or [("Excel files", "*.xlsx"), ("All files", "*.*")]
        self._ext = default_ext
        self._name = default_name
        self._var = ctk.StringVar()

        label_w = ctk.CTkLabel(self, text=label, width=175, anchor="w",
                               font=(FONT_FAMILY, 12), text_color=("#374151", "#cbd5e1"))
        label_w.grid(row=0, column=0, sticky="w", padx=(0, 10))

        self._entry = ctk.CTkEntry(self, textvariable=self._var,
                                   placeholder_text=placeholder,
                                   height=34, font=(FONT_FAMILY, 12))
        self._entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        btn_txt = "Save As…" if save_mode else "Browse…"
        ctk.CTkButton(self, text=btn_txt, width=96, height=34,
                      font=(FONT_FAMILY, 11),
                      fg_color=("#e2e8f0", "#334155"),
                      text_color=("#374151", "#e2e8f0"),
                      hover_color=("#cbd5e1", "#475569"),
                      command=self._browse).grid(row=0, column=2)

        self.columnconfigure(1, weight=1)

    def _browse(self):
        if self.save_mode:
            p = filedialog.asksaveasfilename(
                filetypes=self._filetypes,
                defaultextension=self._ext,
                initialfile=self._name,
            )
        else:
            p = filedialog.askopenfilename(filetypes=self._filetypes)
        if p:
            self._var.set(p)

    def get(self) -> str:
        return self._var.get().strip()

    def set(self, v: str):
        self._var.set(v)


class SectionCard(ctk.CTkFrame):
    """White/dark card that wraps a group of inputs."""

    def __init__(self, parent, **kw):
        super().__init__(parent,
                         fg_color=("white", "#1e293b"),
                         corner_radius=12, **kw)
        self._inner = ctk.CTkFrame(self, fg_color="transparent")
        self._inner.pack(fill="x", padx=24, pady=20)
        self._inner.columnconfigure(0, weight=1)

    @property
    def inner(self):
        return self._inner


# ═══════════════════════════════════════════════════════════════════════════#  Info Accordion Card
# ═══════════════════════════════════════════════════════════════════════

class InfoCard(ctk.CTkFrame):
    """Collapsible ‘About this Tool’ accordion shown below the panel header."""

    def __init__(self, parent, content: str, **kw):
        super().__init__(parent,
                         fg_color=("#eff6ff", "#0c1f3a"),
                         corner_radius=10, **kw)
        self._expanded = False

        self._toggle_btn = ctk.CTkButton(
            self,
            text="ℹ️  About this Tool  ▼",
            anchor="w",
            height=36,
            font=(FONT_FAMILY, 12),
            fg_color="transparent",
            text_color=("#2563eb", "#60a5fa"),
            hover_color=("#dbeafe", "#1e3a5f"),
            corner_radius=8,
            command=self._toggle,
        )
        self._toggle_btn.pack(fill="x", padx=6, pady=(4, 2))

        self._body = ctk.CTkTextbox(
            self,
            font=(FONT_FAMILY, 11),
            fg_color="transparent",
            text_color=("#374151", "#9ca3af"),
            wrap="word",
            state="normal",
            height=230,
        )
        self._body.insert("1.0", content)
        self._body.configure(state="disabled")
        # hidden by default; user clicks to expand

    def _toggle(self):
        if self._expanded:
            self._body.pack_forget()
            self._toggle_btn.configure(text="ℹ️  About this Tool  ▼")
        else:
            self._body.pack(fill="x", padx=14, pady=(0, 10))
            self._toggle_btn.configure(text="ℹ️  About this Tool  ▲")
        self._expanded = not self._expanded


# ═══════════════════════════════════════════════════════════════════════#  Base Tool Panel
# ═══════════════════════════════════════════════════════════════════════════

class ToolPanel(ctk.CTkFrame):
    title     = ""
    subtitle  = ""
    info_text = ""   # override in subclasses to show InfoCard

    def __init__(self, parent, app, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.app = app
        self._build_header()
        self._build_inputs()
        self._build_run_btn()

    def _build_header(self):
        ctk.CTkLabel(self, text=self.title,
                     font=(FONT_FAMILY, 20, "bold"),
                     text_color=("#111827", "#f1f5f9")).pack(anchor="w", pady=(0, 2))
        ctk.CTkLabel(self, text=self.subtitle,
                     font=(FONT_FAMILY, 12),
                     text_color=("#6b7280", "#9ca3af")).pack(anchor="w", pady=(0, 8))
        if self.info_text:
            InfoCard(self, self.info_text).pack(fill="x", pady=(0, 14))

    def _build_inputs(self):
        """Override in subclass to add SectionCard(s)."""

    def _build_run_btn(self):
        self.run_btn = ctk.CTkButton(
            self,
            text="▶   Run",
            height=44,
            font=(FONT_FAMILY, 15, "bold"),
            command=self._start_run,
        )
        self.run_btn.pack(fill="x", pady=(16, 4))

    # ── execution ──────────────────────────────────────────────────────────
    def _start_run(self):
        self.run_btn.configure(state="disabled", text="⏳  Running…")
        self.app.log_clear()
        self.app.log_timestamp()
        threading.Thread(target=self._run_safe, daemon=True).start()

    def _run_safe(self):
        old = sys.stdout
        sys.stdout = self.app.redirector
        try:
            self._run()
        except Exception as exc:
            import traceback
            print(f"\n❌ Unexpected error: {exc}")
            traceback.print_exc()
        finally:
            sys.stdout = old
            self.run_btn.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶   Run"))

    def _run(self):
        """Override in subclass with actual tool call."""


# ═══════════════════════════════════════════════════════════════════════════
#  Individual Tool Panels
# ═══════════════════════════════════════════════════════════════════════════

class MosaicPanel(ToolPanel):
    title    = "MOSAIC Convert"
    subtitle = "Convert TiFo CSV to formatted Excel with full Index + Original sheets"
    info_text = (
        "📥  Input File\n"
        "  • Format   : CSV (.csv) — exported from TiFo / MOSAIC\n"
        "  • Key columns (case-sensitive): TocNumber, Title, Output, Type,\n"
        "    Programs, Developer, QC Programmer, QC Date, QC Status\n\n"
        "📤  Output File  (Excel .xlsx with 2 sheets)\n"
        "  • \"Index\" sheet    — formatted master TOC with colour-coded rows\n"
        "  • \"Original\" sheet — raw source data preserved as-is\n\n"
        "🎨  Highlight Colours  (\"Index\" sheet)\n"
        "  • Green  (✅) — row has a valid, non-placeholder output filename\n"
        "  • Yellow (⚠️) — output filename is blank or contains \"XXXX\"\n"
        "  • Red    (❌) — duplicate TocNumber detected in source data\n"
        "  • No fill — standard row\n\n"
        "📋  Log Checks\n"
        "  • Missing required columns → lists which ones are absent\n"
        "  • Duplicate TocNumber values → reports all affected rows\n"
        "  • Rows with empty or placeholder Output field\n"
        "  • Final row count and output file path\n\n"
        "⚠  Notes\n"
        "  • If \"Output XLSX\" is left blank, a name like\n"
        "    output_MOSAIC_CONVERT_YYYY-MM-DD.xlsx is auto-created\n"
        "    in the same folder as the input CSV.\n"
        "  • Column names in the CSV must match exactly (case-sensitive)."
    )

    def _build_inputs(self):
        card = SectionCard(self)
        card.pack(fill="x")

        self._csv = FileInputRow(card.inner, "Input CSV File:",
                                 "Select TiFo CSV file…",
                                 filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                                 default_ext=".csv")
        self._csv.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self._out = FileInputRow(card.inner, "Output XLSX (optional):",
                                 "Leave blank to auto-name with date…",
                                 save_mode=True,
                                 default_name="output_MOSAIC_CONVERT.xlsx")
        self._out.grid(row=1, column=0, sticky="ew")

    def _run(self):
        csv = self._csv.get()
        out = self._out.get() or None
        if not csv:
            print("❌ Please select an input CSV file."); return
        if not os.path.exists(csv):
            print(f"❌ File not found: {csv}"); return
        mosaic_convert.mosaic_convert(csv, out)


class TemplatePanel(ToolPanel):
    title    = "Fill TLF Template"
    subtitle = "Merge MOSAIC output with People Management — three-tier cascading match"
    info_text = (
        "📥  Input Files\n"
        "  • MOSAIC XLSX  — output from the MOSAIC Convert tool\n"
        "                   (the \"Index\" sheet is read automatically)\n"
        "  • People Mgmt  — existing People Management workbook to be updated\n\n"
        "📤  Output File  (Excel .xlsx)\n"
        "  • Same structure as People Mgmt, with Programmer / QC Program /\n"
        "    QC Programmer columns populated from MOSAIC data\n\n"
        "🔗  Three-Tier Cascading Match  (applied in order; each tier only\n"
        "    fills rows that the previous tier could not match)\n\n"
        "  Tier 1 — Output Name  (highest priority)\n"
        "    Exact match on the \"Output Name\" column between MOSAIC and\n"
        "    People Mgmt.  Matched rows receive no special highlight.\n\n"
        "  Tier 2 — Program Name  (supplement)\n"
        "    For rows still unmatched after Tier 1, exact match on the\n"
        "    \"Program Name\" column.  Matched rows are highlighted Green.\n\n"
        "  Unmatched — neither column found a match\n"
        "    Programmer / QC fields are left blank and the row is\n"
        "    highlighted Yellow as a manual-review flag.\n\n"
        "🎨  Highlight Colours  (Programmer / QC columns)\n"
        "  • No fill — matched via Tier 1 (Output Name)\n"
        "  • Green   — matched via Tier 2 (Program Name)\n"
        "  • Yellow  — unmatched; needs manual attention\n\n"
        "📋  Log Checks\n"
        "  • Reports Tier 1 match count and Tier 2 supplement count\n"
        "  • Lists all rows that could not be matched with a warning\n"
        "  • Warns if People Mgmt is missing expected column headers\n"
        "  • Final totals: matched rows vs. total rows\n\n"
        "⚠  Notes\n"
        "  • The MOSAIC XLSX must be the direct output of the MOSAIC Convert step.\n"
        "  • Do NOT rename columns in the People Mgmt file between runs.\n"
        "  • Column names are case-sensitive (\"Output Name\", \"Program Name\")."
    )

    def _build_inputs(self):
        card = SectionCard(self)
        card.pack(fill="x")

        self._mosaic = FileInputRow(card.inner, "MOSAIC XLSX File:",
                                   "Select MOSAIC_CONVERT_*.xlsx…")
        self._mosaic.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self._ppl = FileInputRow(card.inner, "People Mgmt XLSX:",
                                 "Select people_management.xlsx…")
        self._ppl.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        self._out = FileInputRow(card.inner, "Output XLSX File:",
                                 "people_management_updated.xlsx",
                                 save_mode=True,
                                 default_name="people_management_updated.xlsx")
        self._out.grid(row=2, column=0, sticky="ew")

    def _run(self):
        mosaic = self._mosaic.get()
        ppl    = self._ppl.get()
        out    = self._out.get()
        for label, val in [("MOSAIC XLSX", mosaic), ("People Mgmt XLSX", ppl), ("Output XLSX", out)]:
            if not val:
                print(f"❌ Please fill in: {label}"); return
        for f in (mosaic, ppl):
            if not os.path.exists(f):
                print(f"❌ File not found: {f}"); return
        ftt.fill_tlf_template(mosaic_file=mosaic, people_file=ppl, output_file=out)


class StatusPanel(ToolPanel):
    title    = "Fill TLF Status"
    subtitle = "Merge QC Status from TFL Status file (Match→Pass, Mismatch→Fail)"
    info_text = (
        "📥  Input Files\n"
        "  • People Mgmt XLSX  — output from the Fill TLF Template step\n"
        "  • TFL Status XLSX   — QC status tracking file from the QC team\n\n"
        "📤  Output File  (Excel .xlsx)\n"
        "  • People Mgmt with a new \"QC Status\" column fully populated\n\n"
        "🔗  Matching Logic\n"
        "  Match key   : Output filename (normalised, case-insensitive)\n"
        "  Match + agree  → \"Pass\"  (green highlight)\n"
        "  Match + differ → \"Fail\"  (red highlight)\n"
        "  No match       → cell left blank (no highlight)\n\n"
        "🎨  Highlight Colours  (QC Status column)\n"
        "  • Green — Pass: programmer status and QC reviewer status agree\n"
        "  • Red   — Fail: statuses differ, manual review required\n"
        "  • No fill — output filename not found in TFL Status file\n\n"
        "📋  Log Checks\n"
        "  • Key column detection in both input files\n"
        "  • Reports unmatched rows from People Mgmt\n"
        "  • Reports TFL Status rows that matched nothing in People Mgmt\n"
        "  • Final summary: Pass count, Fail count, unmatched count\n\n"
        "⚠  Notes\n"
        "  • Run this tool AFTER Fill TLF Template, not before.\n"
        "  • Column names in the TFL Status file must match the expected schema."
    )

    def _build_inputs(self):
        card = SectionCard(self)
        card.pack(fill="x")

        self._ppl = FileInputRow(card.inner, "People Mgmt XLSX:",
                                 "Select people_management.xlsx…")
        self._ppl.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self._stat = FileInputRow(card.inner, "TFL Status XLSX:",
                                  "Select tfl_status.xlsx…")
        self._stat.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        self._out = FileInputRow(card.inner, "Output XLSX File:",
                                 "people_management_with_status.xlsx",
                                 save_mode=True,
                                 default_name="people_management_with_status.xlsx")
        self._out.grid(row=2, column=0, sticky="ew")

    def _run(self):
        ppl  = self._ppl.get()
        stat = self._stat.get()
        out  = self._out.get()
        for label, val in [("People Mgmt XLSX", ppl), ("TFL Status XLSX", stat), ("Output XLSX", out)]:
            if not val:
                print(f"❌ Please fill in: {label}"); return
        for f in (ppl, stat):
            if not os.path.exists(f):
                print(f"❌ File not found: {f}"); return
        fts.fill_tlf_status(people_file=ppl, status_file=stat, output_file=out)


class ExtractPanel(ToolPanel):
    title    = "Extract Programs"
    subtitle = "Extract SAS %runpgm script ordered by first appearance in MOSAIC Excel"
    info_text = (
        "📥  Input File\n"
        "  • MOSAIC XLSX — output from the MOSAIC Convert tool\n"
        "    (reads the \"Index\" sheet; looks for a \"Programs\" column)\n\n"
        "📤  Output File  (.txt)\n"
        "  • A plain-text SAS script with one %runpgm() call per program,\n"
        "    ordered by first appearance (TocNumber order) in the spreadsheet\n"
        "  • Duplicate program names are automatically de-duplicated\n\n"
        "📋  Log Checks\n"
        "  • Warns if the \"Programs\" column is missing or entirely empty\n"
        "  • Reports total unique programs extracted\n"
        "  • Lists any rows where the Programs cell is blank\n"
        "  • Final line count in the generated script\n\n"
        "⚠  Notes\n"
        "  • Each Programs cell may contain multiple programme names separated\n"
        "    by commas, semicolons, or line breaks — all formats are handled.\n"
        "  • The output .txt file is UTF-8 encoded; open in SAS EG or any\n"
        "    plain-text editor."
    )

    def _build_inputs(self):
        card = SectionCard(self)
        card.pack(fill="x")

        self._xl = FileInputRow(card.inner, "MOSAIC XLSX File:",
                                "Select MOSAIC_CONVERT_*.xlsx…")
        self._xl.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self._out = FileInputRow(card.inner, "Output TXT File:",
                                 "run_all_pgm_generated.txt",
                                 save_mode=True,
                                 filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                                 default_ext=".txt",
                                 default_name="run_all_pgm_generated.txt")
        self._out.grid(row=1, column=0, sticky="ew")

    def _run(self):
        xl  = self._xl.get()
        out = self._out.get()
        for label, val in [("MOSAIC XLSX", xl), ("Output TXT", out)]:
            if not val:
                print(f"❌ Please fill in: {label}"); return
        if not os.path.exists(xl):
            print(f"❌ File not found: {xl}"); return
        ep.main(input_file=xl, output_file=out)


class XMLPanel(ToolPanel):
    title    = "Generate Batch XML"
    subtitle = "Generate Adobe PDF Builder batch list XML from Excel/CSV"
    info_text = (
        "📥  Input File\n"
        "  • Excel (.xlsx / .xls) or CSV (.csv)\n"
        "  • Required: a column containing PDF output filenames\n"
        "    (auto-detected from common names: Output, Filename, File)\n\n"
        "📤  Output File  (.xml)\n"
        "  • Adobe PDF Builder batch list XML\n"
        "  • One <batchitem> entry per row, with sequential page numbers\n\n"
        "⚙️  Parameters\n"
        "  • Header Text       — label in the batch PDF cover header\n"
        "  • Output PDF Name   — base name of the assembled PDF (no spaces)\n"
        "  • File Location     — SAS-style path to PDF source files\n"
        "                        e.g.  root/cdar/.../output/\n"
        "  • Start Page Number — page number assigned to the first item\n\n"
        "📋  Log Checks\n"
        "  • Skips rows where Output filename is blank\n"
        "  • Warns if Output PDF Name contains spaces (invalid for Adobe)\n"
        "  • Validates Start Page Number is a positive integer\n"
        "  • Reports total items written to XML\n\n"
        "⚠  Notes\n"
        "  • Open the output XML in Acrobat Pro via\n"
        "    Document > Combine Files (Batch PDF Builder plugin).\n"
        "  • File Location must use forward slashes (/), not backslashes."
    )

    def _build_inputs(self):
        # ── File inputs ────────────────────────────────────────────────────
        card1 = SectionCard(self)
        card1.pack(fill="x", pady=(0, 12))

        self._xl = FileInputRow(card1.inner, "Input Excel / CSV:",
                                "Select Excel or CSV file…",
                                filetypes=[
                                    ("Excel/CSV", "*.xlsx *.xls *.csv"),
                                    ("All files", "*.*"),
                                ])
        self._xl.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self._out = FileInputRow(card1.inner, "Output XML File:",
                                 "batch_list.xml",
                                 save_mode=True,
                                 filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
                                 default_ext=".xml",
                                 default_name="batch_list.xml")
        self._out.grid(row=1, column=0, sticky="ew")

        # ── Parameter inputs ───────────────────────────────────────────────
        card2 = SectionCard(self)
        card2.pack(fill="x")
        g = card2.inner
        g.columnconfigure(1, weight=1)

        fields = [
            ("Header Text:",        "AZD0901 CSR DR2 Batch 1 Listings"),
            ("Output PDF Name:",    "AZD0901_CSR_DR2_Batch1_Listings"),
            ("File Location:",      "root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/"),
            ("Start Page Number:",  "2"),
        ]
        self._param_entries = []
        for i, (lbl, ph) in enumerate(fields):
            ctk.CTkLabel(g, text=lbl, width=175, anchor="w",
                         font=(FONT_FAMILY, 12), text_color=("#374151", "#cbd5e1")
                         ).grid(row=i, column=0, sticky="w", padx=(0, 10), pady=4)
            entry = ctk.CTkEntry(g, placeholder_text=ph, height=34, font=(FONT_FAMILY, 12))
            entry.grid(row=i, column=1, sticky="ew", pady=4)
            self._param_entries.append((entry, ph))

    def _run(self):
        xl  = self._xl.get()
        out = self._out.get()
        for label, val in [("Input Excel/CSV", xl), ("Output XML", out)]:
            if not val:
                print(f"❌ Please fill in: {label}"); return
        if not os.path.exists(xl):
            print(f"❌ File not found: {xl}"); return

        header_text  = self._param_entries[0][0].get().strip() or self._param_entries[0][1]
        pdf_name     = self._param_entries[1][0].get().strip() or self._param_entries[1][1]
        file_loc     = self._param_entries[2][0].get().strip() or self._param_entries[2][1]
        start_raw    = self._param_entries[3][0].get().strip()

        try:
            start_num = int(start_raw) if start_raw else 2
        except ValueError:
            print("⚠ Start page is not a number — defaulting to 2")
            start_num = 2

        if " " in pdf_name:
            print(f"❌ Output PDF Name cannot contain spaces. Suggested: {pdf_name.replace(' ', '_')}")
            return

        gbx.run_from_gui(
            input_file=xl,
            header_text=header_text,
            output_filename=pdf_name,
            file_location=file_loc,
            start_number=start_num,
            output_path=out,
        )


# ═══════════════════════════════════════════════════════════════════════════
#  Main Application Window
# ═══════════════════════════════════════════════════════════════════════════

class AegisFlowApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("AegisFlow  ·  Transform, Validate, Deliver")
        self.geometry("1040x700")
        self.minsize(860, 560)

        self._cur_id = None
        self._panels: dict[str, ToolPanel] = {}
        self._nav_btns: dict[str, ctk.CTkButton] = {}

        self._build_layout()
        self._build_sidebar()
        self._build_content()
        self._build_log()

        # Redirector
        self.redirector = StdoutRedirector(self._log_box)

        # Show first tool
        self._switch("mosaic")

    # ── Layout ─────────────────────────────────────────────────────────────

    def _build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=230, corner_radius=0,
                          fg_color=("#1e293b", "#0f172a"))
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        self._sidebar = sb

        # Brand header
        brand = ctk.CTkFrame(sb, fg_color="transparent")
        brand.pack(fill="x", padx=18, pady=(28, 6))
        ctk.CTkLabel(brand, text="⚡ AegisFlow",
                     font=(FONT_FAMILY, 21, "bold"),
                     text_color="#f8fafc").pack(anchor="w")
        ctk.CTkLabel(brand, text="Transform · Validate · Deliver",
                     font=(FONT_FAMILY, 10),
                     text_color="#475569").pack(anchor="w", pady=(2, 0))

        ctk.CTkFrame(sb, height=1, fg_color="#334155").pack(fill="x", padx=18, pady=14)
        ctk.CTkLabel(sb, text="T O O L S", font=(FONT_FAMILY, 9, "bold"),
                     text_color="#475569").pack(anchor="w", padx=22, pady=(0, 6))

        for t in TOOL_DEFS:
            btn = ctk.CTkButton(
                sb,
                text=f"  {t['icon']}  {t['name']}",
                anchor="w",
                height=44,
                font=(FONT_FAMILY, 13),
                fg_color="transparent",
                text_color="#94a3b8",
                hover_color="#1e3a5f",
                corner_radius=8,
                command=lambda tid=t["id"]: self._switch(tid),
            )
            btn.pack(fill="x", padx=10, pady=2)
            self._nav_btns[t["id"]] = btn

        # Bottom bar
        ctk.CTkFrame(sb, height=1, fg_color="#334155").pack(
            fill="x", padx=18, side="bottom", pady=(0, 6))
        self._mode_btn = ctk.CTkButton(
            sb, text="🌙  Dark Mode", anchor="w",
            height=36, font=(FONT_FAMILY, 11),
            fg_color="transparent", text_color="#64748b",
            hover_color="#1e293b", corner_radius=8,
            command=self._toggle_mode)
        self._mode_btn.pack(side="bottom", fill="x", padx=10, pady=4)

        ver = ctk.CTkLabel(sb, text="v 2026-03", font=(FONT_FAMILY, 10), text_color="#334155")
        ver.pack(side="bottom", pady=(0, 4))

    def _build_content(self):
        right = ctk.CTkFrame(self, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)
        self._right = right

        scr = ctk.CTkScrollableFrame(right, fg_color=("#f1f5f9", "#071428"),
                                      corner_radius=0)
        scr.grid(row=0, column=0, sticky="nsew")
        scr.grid_columnconfigure(0, weight=1)
        self._scroll = scr

        panel_map = {
            "mosaic":   MosaicPanel,
            "template": TemplatePanel,
            "status":   StatusPanel,
            "extract":  ExtractPanel,
            "xml":      XMLPanel,
        }
        for tid, cls in panel_map.items():
            p = cls(scr, self)
            p.grid(row=0, column=0, sticky="nsew", padx=30, pady=26)
            self._panels[tid] = p
            p.grid_remove()

    def _build_log(self):
        LOG_H = 210
        log_frame = ctk.CTkFrame(self._right, fg_color=("white", "#0a1628"),
                                  corner_radius=0, height=LOG_H)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.grid_propagate(False)
        self._right.grid_rowconfigure(1, minsize=LOG_H)

        bar = ctk.CTkFrame(log_frame, fg_color="transparent")
        bar.pack(fill="x", padx=16, pady=(8, 2))
        ctk.CTkLabel(bar, text="📋  Log Output",
                     font=(FONT_FAMILY, 12, "bold"),
                     text_color=("#374151", "#94a3b8")).pack(side="left")
        ctk.CTkButton(bar, text="Clear", width=58, height=24, font=(FONT_FAMILY, 11),
                      fg_color="transparent",
                      text_color=("#9ca3af", "#64748b"),
                      hover_color=("#f3f4f6", "#1e293b"),
                      command=self.log_clear).pack(side="right")

        self._log_box = ctk.CTkTextbox(
            log_frame,
            font=(FONT_MONO, 11),
            fg_color="transparent",
            text_color=("#1e293b", "#d1d5db"),
            wrap="word",
            state="disabled",
        )
        self._log_box.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # Colour tags
        self._log_box.tag_config("ok",   foreground="#16a34a")   # green
        self._log_box.tag_config("err",  foreground="#dc2626")   # red
        self._log_box.tag_config("warn", foreground="#d97706")   # amber

    # ── Actions ────────────────────────────────────────────────────────────

    def _switch(self, tool_id: str):
        if self._cur_id:
            self._panels[self._cur_id].grid_remove()
            self._nav_btns[self._cur_id].configure(
                fg_color="transparent", text_color="#94a3b8", font=(FONT_FAMILY, 13))

        self._cur_id = tool_id
        self._panels[tool_id].grid()
        self._nav_btns[tool_id].configure(
            fg_color="#1e3a5f", text_color="#60a5fa", font=(FONT_FAMILY, 13, "bold"))

    def log_clear(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def log_timestamp(self):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_box.configure(state="normal")
        self._log_box.insert("end", f"\n── {ts} ─────────────────────────────\n", "ts")
        self._log_box.tag_config("ts", foreground="#64748b")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _toggle_mode(self):
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("Light")
            self._mode_btn.configure(text="🌙  Dark Mode")
        else:
            ctk.set_appearance_mode("Dark")
            self._mode_btn.configure(text="☀️  Light Mode")


# ═══════════════════════════════════════════════════════════════════════════
#  Entry Point
# ═══════════════════════════════════════════════════════════════════════════

def main():
    if not TOOLS_LOADED:
        # Fallback: show error in a plain Tk window  
        import tkinter as tk
        import tkinter.messagebox as mb
        root = tk.Tk()
        root.withdraw()
        mb.showerror("Import Error",
                     f"Failed to import required modules:\n\n{IMPORT_ERROR}\n\n"
                     "Please ensure pandas and openpyxl are installed.")
        root.destroy()
        return
    app = AegisFlowApp()
    app.mainloop()


if __name__ == "__main__":
    main()
