# main.py
# ─── Visiting Card Scanner v2 — Redesigned UI ────────────────────────────────

import os
import sys
import ctypes
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# ── HD / DPI fix ──────────────────────────────────────────────────────────────
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from config import APP_TITLE, APP_VERSION, MAX_FILES
from scanner import scan_file, ALL_SUPPORTED
from excel_manager import (
    get_or_create_workbook, append_contact,
    save_workbook, get_contact_count,
    is_duplicate_email, open_excel
)

# ─── Palette ──────────────────────────────────────────────────────────────────
BG          = "#080a0f"
SURFACE     = "#0e1117"
SURFACE2    = "#141820"
BORDER      = "#1e2535"
BORDER2     = "#2a3347"

BLUE        = "#2563ff"
BLUE_BRIGHT = "#4d84ff"
BLUE_DIM    = "#1a3a99"
CYAN        = "#00c8e0"
TEAL        = "#00e5b0"

SUCCESS     = "#00d97e"
SUCCESS_DIM = "#003d24"
WARN        = "#ffab00"
WARN_DIM    = "#3d2900"
DANGER      = "#ff4458"
DANGER_DIM  = "#3d0010"

TEXT        = "#e2e8f4"
TEXT_DIM    = "#8894aa"
TEXT_MUTED  = "#424e63"
WHITE       = "#ffffff"

# ─── Fonts ────────────────────────────────────────────────────────────────────
# Using Consolas (built into Windows, crisp at all sizes) for mono
# and Segoe UI (Windows native, sharp) for sans

F_EYEBROW   = ("Consolas", 9)
F_TITLE     = ("Segoe UI", 26, "bold")
F_SUBTITLE  = ("Segoe UI", 10)
F_LABEL     = ("Consolas", 9)
F_ENTRY     = ("Consolas", 10)
F_BTN_MAIN  = ("Segoe UI", 12, "bold")
F_BTN_SMALL = ("Consolas", 9)
F_STATUS    = ("Consolas", 10)
F_DATA      = ("Consolas", 9)
F_STAT_VAL  = ("Segoe UI", 28, "bold")
F_STAT_LBL  = ("Consolas", 8)


def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(f"{APP_TITLE}  v{APP_VERSION}")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(860, 640)
        self.geometry("980x900")

        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
            self.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            self.tk.call("tk", "scaling", 1.75)

        self.file_rows     = []
        self.is_scanning   = False
        self.session_count = 0

        self._build_ui()
        self._refresh_total()

    # ── Build UI ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Scrollable canvas
        self._canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self._vsb    = tk.Scrollbar(self, orient="vertical",
                                    command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vsb.set)
        self._vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg=BG)
        self._cwin  = self._canvas.create_window(
            (0, 0), window=self._inner, anchor="nw")

        self._inner.bind("<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._cwin, width=e.width))
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(
                int(-1*(e.delta/120)), "units"))

        root = self._inner
        PAD  = 40

        # ── Header bar ────────────────────────────────────────────────────────
        hdr = tk.Frame(root, bg=BG)
        hdr.pack(fill="x", padx=PAD, pady=(40, 0))

        # Left side — titles
        left = tk.Frame(hdr, bg=BG)
        left.pack(side="left", fill="x", expand=True)

        eyebrow = tk.Frame(left, bg=BG)
        eyebrow.pack(anchor="w")
        tk.Label(eyebrow, text="●", font=("Segoe UI", 8),
                 fg=CYAN, bg=BG).pack(side="left")
        tk.Label(eyebrow, text="  OCR TOOL  ·  v2.0",
                 font=F_EYEBROW, fg=TEXT_MUTED, bg=BG).pack(side="left")

        tk.Label(left, text="Visiting Card Scanner",
                 font=F_TITLE, fg=WHITE, bg=BG).pack(anchor="w", pady=(4, 0))

        tk.Label(left,
                 text="Drop files below · Extracts name, phone & email · Saves to contacts.xlsx",
                 font=F_SUBTITLE, fg=TEXT_DIM, bg=BG).pack(anchor="w", pady=(2, 0))

        # Right side — total stat pill
        pill = tk.Frame(hdr, bg=SURFACE2,
                        highlightthickness=1,
                        highlightbackground=BORDER2)
        pill.pack(side="right", anchor="n", padx=(0, 0))
        tk.Label(pill, text="TOTAL CONTACTS",
                 font=F_STAT_LBL, fg=TEXT_MUTED, bg=SURFACE2).pack(padx=24, pady=(14, 0))
        self.stat_total = tk.Label(pill, text="—",
                                   font=F_STAT_VAL, fg=WHITE, bg=SURFACE2)
        self.stat_total.pack(padx=24, pady=(0, 14))

        # ── Divider ───────────────────────────────────────────────────────────
        div = tk.Frame(root, bg=BG, height=28)
        div.pack(fill="x", padx=PAD)
        tk.Frame(div, bg=BORDER, height=1).pack(fill="x", pady=14)

        # ── File import card ──────────────────────────────────────────────────
        card1 = self._card(root, PAD)

        # Card header
        ch1 = tk.Frame(card1, bg=SURFACE)
        ch1.pack(fill="x", padx=24, pady=(20, 12))
        tk.Label(ch1, text="IMPORT FILES",
                 font=F_LABEL, fg=CYAN, bg=SURFACE).pack(side="left")
        self._counter_lbl = tk.Label(ch1, text="0 / 10",
                                     font=F_LABEL, fg=TEXT_MUTED, bg=SURFACE)
        self._counter_lbl.pack(side="right")

        # Rows container
        self._rows_frame = tk.Frame(card1, bg=SURFACE)
        self._rows_frame.pack(fill="x", padx=24)

        # Add file button
        self._add_btn = self._ghost_btn(
            card1, "+ Add Another File",
            command=self._add_row,
            pady_outer=(10, 20)
        )

        # ── Action row ────────────────────────────────────────────────────────
        action = tk.Frame(root, bg=BG)
        action.pack(fill="x", padx=PAD, pady=(20, 0))

        # Main scan button
        self._scan_btn = tk.Button(
            action,
            text="  GET DETAILS  ",
            font=F_BTN_MAIN,
            fg=WHITE,
            bg=BLUE,
            activebackground=BLUE_BRIGHT,
            activeforeground=WHITE,
            relief="flat", bd=0,
            cursor="hand2",
            padx=32, pady=16,
            command=self._start_scan
        )
        self._scan_btn.pack(side="left")

        # Open sheet button
        tk.Button(
            action,
            text="  OPEN CONTACTS SHEET  ",
            font=("Consolas", 10),
            fg=TEAL,
            bg=SURFACE2,
            activebackground=SURFACE,
            activeforeground=WHITE,
            relief="flat", bd=0,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=BORDER2,
            padx=24, pady=16,
            command=self._open_excel
        ).pack(side="left", padx=(12, 0))

        # Session pill
        sess = tk.Frame(action, bg=SURFACE2,
                        highlightthickness=1,
                        highlightbackground=BORDER2)
        sess.pack(side="right")
        tk.Label(sess, text="THIS SESSION",
                 font=F_STAT_LBL, fg=TEXT_MUTED, bg=SURFACE2).pack(
                     padx=20, pady=(10, 0))
        self.stat_session = tk.Label(sess, text="0",
                                     font=("Segoe UI", 20, "bold"),
                                     fg=WHITE, bg=SURFACE2)
        self.stat_session.pack(padx=20, pady=(0, 10))

        # ── Status card ───────────────────────────────────────────────────────
        div2 = tk.Frame(root, bg=BG, height=28)
        div2.pack(fill="x", padx=PAD)
        tk.Frame(div2, bg=BORDER, height=1).pack(fill="x", pady=14)

        card2 = self._card(root, PAD, pady_bottom=40)

        sh = tk.Frame(card2, bg=SURFACE)
        sh.pack(fill="x", padx=24, pady=(20, 12))
        tk.Label(sh, text="SCAN RESULTS",
                 font=F_LABEL, fg=CYAN, bg=SURFACE).pack(side="left")

        self._status_frame = tk.Frame(card2, bg=SURFACE)
        self._status_frame.pack(fill="x", padx=24, pady=(0, 20))

        tk.Label(self._status_frame,
                 text="No results yet — import files and click GET DETAILS.",
                 font=F_DATA, fg=TEXT_MUTED, bg=SURFACE).pack(anchor="w")

        # Add first row
        self._add_row()
        self._update_counter()

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _card(self, parent, pad, pady_bottom=0):
        outer = tk.Frame(parent, bg=SURFACE,
                         highlightthickness=1,
                         highlightbackground=BORDER)
        outer.pack(fill="x", padx=pad, pady=(0, pady_bottom))
        return outer

    def _ghost_btn(self, parent, text, command, pady_outer=(8, 16)):
        btn = tk.Button(
            parent, text=text,
            font=("Consolas", 9),
            fg=TEXT_MUTED, bg=SURFACE,
            activebackground=SURFACE, activeforeground=TEAL,
            relief="flat", bd=0, cursor="hand2",
            padx=0, pady=0,
            command=command
        )
        btn.pack(anchor="w", padx=24,
                 pady=(pady_outer[0], pady_outer[1]))
        return btn

    # ── File rows ─────────────────────────────────────────────────────────────

    def _add_row(self):
        if len(self.file_rows) >= MAX_FILES:
            return

        row = tk.Frame(self._rows_frame, bg=SURFACE)
        row.pack(fill="x", pady=4)

        # Row number badge
        n = len(self.file_rows) + 1
        badge = tk.Label(row, text=f"{n:02d}",
                         font=("Consolas", 9, "bold"),
                         fg=TEXT_MUTED, bg=BORDER,
                         width=3, padx=6, pady=8)
        badge.pack(side="left")

        path_var = tk.StringVar()

        entry = tk.Entry(row, textvariable=path_var,
                         font=F_ENTRY,
                         fg=TEXT, bg=SURFACE2,
                         insertbackground=CYAN,
                         relief="flat", bd=0,
                         highlightthickness=1,
                         highlightbackground=BORDER,
                         highlightcolor=BLUE)
        entry.pack(side="left", fill="x", expand=True,
                   ipady=8, ipadx=10, padx=(6, 0))

        # Browse button
        tk.Button(row, text="BROWSE",
                  font=("Consolas", 8, "bold"),
                  fg=BLUE_BRIGHT, bg=SURFACE2,
                  activebackground=BORDER, activeforeground=WHITE,
                  relief="flat", bd=0, cursor="hand2",
                  highlightthickness=1,
                  highlightbackground=BORDER2,
                  padx=12, pady=8,
                  command=lambda pv=path_var: self._browse(pv)
                  ).pack(side="left", padx=(6, 0))

        # First row clears text only; subsequent rows remove entire row
        is_first = len(self.file_rows) == 0
        btn_cmd  = (lambda pv=path_var: pv.set("")) if is_first else (lambda rf=row: self._remove_row(rf))

        tk.Button(row, text="✕",
                  font=("Segoe UI", 11),
                  fg=TEXT_MUTED, bg=SURFACE,
                  activebackground=SURFACE, activeforeground=DANGER,
                  relief="flat", bd=0, cursor="hand2",
                  padx=8,
                  command=btn_cmd
                  ).pack(side="left", padx=(4, 0))

        self.file_rows.append({"frame": row, "path_var": path_var})
        self._update_counter()

    def _remove_row(self, row_frame):
        self.file_rows = [r for r in self.file_rows
                          if r["frame"] is not row_frame]
        row_frame.destroy()
        self._renumber_rows()
        self._update_counter()

    def _renumber_rows(self):
        for i, r in enumerate(self.file_rows):
            # Update badge number
            for child in r["frame"].winfo_children():
                if isinstance(child, tk.Label) and child.cget("bg") == BORDER:
                    child.config(text=f"{i+1:02d}")
                    break

    def _browse(self, path_var):
        exts = " ".join(f"*{e}" for e in sorted(ALL_SUPPORTED))
        path = filedialog.askopenfilename(
            filetypes=[("Supported files", exts), ("All files", "*.*")]
        )
        if path:
            path_var.set(path)

    def _update_counter(self):
        n = len(self.file_rows)
        self._counter_lbl.config(
            text=f"{n} / {MAX_FILES}",
            fg=WARN if n >= MAX_FILES else TEXT_MUTED
        )
        self._add_btn.config(
            state="disabled" if n >= MAX_FILES else "normal",
            fg=TEXT_MUTED if n < MAX_FILES else TEXT_MUTED
        )

    # ── Scanning ──────────────────────────────────────────────────────────────

    def _start_scan(self):
        if self.is_scanning:
            return

        paths = [r["path_var"].get().strip() for r in self.file_rows
                 if r["path_var"].get().strip()]

        if not paths:
            self._show_status([{
                "file": "", "status": "error",
                "message": "Add at least one file path before scanning."
            }])
            return

        self.is_scanning = True
        self._scan_btn.config(
            text="  SCANNING…  ",
            state="disabled",
            bg=BLUE_DIM
        )
        self._clear_status()

        threading.Thread(
            target=self._scan_worker, args=(paths,), daemon=True
        ).start()

    def _scan_worker(self, paths):
        results = []
        wb, ws  = get_or_create_workbook()
        added   = 0

        for path in paths:
            file_name = Path(path).name
            result    = scan_file(path)
            result["file"] = file_name

            if result["status"] != "error" and result["data"]:
                contacts = result["data"]
                if (contacts["email"] != "N/A"
                        and is_duplicate_email(contacts["email"])):
                    result["status"]  = "partial"
                    result["message"] = "Duplicate — this email already exists in contacts"
                else:
                    append_contact(ws, contacts["name"],
                                   contacts["phone"],
                                   contacts["email"], file_name)
                    added += 1

            results.append(result)

        save_workbook(wb)
        self.after(0, self._scan_done, results, added)

    def _scan_done(self, results, added):
        self.is_scanning = False
        self._scan_btn.config(
            text="  GET DETAILS  ",
            state="normal",
            bg=BLUE
        )
        self._show_status(results)
        self.session_count += added
        self.stat_session.config(text=str(self.session_count))
        self._refresh_total()

    # ── Status ────────────────────────────────────────────────────────────────

    def _clear_status(self):
        for w in self._status_frame.winfo_children():
            w.destroy()


    def _show_status(self, results):
        self._clear_status()

        CFG = {
            "success": (SUCCESS,  SUCCESS_DIM, "✓"),
            "partial":  (WARN,    WARN_DIM,    "⚠"),
            "error":    (DANGER,  DANGER_DIM,  "✕"),
        }

        for r in results:
            status = r.get("status", "error")
            color, bg_dim, icon = CFG.get(status, CFG["error"])

            # Outer wrapper with left accent border effect
            wrap = tk.Frame(self._status_frame, bg=SURFACE)
            wrap.pack(fill="x", pady=4)

            # Coloured left stripe
            tk.Frame(wrap, bg=color, width=4).pack(side="left", fill="y")

            # Content area
            inner = tk.Frame(wrap, bg=bg_dim, padx=16, pady=12)
            inner.pack(side="left", fill="x", expand=True)

            # Top row: icon + filename
            top = tk.Frame(inner, bg=bg_dim)
            top.pack(fill="x")

            tk.Label(top, text=f"{icon}",
                     font=("Segoe UI", 13, "bold"),
                     fg=color, bg=bg_dim).pack(side="left")

            if r.get("file"):
                tk.Label(top, text=f"  {r['file']}",
                         font=("Segoe UI", 10, "bold"),
                         fg=WHITE, bg=bg_dim).pack(side="left")

            # Message
            tk.Label(inner, text=r.get("message", ""),
                     font=F_STATUS, fg=color, bg=bg_dim,
                     wraplength=700, justify="left").pack(
                         anchor="w", pady=(4, 0))

            # Extracted data row
            if r.get("data"):
                d = r["data"]
                data_frame = tk.Frame(inner, bg=bg_dim)
                data_frame.pack(anchor="w", pady=(6, 0))

                for label, val, col in [
                    ("NAME",  d["name"],  TEXT),
                    ("PHONE", d["phone"], TEXT),
                    ("EMAIL", d["email"], CYAN),
                ]:
                    chip = tk.Frame(data_frame, bg=SURFACE2,
                                    highlightthickness=1,
                                    highlightbackground=BORDER2)
                    chip.pack(side="left", padx=(0, 8))
                    tk.Label(chip, text=label,
                             font=("Consolas", 7, "bold"),
                             fg=TEXT_MUTED, bg=SURFACE2).pack(
                                 side="left", padx=(8, 4), pady=5)
                    tk.Label(chip, text=val,
                             font=("Consolas", 9),
                             fg=col, bg=SURFACE2).pack(
                                 side="left", padx=(0, 8), pady=5)

    # ── Excel ─────────────────────────────────────────────────────────────────

    def _open_excel(self):
        ok, err = open_excel()
        if not ok:
            messagebox.showerror("Error", err)

    def _refresh_total(self):
        self.stat_total.config(text=str(get_contact_count()))


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(sys.executable))

    app = App()
    app.mainloop()
