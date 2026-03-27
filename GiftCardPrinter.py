"""
Gift Card Printer - GUI App
Auto-detects Poppler, SumatraPDF, and Printers.

Requirements:
    pip install pypdf pdf2image pillow fpdf2 pywin32
"""

import os
import sys
import time
import json
import tempfile
import subprocess
import threading
import urllib.request
import winreg
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject
from pdf2image import convert_from_path
from fpdf import FPDF
import win32print

# ─── CONFIG ───────────────────────────────────────────────────────────────────
VERSION    = "1.1.0"
UPDATE_URL = "https://raw.githubusercontent.com/suresh2216g/giftcardprinter/refs/heads/main/GiftCardPrinter.py"
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), "GiftCardPrinter_settings.json")
DEFAULT_CROPBOX = [400, 85, 790, 270]
# ──────────────────────────────────────────────────────────────────────────────


def find_poppler():
    """Search common locations for Poppler."""
    common = [
        r"C:\Program Files\poppler\Library\bin",
        r"C:\Program Files (x86)\poppler\Library\bin",
        r"C:\poppler\Library\bin",
    ]
    # Search Program Files for any poppler folder
    for base in [r"C:\Program Files", r"C:\Program Files (x86)", r"C:\\"]:
        try:
            for folder in os.listdir(base):
                if "poppler" in folder.lower():
                    candidate = os.path.join(base, folder, "Library", "bin")
                    if os.path.exists(candidate):
                        return candidate
                    candidate = os.path.join(base, folder, "bin")
                    if os.path.exists(candidate):
                        return candidate
        except:
            pass
    for p in common:
        if os.path.exists(p):
            return p
    return ""


def find_sumatra():
    """Search common locations for SumatraPDF."""
    common = [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    for p in common:
        if os.path.exists(p):
            return p
    return ""


def get_printers():
    """Get list of installed printers."""
    try:
        printers = win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
        return [p[2] for p in printers]
    except:
        return []


def load_settings():
    try:
        with open(SETTINGS_FILE, "r") as f:
            return json.load(f)
    except:
        return {}


def save_settings(data):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"🎁 Gift Card Printer v{VERSION}")
        self.resizable(True, True)
        self.minsize(660, 560)
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w, h = 680, min(750, sh - 80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.configure(bg="#1e1e2e")

        # Load saved settings
        s = load_settings()
        self.folder   = tk.StringVar(value=s.get("folder", ""))
        self.poppler  = tk.StringVar(value=s.get("poppler", find_poppler()))
        self.sumatra  = tk.StringVar(value=s.get("sumatra", find_sumatra()))
        self.printer  = tk.StringVar(value=s.get("printer", ""))

        self._build_ui()

        # Auto-select default printer if not saved
        if not self.printer.get() and self.printer_combo["values"]:
            try:
                default = win32print.GetDefaultPrinter()
                self.printer.set(default)
            except:
                self.printer.set(self.printer_combo["values"][0])

    def _build_ui(self):
        # ── Title ──
        tk.Label(self, text="🎁 Gift Card Printer", font=("Segoe UI", 18, "bold"),
                 bg="#1e1e2e", fg="#cdd6f4").pack(pady=(15, 2))
        tk.Label(self, text=f"Crop · Merge · Print to Rollo   v{VERSION}",
                 font=("Segoe UI", 10), bg="#1e1e2e", fg="#a6adc8").pack()

        # ── Settings frame ──
        sf = tk.LabelFrame(self, text=" Settings ", font=("Segoe UI", 9),
                           bg="#1e1e2e", fg="#a6adc8", bd=1, relief="groove")
        sf.pack(fill="x", padx=20, pady=(15, 5))

        def row(parent, label, var, browse_cmd=None, is_combo=False, values=[]):
            f = tk.Frame(parent, bg="#1e1e2e")
            f.pack(fill="x", padx=10, pady=3)
            tk.Label(f, text=label, font=("Segoe UI", 9), bg="#1e1e2e",
                     fg="#cdd6f4", width=14, anchor="w").pack(side="left")
            if is_combo:
                cb = ttk.Combobox(f, textvariable=var, values=values,
                                  font=("Segoe UI", 9), width=42)
                cb.pack(side="left", padx=(5, 0))
                return cb
            else:
                tk.Entry(f, textvariable=var, font=("Segoe UI", 9),
                         bg="#313244", fg="#cdd6f4", insertbackground="white",
                         relief="flat", width=44).pack(side="left", padx=(5, 5))
                if browse_cmd:
                    tk.Button(f, text="Browse", command=browse_cmd,
                              bg="#585b70", fg="white", font=("Segoe UI", 8),
                              relief="flat", padx=6, cursor="hand2").pack(side="left")

        row(sf, "PDF Folder:",   self.folder,  self._browse_folder)
        row(sf, "Poppler Path:", self.poppler, self._browse_poppler)
        row(sf, "SumatraPDF:",   self.sumatra, self._browse_sumatra)
        self.printer_combo = row(sf, "Printer:", self.printer,
                                  is_combo=True, values=get_printers())

        # Save settings button
        tk.Button(sf, text="💾 Save Settings", command=self._save,
                  bg="#313244", fg="#a6adc8", font=("Segoe UI", 8),
                  relief="flat", padx=8, pady=3, cursor="hand2").pack(anchor="e", padx=10, pady=(0, 8))

        # ── Log box ──
        log_frame = tk.Frame(self, bg="#181825")
        log_frame.pack(fill="both", expand=True, padx=20, pady=(5, 5))
        self.log = tk.Text(log_frame, bg="#181825", fg="#cdd6f4",
                           font=("Consolas", 9), relief="flat",
                           state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=10)
        self.log.tag_config("ok",   foreground="#a6e3a1")
        self.log.tag_config("fail", foreground="#f38ba8")
        self.log.tag_config("info", foreground="#89b4fa")
        self.log.tag_config("bold", foreground="#f9e2af", font=("Consolas", 9, "bold"))

        # ── Progress bar ──
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=20, pady=(0, 8))

        # ── Buttons ──
        btn_frame = tk.Frame(self, bg="#1e1e2e")
        btn_frame.pack(pady=(0, 15))
        self._btn("Crop & Merge", "#89b4fa", "#1e66f5", self._run_merge,  btn_frame)
        self._btn("Crop & Print", "#a6e3a1", "#40a02b", self._run_print,  btn_frame)
        self._btn("Clear Log",    "#585b70", "#45475a", self._clear_log,  btn_frame)
        self._btn("⬆ Update",    "#f9e2af", "#df8e1d", self._run_update, btn_frame)

    def _btn(self, text, fg, bg, cmd, parent):
        tk.Button(parent, text=text, command=cmd, bg=bg, fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat",
                  padx=14, pady=8, cursor="hand2",
                  activebackground=fg, activeforeground="white"
                  ).pack(side="left", padx=5)

    def _browse_folder(self):
        d = filedialog.askdirectory(initialdir=self.folder.get() or "C:\\")
        if d: self.folder.set(d)

    def _browse_poppler(self):
        d = filedialog.askdirectory(title="Select Poppler bin folder")
        if d: self.poppler.set(d)

    def _browse_sumatra(self):
        f = filedialog.askopenfilename(title="Select SumatraPDF.exe",
                                        filetypes=[("EXE", "*.exe")])
        if f: self.sumatra.set(f)

    def _save(self):
        save_settings({
            "folder":  self.folder.get(),
            "poppler": self.poppler.get(),
            "sumatra": self.sumatra.get(),
            "printer": self.printer.get(),
        })
        self._log("  ✅ Settings saved!", "ok")

    def _log(self, msg, tag="info"):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy):
        if busy: self.progress.start(10)
        else:    self.progress.stop()

    def _validate(self):
        if not self.folder.get():
            messagebox.showerror("Error", "Please select a PDF folder!"); return False
        if not self.poppler.get() or not os.path.exists(self.poppler.get()):
            messagebox.showerror("Error", "Poppler path not found! Please browse to it."); return False
        if not self.printer.get():
            messagebox.showerror("Error", "Please select a printer!"); return False
        return True

    # ── Update ──────────────────────────────────────────────────────────────

    def _run_update(self):
        def task():
            self._set_busy(True)
            self._log("── Checking for updates ──────────────", "bold")
            try:
                current_file = os.path.abspath(sys.argv[0])
                self._log("  Downloading latest version...", "info")
                tmp = current_file + ".new"
                urllib.request.urlretrieve(UPDATE_URL, tmp)
                with open(tmp, "r") as f:
                    new_content = f.read()
                if f'VERSION    = "{VERSION}"' in new_content:
                    os.unlink(tmp)
                    self._log("  ✅ Already on latest version!", "ok")
                else:
                    os.replace(tmp, current_file)
                    self._log("  ✅ Updated! Restarting...", "ok")
                    time.sleep(2)
                    subprocess.Popen([sys.executable, current_file])
                    self.destroy()
            except Exception as e:
                self._log(f"  [FAIL] {e}", "fail")
            self._set_busy(False)
        threading.Thread(target=task, daemon=True).start()

    # ── Core logic ──────────────────────────────────────────────────────────

    def _crop(self) -> list:
        input_path  = Path(self.folder.get())
        output_path = input_path / "cropped"
        output_path.mkdir(exist_ok=True)

        pdf_files = sorted(input_path.glob("*.pdf"))
        if not pdf_files:
            self._log("No PDFs found in folder!", "fail"); return []

        self._log(f"Found {len(pdf_files)} PDF(s)\n", "bold")
        cropped = []
        for pdf_file in pdf_files:
            try:
                reader = PdfReader(str(pdf_file))
                writer = PdfWriter()
                for page in reader.pages:
                    l, b, r, t = DEFAULT_CROPBOX
                    box = RectangleObject((l, b, r, t))
                    page.cropbox = box
                    page.mediabox = box
                    writer.add_page(page)
                out = output_path / pdf_file.name
                with open(out, "wb") as f:
                    writer.write(f)
                self._log(f"  [CROPPED] {pdf_file.name}", "ok")
                cropped.append(out)
            except Exception as e:
                self._log(f"  [FAIL] {pdf_file.name} — {e}", "fail")
        return cropped

    def _merge(self, cropped: list) -> str:
        self._log(f"\nMerging {len(cropped)} pages into 4x6 PDF...", "bold")
        merged = FPDF(orientation="L", unit="in", format=(4, 6))
        count  = 0
        for pdf in cropped:
            try:
                images = convert_from_path(str(pdf), dpi=300, poppler_path=self.poppler.get())
                img = images[0].convert("RGB")
                tmp_img = tempfile.mktemp(suffix=".jpg")
                img.save(tmp_img, "JPEG", quality=95)
                merged.add_page()
                merged.image(tmp_img, x=0, y=0, w=6, h=4)
                count += 1
                os.unlink(tmp_img)
                self._log(f"  [CONVERTED] {pdf.name}", "ok")
            except Exception as e:
                self._log(f"  [FAIL] {pdf.name} — {e}", "fail")

        out_path = str(Path(self.folder.get()) / "merged_4x6.pdf")
        merged.output(out_path)
        self._log(f"\n  [SAVED] merged_4x6.pdf ({count} pages)", "bold")
        return out_path

    def _print(self, merged_path: str):
        self._log(f"\nSending to {self.printer.get()}...", "bold")
        subprocess.run([
            self.sumatra.get(),
            "-print-to", self.printer.get(),
            "-print-settings", "noscale",
            merged_path
        ])
        self._log("  [PRINTED] ✓", "ok")

    # ── Button handlers ──────────────────────────────────────────────────────

    def _run_merge(self):
        if not self._validate(): return
        def task():
            self._set_busy(True)
            self._log("── Crop & Merge ─────────────────", "bold")
            cropped = self._crop()
            if cropped:
                self._merge(cropped)
                self._log("\n✅ Done! Check merged_4x6.pdf", "ok")
            self._set_busy(False)
        threading.Thread(target=task, daemon=True).start()

    def _run_print(self):
        if not self._validate(): return
        def task():
            self._set_busy(True)
            self._log("── Crop & Print ─────────────────", "bold")
            cropped = self._crop()
            if cropped:
                merged = self._merge(cropped)
                self._print(merged)
                self._log("\n✅ All done!", "ok")
            self._set_busy(False)
        threading.Thread(target=task, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
