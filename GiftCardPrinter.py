"""
Gift Card Printer - GUI App
Drop PDFs in AMZ folder, click Print!

Requirements:
    pip install pypdf pdf2image pillow fpdf2
    Poppler: https://github.com/oschwartz10612/poppler-windows/releases
"""

import os
import sys
import time
import tempfile
import subprocess
import threading
import urllib.request
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject
from pdf2image import convert_from_path
from fpdf import FPDF

# ─── CONFIG ───────────────────────────────────────────────────────────────────
VERSION              = "1.0.0"
UPDATE_URL           = "https://raw.githubusercontent.com/suresh2216g/giftcardprinter/refs/heads/main/GiftCardPrinter.py"
DEFAULT_INPUT_FOLDER = r"C:\Users\ankit\OneDrive\Desktop\AMZ"
DEFAULT_CROPBOX      = [400, 85, 790, 270]
POPPLER_PATH         = r"C:\Program Files\poppler-25.12.0\Library\bin"
SUMATRA_PATH         = r"C:\Users\ankit\AppData\Local\SumatraPDF\SumatraPDF.exe"
ROLLO_PRINTER        = "Rollo X1040 (Copy 1)"
MERGED_OUTPUT        = r"C:\Users\ankit\OneDrive\Desktop\AMZ\merged_4x6.pdf"
# ──────────────────────────────────────────────────────────────────────────────


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"🎁 Gift Card Printer v{VERSION}")
        self.resizable(True, True)
        self.minsize(600, 500)
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w, h = 620, min(700, sh - 80)
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.configure(bg="#1e1e2e")

        self.folder = tk.StringVar(value=DEFAULT_INPUT_FOLDER)
        self._build_ui()

    def _build_ui(self):
        # ── Title ──
        tk.Label(self, text="🎁 Gift Card Printer", font=("Segoe UI", 18, "bold"),
                 bg="#1e1e2e", fg="#cdd6f4").pack(pady=(20, 5))
        tk.Label(self, text=f"Crop · Merge · Print to Rollo   v{VERSION}",
                 font=("Segoe UI", 10), bg="#1e1e2e", fg="#a6adc8").pack()

        # ── Folder picker ──
        frame = tk.Frame(self, bg="#1e1e2e")
        frame.pack(fill="x", padx=20, pady=(20, 5))
        tk.Label(frame, text="PDF Folder:", font=("Segoe UI", 10),
                 bg="#1e1e2e", fg="#cdd6f4").pack(side="left")
        tk.Entry(frame, textvariable=self.folder, font=("Segoe UI", 9),
                 bg="#313244", fg="#cdd6f4", insertbackground="white",
                 relief="flat", width=45).pack(side="left", padx=(8, 5))
        tk.Button(frame, text="Browse", command=self._browse,
                  bg="#585b70", fg="white", font=("Segoe UI", 9),
                  relief="flat", padx=8, cursor="hand2").pack(side="left")

        # ── Log box ──
        log_frame = tk.Frame(self, bg="#181825", relief="flat")
        log_frame.pack(fill="both", expand=True, padx=20, pady=(10, 10))
        self.log = tk.Text(log_frame, bg="#181825", fg="#cdd6f4",
                           font=("Consolas", 9), relief="flat",
                           state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=10)
        self.log.tag_config("ok",   foreground="#a6e3a1")
        self.log.tag_config("fail", foreground="#f38ba8")
        self.log.tag_config("info", foreground="#89b4fa")
        self.log.tag_config("bold", foreground="#f9e2af", font=("Consolas", 9, "bold"))

        # ── Progress bar ──
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=560)
        self.progress.pack(padx=20, pady=(0, 10))

        # ── Buttons ──
        btn_frame = tk.Frame(self, bg="#1e1e2e")
        btn_frame.pack(pady=(0, 20))
        self._btn("Crop & Merge", "#89b4fa", "#1e66f5", self._run_merge, btn_frame)
        self._btn("Crop & Print", "#a6e3a1", "#40a02b", self._run_print, btn_frame)
        self._btn("Clear Log",    "#585b70", "#45475a", self._clear_log, btn_frame)
        self._btn("⬆ Update",    "#f9e2af", "#df8e1d", self._run_update, btn_frame)

    def _btn(self, text, fg, bg, cmd, parent):
        tk.Button(parent, text=text, command=cmd,
                  bg=bg, fg="white", font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=16, pady=8, cursor="hand2",
                  activebackground=fg, activeforeground="white"
                  ).pack(side="left", padx=6)

    def _browse(self):
        folder = filedialog.askdirectory(initialdir=self.folder.get())
        if folder:
            self.folder.set(folder)

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
        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()

    # ── Update ──────────────────────────────────────────────────────────────

    def _run_update(self):
        def task():
            self._set_busy(True)
            self._log("── Checking for updates ─────────────", "bold")
            try:
                current_file = os.path.abspath(sys.argv[0])
                self._log(f"  Downloading latest version...", "info")
                urllib.request.urlretrieve(UPDATE_URL, current_file + ".new")

                # Check version in new file
                with open(current_file + ".new", "r") as f:
                    new_content = f.read()

                if f'VERSION              = "{VERSION}"' in new_content:
                    os.unlink(current_file + ".new")
                    self._log("  ✅ Already on latest version!", "ok")
                else:
                    # Replace current file with new
                    os.replace(current_file + ".new", current_file)
                    self._log("  ✅ Updated! Restarting...", "ok")
                    time.sleep(2)
                    subprocess.Popen([sys.executable, current_file])
                    self.destroy()

            except Exception as e:
                self._log(f"  [FAIL] {e}", "fail")
                self._log("  Make sure UPDATE_URL is set correctly in the script.", "fail")
            self._set_busy(False)
        threading.Thread(target=task, daemon=True).start()

    # ── Core logic ──────────────────────────────────────────────────────────

    def _crop(self) -> list:
        input_path = Path(self.folder.get())
        output_path = input_path / "cropped"
        output_path.mkdir(exist_ok=True)

        pdf_files = sorted(input_path.glob("*.pdf"))
        if not pdf_files:
            self._log("No PDFs found in folder!", "fail")
            return []

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
        count = 0
        for pdf in cropped:
            try:
                images = convert_from_path(str(pdf), dpi=300, poppler_path=POPPLER_PATH)
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

        merged.output(MERGED_OUTPUT)
        self._log(f"\n  [SAVED] merged_4x6.pdf ({count} pages)", "bold")
        return MERGED_OUTPUT

    def _print(self, merged_path: str):
        self._log(f"\nSending to {ROLLO_PRINTER}...", "bold")
        subprocess.run([
            SUMATRA_PATH,
            "-print-to", ROLLO_PRINTER,
            "-print-settings", "noscale",
            merged_path
        ])
        self._log("  [PRINTED] ✓", "ok")

    # ── Button handlers ──────────────────────────────────────────────────────

    def _run_merge(self):
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
