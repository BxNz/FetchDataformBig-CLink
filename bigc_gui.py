"""
Big C Scraper - GUI
Keep this file in the SAME folder as bigc_scraper.py

Run:
    uv run --python 3.13 --with playwright --with openpyxl --with requests --with pillow python bigc_gui.py
"""
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import bigc_scraper as core


class QueueWriter:
    """Sends print() output from the worker thread to the log box."""

    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text.strip():
            self.q.put(text.rstrip())

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Big C Scraper -> Excel")
        self.geometry("760x640")
        self.minsize(640, 560)
        self.log_q = queue.Queue()
        self.last_file = None

        pad = {"padx": 12, "pady": 4}

        ttk.Label(self, text="วางลิงก์สินค้า / Paste links (บรรทัดละ 1 ลิงก์)").pack(anchor="w", **pad)
        self.links = scrolledtext.ScrolledText(self, height=8, wrap="none", font=("Consolas", 10))
        self.links.pack(fill="x", padx=12)
        self.links.insert("1.0", core.DEFAULT_URL)

        # file name + folder
        row = ttk.Frame(self)
        row.pack(fill="x", **pad)
        ttk.Label(row, text="ชื่อไฟล์ / File name:").grid(row=0, column=0, sticky="w")
        self.filename = tk.StringVar(value="bigc_products")
        ttk.Entry(row, textvariable=self.filename, width=34).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(row, text=".xlsx").grid(row=0, column=2, sticky="w")

        ttk.Label(row, text="บันทึกที่ / Save to:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.folder = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))
        ttk.Entry(row, textvariable=self.folder, width=60).grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=(6, 0))
        ttk.Button(row, text="เลือกโฟลเดอร์...", command=self.pick_folder).grid(row=1, column=3, padx=6, pady=(6, 0))

        # options
        opt = ttk.Frame(self)
        opt.pack(fill="x", **pad)
        self.embed = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="ฝังรูปในช่อง Excel (ช้ากว่า)", variable=self.embed).pack(side="left")
        ttk.Label(opt, text="   จำนวนหน้าสูงสุดต่อลิงก์:").pack(side="left")
        self.pages = tk.IntVar(value=1)
        ttk.Spinbox(opt, from_=1, to=50, width=4, textvariable=self.pages).pack(side="left", padx=4)

        # buttons
        btns = ttk.Frame(self)
        btns.pack(fill="x", **pad)
        self.start_btn = ttk.Button(btns, text="▶ เริ่มดึงข้อมูล", command=self.start)
        self.start_btn.pack(side="left")
        self.open_btn = ttk.Button(btns, text="เปิดไฟล์", command=self.open_file, state="disabled")
        self.open_btn.pack(side="left", padx=8)
        self.open_dir_btn = ttk.Button(btns, text="เปิดโฟลเดอร์", command=self.open_folder)
        self.open_dir_btn.pack(side="left")

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=12, pady=(4, 0))

        ttk.Label(self, text="สถานะ / Log").pack(anchor="w", **pad)
        self.log = scrolledtext.ScrolledText(self, height=12, state="disabled", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.after(200, self.poll_log)

    # ---------- helpers ----------
    def pick_folder(self):
        d = filedialog.askdirectory(initialdir=self.folder.get())
        if d:
            self.folder.set(d)

    def write_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def poll_log(self):
        try:
            while True:
                self.write_log(self.log_q.get_nowait())
        except queue.Empty:
            pass
        self.after(200, self.poll_log)

    def open_file(self):
        if self.last_file and os.path.exists(self.last_file):
            os.startfile(self.last_file)

    def open_folder(self):
        d = self.folder.get()
        if os.path.isdir(d):
            os.startfile(d)

    # ---------- run ----------
    def start(self):
        urls = [l.strip() for l in self.links.get("1.0", "end").splitlines() if l.strip() and not l.strip().startswith("#")]
        if not urls:
            messagebox.showwarning("ไม่มีลิงก์", "กรุณาวางลิงก์อย่างน้อย 1 ลิงก์")
            return
        name = self.filename.get().strip()
        for ch in '\\/:*?"<>|':
            name = name.replace(ch, "_")
        if not name:
            messagebox.showwarning("ไม่มีชื่อไฟล์", "กรุณาตั้งชื่อไฟล์")
            return
        if name.lower().endswith(".xlsx"):
            name = name[:-5]
        folder = self.folder.get().strip()
        if not os.path.isdir(folder):
            messagebox.showwarning("โฟลเดอร์ไม่ถูกต้อง", "กรุณาเลือกโฟลเดอร์ที่มีอยู่จริง")
            return
        path = os.path.join(folder, name + ".xlsx")
        if os.path.exists(path) and not messagebox.askyesno("ไฟล์มีอยู่แล้ว", f"{name}.xlsx มีอยู่แล้ว ต้องการเขียนทับหรือไม่?\n(ปิดไฟล์ใน Excel ก่อน)"):
            return

        try:
            pages = max(1, int(self.pages.get()))
        except Exception:
            pages = 1

        self.start_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self.progress.start(12)
        self.write_log(f"--- เริ่ม: {len(urls)} ลิงก์ ---")
        threading.Thread(target=self.worker, args=(urls, pages, path, self.embed.get()), daemon=True).start()

    def worker(self, urls, pages, path, embed):
        old = sys.stdout
        sys.stdout = QueueWriter(self.log_q)
        try:
            data = core.scrape(urls, pages)
            if not data:
                print("ไม่พบสินค้า (เว็บอาจบล็อกหรือหน้าเว็บเปลี่ยน)")
                self.after(0, self.done, None)
            else:
                core.save_excel(data, path, embed)
                self.after(0, self.done, path)
        except PermissionError:
            print("บันทึกไม่ได้: ไฟล์ถูกเปิดอยู่ใน Excel กรุณาปิดแล้วลองใหม่")
            self.after(0, self.done, None)
        except Exception as e:
            print(f"ผิดพลาด: {e}")
            self.after(0, self.done, None)
        finally:
            sys.stdout = old

    def done(self, path):
        self.progress.stop()
        self.start_btn.configure(state="normal")
        if path:
            self.last_file = path
            self.open_btn.configure(state="normal")
            self.write_log(f"เสร็จแล้ว ✔  {path}")
            messagebox.showinfo("เสร็จแล้ว", f"บันทึกไฟล์แล้ว:\n{path}")


if __name__ == "__main__":
    App().mainloop()
