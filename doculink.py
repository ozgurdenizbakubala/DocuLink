#!/usr/bin/env python3
"""
Word ↔ PDF Çift Yönlü Dönüştürücü
Word → PDF  |  PDF → Word
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import threading
import os
import sys

# ---------- pdf2docx kurulum kontrolü ----------
def ensure_pdf2docx():
    try:
        import pdf2docx  # noqa
        return True
    except ImportError:
        return False

# ================================================
#  ANA UYGULAMA
# ================================================
class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DocuLink - Word ↔ PDF")
        self.root.geometry("620x480")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f4f8")

        self._build_header()
        self._build_tabs()

    # ── Başlık ──────────────────────────────────
    def _build_header(self):
        frame = tk.Frame(self.root, bg="#2563eb", pady=14)
        frame.pack(fill="x")
        tk.Label(frame, text="📄 DocuLink: Akıllı Dönüştürücü",
                 font=("Segoe UI", 16, "bold"), fg="white", bg="#2563eb").pack()
        tk.Label(frame, text="LibreOffice & pdf2docx motoru • Format bütünlüğü korunur",
                 font=("Segoe UI", 9), fg="#bfdbfe", bg="#2563eb").pack()

    # ── Sekmeler ────────────────────────────────
    def _build_tabs(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background="#f0f4f8", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"),
                        padding=[20, 8], background="#dbeafe", foreground="#1e40af")
        style.map("TNotebook.Tab",
                  background=[("selected", "#ff007f")],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        # Sekme 1: Word → PDF
        tab1 = tk.Frame(nb, bg="#f0f4f8")
        nb.add(tab1, text="  📄 Word → PDF  ")
        WordToPDFPane(tab1)

        # Sekme 2: PDF → Word
        tab2 = tk.Frame(nb, bg="#f0f4f8")
        nb.add(tab2, text="  📝 PDF → Word  ")
        PDFToWordPane(tab2)


# ================================================
#  ORTAK TABAN PANEL
# ================================================
class BasePane:
    def __init__(self, parent):
        self.parent = parent
        self.files = []
        self.output_dir = tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        body = tk.Frame(self.parent, bg="#f0f4f8", padx=20, pady=15)
        body.pack(fill="both", expand=True)

        # Dosya seçim kutusu
        file_frame = tk.LabelFrame(body, text=f" 1. {self.source_label} Seç ",
                                    font=("Segoe UI", 10, "bold"),
                                    bg="#f0f4f8", fg="#1e40af", padx=10, pady=8)
        file_frame.pack(fill="x", pady=(0, 10))

        btn_row = tk.Frame(file_frame, bg="#f0f4f8")
        btn_row.pack(fill="x")

        tk.Button(btn_row, text="+ Dosya Ekle", command=self.add_files,
                  bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row, text="🗑 Seçileni Kaldır", command=self.remove_selected,
                  bg="#ef4444", fg="white", font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row, text="Tümünü Temizle", command=self.clear_files,
                  bg="#6b7280", fg="white", font=("Segoe UI", 10),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left")

        self.file_count_label = tk.Label(btn_row, text="0 dosya",
                                          font=("Segoe UI", 10), fg="#6b7280", bg="#f0f4f8")
        self.file_count_label.pack(side="right")

        list_frame = tk.Frame(file_frame, bg="#f0f4f8")
        list_frame.pack(fill="x", pady=(8, 0))
        sb = tk.Scrollbar(list_frame, orient="vertical")
        self.file_listbox = tk.Listbox(list_frame, height=5,
                                        font=("Segoe UI", 9),
                                        selectbackground="#2563eb",
                                        yscrollcommand=sb.set,
                                        relief="solid", bd=1,
                                        bg="white", fg="#374151")
        sb.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")

        # Çıktı klasörü
        out_frame = tk.LabelFrame(body, text=" 2. Çıktı Klasörü ",
                                   font=("Segoe UI", 10, "bold"),
                                   bg="#f0f4f8", fg="#1e40af", padx=10, pady=8)
        out_frame.pack(fill="x", pady=(0, 10))

        out_row = tk.Frame(out_frame, bg="#f0f4f8")
        out_row.pack(fill="x")

        self.out_entry = tk.Entry(out_row, textvariable=self.output_dir,
                                   font=("Segoe UI", 10), relief="solid", bd=1)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=4)
        tk.Button(out_row, text="Seç...", command=self.select_output_dir,
                  bg="#e5e7eb", fg="#374151", font=("Segoe UI", 10),
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="right")

        self.same_dir_var = tk.BooleanVar(value=True)
        tk.Checkbutton(out_frame, text="Kaynak dosyalarla aynı klasöre kaydet",
                       variable=self.same_dir_var, command=self.toggle_output_dir,
                       bg="#f0f4f8", font=("Segoe UI", 9), fg="#6b7280").pack(anchor="w", pady=(4, 0))
        self.toggle_output_dir()

        # Dönüştür butonu
        tk.Button(body, text=f"🚀  {self.btn_label}", command=self.start_conversion,
                  bg="#16a34a", fg="white", font=("Segoe UI", 13, "bold"),
                  relief="flat", pady=10, cursor="hand2").pack(fill="x")

        self.progress = ttk.Progressbar(body, mode="determinate")
        self.progress.pack(fill="x", pady=(10, 4))

        self.status_label = tk.Label(body, text="Dosya ekleyip Dönüştür butonuna basın.",
                                      font=("Segoe UI", 9), fg="#6b7280", bg="#f0f4f8")
        self.status_label.pack()

    # ── Ortak metodlar ──────────────────────────
    def toggle_output_dir(self):
        if self.same_dir_var.get():
            self.out_entry.config(state="disabled")
            self.output_dir.set("(kaynak dosyayla aynı klasör)")
        else:
            self.out_entry.config(state="normal")
            self.output_dir.set("")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title=f"{self.source_label} Seç",
            filetypes=self.filetypes
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.file_listbox.insert("end", f"  {os.path.basename(p)}")
        self.file_count_label.config(text=f"{len(self.files)} dosya")

    def remove_selected(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.file_listbox.delete(idx)
        self.files.pop(idx)
        self.file_count_label.config(text=f"{len(self.files)} dosya")

    def clear_files(self):
        self.files.clear()
        self.file_listbox.delete(0, "end")
        self.file_count_label.config(text="0 dosya")

    def select_output_dir(self):
        d = filedialog.askdirectory(title="Çıktı Klasörü Seç")
        if d:
            self.output_dir.set(d)

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("Uyarı", f"Lütfen en az bir {self.source_label} ekleyin.")
            return
        if not self.same_dir_var.get() and not self.output_dir.get():
            messagebox.showwarning("Uyarı", "Lütfen çıktı klasörü seçin.")
            return
        threading.Thread(target=self.convert_all, daemon=True).start()

    def get_out_dir(self, filepath):
        if self.same_dir_var.get():
            return os.path.dirname(filepath)
        return self.output_dir.get()

    def finish(self, success, fail):
        if fail == 0:
            self.status_label.config(
                text=f"✅ Tamamlandı! {success} dosya başarıyla dönüştürüldü.", fg="#16a34a")
            messagebox.showinfo("Başarılı", f"{success} dosya dönüştürüldü!")
        else:
            self.status_label.config(
                text=f"⚠️ {success} başarılı, {fail} başarısız.", fg="#b45309")
            messagebox.showwarning("Kısmi Başarı",
                                   f"{success} dosya dönüştürüldü.\n{fail} dosya dönüştürülemedi.")


# ================================================
#  SEKME 1 — Word → PDF
# ================================================
class WordToPDFPane(BasePane):
    source_label = "Word Dosyası"
    btn_label    = "PDF'e Dönüştür"
    filetypes    = [("Word Dosyaları", "*.docx *.doc"), ("Tüm Dosyalar", "*.*")]

    def convert_all(self):
        total = len(self.files)
        self.progress["maximum"] = total
        self.progress["value"] = 0
        success = fail = 0

        for i, filepath in enumerate(self.files):
            name = os.path.basename(filepath)
            self.status_label.config(text=f"Dönüştürülüyor: {name} ({i+1}/{total})")
            self.parent.update_idletasks()
            out_dir = self.get_out_dir(filepath)
            try:
                result = subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf",
                     "--outdir", out_dir, filepath],
                    capture_output=True, text=True, timeout=60
                )
                success += 1 if result.returncode == 0 else 0
                fail    += 0 if result.returncode == 0 else 1
            except Exception:
                fail += 1
            self.progress["value"] = i + 1
            self.parent.update_idletasks()

        self.finish(success, fail)


# ================================================
#  SEKME 2 — PDF → Word
# ================================================
class PDFToWordPane(BasePane):
    source_label = "PDF Dosyası"
    btn_label    = "Word'e Dönüştür"
    filetypes    = [("PDF Dosyaları", "*.pdf"), ("Tüm Dosyalar", "*.*")]

    def convert_all(self):
        # pdf2docx kurulu değilse otomatik kur
        if not ensure_pdf2docx():
            self.status_label.config(text="⏳ pdf2docx kuruluyor, lütfen bekleyin…", fg="#b45309")
            self.parent.update_idletasks()
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "pdf2docx", "--quiet"],
                    check=True, timeout=120
                )
            except Exception as e:
                messagebox.showerror("Kurulum Hatası",
                                     f"pdf2docx kurulamadı:\n{e}\n\n"
                                     "Terminal'de şunu çalıştırın:\n  pip install pdf2docx")
                self.status_label.config(text="❌ pdf2docx kurulamadı.", fg="#ef4444")
                return

        from pdf2docx import Converter  # noqa

        total = len(self.files)
        self.progress["maximum"] = total
        self.progress["value"] = 0
        success = fail = 0

        for i, filepath in enumerate(self.files):
            name = os.path.basename(filepath)
            self.status_label.config(text=f"Dönüştürülüyor: {name} ({i+1}/{total})")
            self.parent.update_idletasks()

            out_dir  = self.get_out_dir(filepath)
            stem     = os.path.splitext(name)[0]
            out_path = os.path.join(out_dir, stem + ".docx")

            try:
                cv = Converter(filepath)
                cv.convert(out_path, start=0, end=None)
                cv.close()
                success += 1
            except Exception as e:
                fail += 1

            self.progress["value"] = i + 1
            self.parent.update_idletasks()

        self.finish(success, fail)


# ================================================
if __name__ == "__main__":
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()
