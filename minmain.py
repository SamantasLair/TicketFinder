# ==========================================
# IMPOR PUSTAKA (HANYA YANG DIPERLUKAN)
# ==========================================
# Menggunakan 'from ... import ...' untuk meminimalkan namespace
from tkinter import Tk, Frame, Label, Entry, Button, Scrollbar, Toplevel, filedialog, messagebox, ttk
from tkinter import BOTH, RIGHT, BOTTOM, LEFT, Y, X, END, W, E
from threading import Thread
from datetime import datetime
from re import search, compile, IGNORECASE
from logging import basicConfig, DEBUG, info, warning, error, critical
from os.path import basename
import xlwings as xw  # Pustaka Inti untuk Excel Automation

# ==========================================
# KONFIGURASI LOG
# ==========================================
basicConfig(
    filename='debug_log.txt',
    filemode='w',
    level=DEBUG,
    format='%(asctime)s %(message)s',
    datefmt='%H:%M:%S'
)

class BRIProSystem:
    
    # =========================================================================
    # [ KONSTRUKTOR & UI ]
    # =========================================================================

    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Rekapitulasi Operasional (Versi Ringan)")
        self.root.geometry("1350x850") 
        self.root.configure(bg="#F4F5F7")

        info("--------------------------------------------------")
        info("[I] MULAI SISTEM - MODE RINGAN (TANPA PANDAS)")
        info("--------------------------------------------------")

        # Variabel Warna
        self.c_pri = "#00529C"
        self.c_wht = "#FFFFFF"
        self.c_acc = "#F37021"
        self.c_suc = "#27AE60"
        self.c_err = "#C0392B"
        self.c_txt = "#2C3E50"

        # Kontainer Data (Native List & Dict)
        self.master_data = [] 
        self.failed_files = [] 
        self.seen_cache = {} # Pengganti drop_duplicates Pandas
        self.is_proc = False

        self.setup_styles()
        self.create_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", background=self.c_pri, foreground="white", font=("Segoe UI", 10, "bold"), relief="flat")
        style.configure("Treeview", background="white", fieldbackground="white", foreground=self.c_txt, rowheight=28, font=("Segoe UI", 10))
        style.configure("Error.Treeview.Heading", background=self.c_err, foreground="white", font=("Segoe UI", 10, "bold"), relief="flat")
        style.configure("Horizontal.TProgressbar", background=self.c_acc, troughcolor="#E0E0E0")

    def create_ui(self):
        # --- Header ---
        hdr = Frame(self.root, bg=self.c_pri, height=90)
        hdr.pack(fill=X, side="top")
        hdr.pack_propagate(False)
        
        Label(hdr, text="Otomatisasi Terpadu BRI", bg=self.c_pri, fg="white", font=("Segoe UI", 20, "bold")).pack(side=LEFT, padx=25, pady=5)
        Label(hdr, text="Mode: Native Python (No Pandas)", bg=self.c_pri, fg="#BDC3C7", font=("Segoe UI", 10)).pack(side=RIGHT, padx=25, pady=15, anchor="e")

        # --- Kontrol ---
        ctl = Frame(self.root, bg="#F4F5F7")
        ctl.pack(fill=X, padx=20, pady=15)
        
        inp_frm = Frame(ctl, bg="white")
        inp_frm.pack(side=LEFT, fill=Y, padx=(0, 20))
        
        Label(inp_frm, text="Kata Kunci Baris (Regex):", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky=W)
        self.ent_row = Entry(inp_frm, width=30, font=("Segoe UI", 10))
        self.ent_row.insert(0, r"bandar.*lampung") 
        self.ent_row.grid(row=0, column=1, padx=10, pady=5)

        Label(inp_frm, text="Kata Kunci Kolom:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky=W)
        self.ent_col = Entry(inp_frm, width=30, font=("Segoe UI", 10))
        self.ent_col.insert(0, "Grand Total")
        self.ent_col.grid(row=1, column=1, padx=10, pady=5)

        Label(inp_frm, text="Kode Filter (Regex):", bg="white", font=("Segoe UI", 9, "bold"), fg=self.c_acc).grid(row=2, column=0, padx=10, pady=5, sticky=W)
        self.ent_code = Entry(inp_frm, width=30, font=("Segoe UI", 10))
        self.ent_code.insert(0, "8204") 
        self.ent_code.grid(row=2, column=1, padx=10, pady=5)

        self.btn_run = Button(ctl, text="▶ MULAI", bg=self.c_acc, fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=25, cursor="hand2", command=self.start_proc)
        self.btn_run.pack(side=LEFT, fill=Y, padx=(0, 10))

        self.btn_rst = Button(ctl, text="⟳ RESET", bg=self.c_err, fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=25, cursor="hand2", command=self.reset_app)
        self.btn_rst.pack(side=LEFT, fill=Y, padx=(0, 10))

        self.btn_exp = Button(ctl, text="💾 EXCEL", bg=self.c_suc, fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=25, cursor="hand2", state="disabled", command=self.export_xls)
        self.btn_exp.pack(side=LEFT, fill=Y)

        # --- Status & Tabel ---
        prg_frm = Frame(self.root, bg="#F4F5F7")
        prg_frm.pack(fill=X, padx=20, pady=(0, 10))
        
        self.lbl_stat = Label(prg_frm, text="Status: Siap", bg="#F4F5F7", fg="#7F8C8D", font=("Segoe UI", 10))
        self.lbl_stat.pack(anchor=W)

        self.pbar = ttk.Progressbar(prg_frm, orient="horizontal", length=100, mode="determinate", style="Horizontal.TProgressbar")
        self.pbar.pack(fill=X, pady=5)

        tbl_frm = Frame(self.root, bg="white")
        tbl_frm.pack(fill=BOTH, expand=True, padx=20, pady=(0, 20))

        sy = Scrollbar(tbl_frm, orient="vertical")
        sx = Scrollbar(tbl_frm, orient="horizontal")
        
        self.cols = ["No Kasus", "Tipe Kasus", "Deskripsi", "Tanggal", "Unit Kerja", "Sumber"]
        self.tree = ttk.Treeview(tbl_frm, columns=self.cols, show="headings", yscrollcommand=sy.set, xscrollcommand=sx.set)
        
        sy.config(command=self.tree.yview)
        sx.config(command=self.tree.xview)
        sy.pack(side=RIGHT, fill=Y)
        sx.pack(side=BOTTOM, fill=X)
        self.tree.pack(fill=BOTH, expand=True)

        for c in self.cols:
            self.tree.heading(c, text=c, anchor=W)
            w = 250 if c == "Sumber" else 180
            self.tree.column(c, width=w)

        self.tree.bind("<Control-c>", self.copy_tree)

    # =========================================================================
    # [ FUNGSI UTILITAS (NATIVE PYTHON) ]
    # =========================================================================

    def normalize(self, val):
        """ Membersihkan data (Angka/Teks) menjadi String Murni tanpa dependensi Numpy/Pandas """
        if val is None: return ""
        if isinstance(val, (float, int)):
            # Cek apakah float murni (x.0) atau desimal
            if float(val).is_integer(): return str(int(val)) 
            return str(val)
        return str(val).strip()

    def clean_err(self, err):
        msg = str(err)
        if "ShowDetail" in msg: return "Gagal Drill-Down: Sel terkunci/Bukan Pivot."
        if "-2147" in msg: return "Galat Interaksi Excel (COM Error)."
        return msg

    # =========================================================================
    # [ LOGIKA UTAMA ]
    # =========================================================================

    def start_proc(self):
        row_kw = self.ent_row.get()
        col_kw = self.ent_col.get()
        code_kw = self.ent_code.get().strip()
        
        if not code_kw:
             messagebox.showwarning("Peringatan", "Kode Filter kosong.")
             return

        info(f"[>] START: Row='{row_kw}', Col='{col_kw}', Code='{code_kw}'")
        
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx;*.xls;*.xlsb")])
        if not files: return

        self.is_proc = True
        self.toggle_ui(False)
        self.pbar["maximum"] = len(files)
        self.pbar["value"] = 0
        self.failed_files = [] 

        t = Thread(target=self.worker, args=(files, row_kw, col_kw, code_kw))
        t.daemon = True
        t.start()

    def toggle_ui(self, state):
        st = "normal" if state else "disabled"
        bg = self.c_acc if state else "#95A5A6"
        self.btn_run.config(state=st, bg=bg)
        self.btn_rst.config(state=st)
        self.btn_exp.config(state=st)

    def worker(self, files, r_kw, c_kw, code):
        app = None
        try:
            self.upd_ui(0, "Membuka Mesin Excel...")
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            
            for idx, path in enumerate(files):
                fname = basename(path)
                self.upd_ui(idx, f"Memproses: {fname}")
                
                try:
                    self.process_file(app, path, r_kw, c_kw, code)
                except Exception as e:
                    error(f"[!] Gagal {fname}: {e}")
                    self.failed_files.append((fname, self.clean_err(e)))
                
                self.upd_ui(idx + 1, f"Selesai: {fname}")

        except Exception as e:
            critical(f"[!] CRASH: {e}")
            self.failed_files.append(("SYSTEM", str(e)))
        finally:
            if app: 
                try: app.quit() 
                except: pass
            self.root.after(0, self.finish)

    def process_file(self, app, path, r_regex, c_kw, code):
        wb = app.books.open(path)
        found = False
        
        try:
            for sht in wb.sheets:
                if sht.name.strip().upper() in ["TABEL", "TABLE", "SHEET1"]: continue

                try:
                    # Ambil semua data sebagai List of Lists (Sangat Cepat & Ringan)
                    raw = sht.used_range.value 
                    if not raw: continue 

                    r_idx, c_idx = -1, -1
                    rows, cols = len(raw), len(raw[0]) if len(raw) > 0 else 0

                    # SCANNING (Native Loop)
                    # 1. Cari Baris (Scan Kolom dulu)
                    for c in range(cols):
                        for r in range(rows):
                            val = self.normalize(raw[r][c])
                            if search(r_regex, val, IGNORECASE):
                                r_idx = sht.used_range.row + r
                                info(f"[+] Baris ketemu di ({r_idx}, {sht.used_range.column + c})")
                                break
                        if r_idx != -1: break
                    
                    # 2. Cari Kolom (Scan Baris dulu)
                    for r in range(rows):
                        for c in range(cols):
                            val = self.normalize(raw[r][c])
                            if c_kw.lower() in val.lower():
                                c_idx = sht.used_range.column + c
                                info(f"[+] Kolom ketemu di ({sht.used_range.row + r}, {c_idx})")
                                break 
                        if c_idx != -1: break 

                    if r_idx == -1 or c_idx == -1: continue 

                    # Drill Down
                    cell = sht.cells(r_idx, c_idx)
                    if cell.value is None: continue 
                    
                    cnt_before = len(wb.sheets)
                    cell.api.ShowDetail = True 
                    if len(wb.sheets) <= cnt_before: continue

                    # Ekstraksi
                    new_sht = wb.sheets.active 
                    extracted = new_sht.used_range.value
                    self.parse_data(extracted, basename(path), code)
                    found = True
                    break 

                except Exception as e:
                    warning(f"[-] Skip sheet {sht.name}: {e}")
                    continue 

            if not found: raise Exception("Koordinat tidak ditemukan.")

        finally:
            wb.close()

    def parse_data(self, data, fname, filter_code):
        if not data or len(data) < 2: return
        
        # Skip Header, data mulai index 1
        rows = data[1:]
        reg = compile(filter_code, IGNORECASE)
        match = 0

        for r in rows:
            # Padding jika kolom kurang
            if len(r) < 19: r = r + [None] * (19 - len(r))

            # Kolom B (Index 1) adalah Kode
            code_clean = self.normalize(r[1])
            
            if reg.search(code_clean):
                # Ambil Data
                va = self.normalize(r[0]) # Case No
                vb = code_clean           # Case Type
                vd = self.normalize(r[3]) # Desc
                ve = self.normalize(r[4]) # Date
                vs = self.normalize(r[18])# Unit
                
                # Cek Duplikasi (Pengganti Pandas Drop Duplicate)
                key = (va, vb)
                if key not in self.seen_cache:
                    self.seen_cache[key] = fname
                    self.master_data.append([va, vb, vd, ve, vs, fname])
                    match += 1
                else:
                    info(f"[D] Duplikat: {va} di {fname}")

        info(f"[+] Disimpan {match} data dari {fname}")

    def upd_ui(self, val, txt):
        self.root.after(0, lambda: self._do_upd(val, txt))

    def _do_upd(self, val, txt):
        self.pbar["value"] = val
        self.lbl_stat.config(text=txt)

    def finish(self):
        self.is_proc = False
        self.toggle_ui(True)
        
        # Refresh Tabel GUI
        for i in self.tree.get_children(): self.tree.delete(i)
        for row in self.master_data: self.tree.insert("", END, values=row)

        msg = f"Selesai! Data: {len(self.master_data)}. Galat: {len(self.failed_files)}"
        self.lbl_stat.config(text=msg)

        if self.failed_files: self.show_err()
        else: messagebox.showinfo("Sukses", f"Proses Selesai.\nTotal Data: {len(self.master_data)}")

    def reset_app(self):
        if self.is_proc: return
        for i in self.tree.get_children(): self.tree.delete(i)
        self.master_data = []
        self.failed_files = []
        self.seen_cache = {}
        self.btn_exp.config(state="disabled")
        self.lbl_stat.config(text="Status: Atur Ulang Selesai")
        self.pbar["value"] = 0
        messagebox.showinfo("Reset", "Memori dibersihkan.")

    def export_xls(self):
        if not self.master_data: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=f"Rekap_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not path: return
        
        try:
            # Menggunakan XLWINGS untuk tulis Excel (Tanpa Pandas)
            app = xw.App(visible=False)
            wb = app.books.add()
            sht = wb.sheets[0]
            
            # Tulis Header & Data
            sht.range("A1").value = [self.cols] + self.master_data
            
            wb.save(path)
            wb.close()
            app.quit()
            messagebox.showinfo("Sukses", f"Tersimpan di:\n{path}")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    def show_err(self):
        top = Toplevel(self.root)
        top.title("Laporan Galat")
        top.geometry("800x400")
        
        frm = Frame(top)
        frm.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        sy = Scrollbar(frm, orient="vertical")
        tr = ttk.Treeview(frm, columns=("File", "Err"), show="headings", yscrollcommand=sy.set)
        tr.heading("File", text="Nama Berkas")
        tr.heading("Err", text="Penyebab")
        tr.column("File", width=250)
        tr.column("Err", width=500)
        
        sy.config(command=tr.yview)
        sy.pack(side=RIGHT, fill=Y)
        tr.pack(fill=BOTH, expand=True)
        
        for f, m in self.failed_files: tr.insert("", END, values=(f, m))

    def copy_tree(self, e):
        sel = self.tree.selection()
        if not sel: return
        res = ""
        for i in sel: res += "\t".join([str(x) for x in self.tree.item(i, 'values')]) + "\n"
        self.root.clipboard_clear()
        self.root.clipboard_append(res)

if __name__ == "__main__":
    root = Tk()
    app = BRIProSystem(root)
    root.mainloop()