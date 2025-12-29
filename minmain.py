import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlwings as xw
import os
import threading
import re
import logging
from datetime import datetime

# =============================================================================
# KONFIGURASI PENCATATAN LOG (SISTEM JURNAL)
# =============================================================================
# Format log menggunakan simbol untuk efisiensi pembacaan visual:
# [I] INFO      : Operasi standar sistem.
# [+] COCOK     : Data sesuai kriteria ditemukan.
# [-] ABAIKAN   : Data/Lembar kerja diabaikan.
# [!] GALAT     : Kesalahan runtime atau pengecualian.
# [D] DUPLIKASI : Redundansi data terdeteksi.
# [>] PROSES    : Memulai sub-rutin baru.

logging.basicConfig(
    filename='debug_log.txt',
    filemode='w',
    level=logging.DEBUG,
    format='%(asctime)s %(message)s',
    datefmt='%H:%M:%S'
)

class BRIProSystem:
    
    # =========================================================================
    # [ KONSTRUKTOR & INISIALISASI ANTARMUKA ]
    # =========================================================================

    def __init__(self, root):
        """
        Konstruktor Kelas Utama.
        
        Fungsi Teknis:
        1. Menginisialisasi objek root Tkinter.
        2. Menetapkan variabel global untuk penampungan data (List) dan cache (Dictionary).
        3. Memanggil metode konstruksi antarmuka grafis (GUI).
        """
        self.root = root
        self.root.title("Sistem Rekapitulasi Operasional (Pencarian Kolom Bertingkat)")
        self.root.geometry("1350x850") 
        self.root.configure(bg="#F4F5F7")

        logging.info("--------------------------------------------------")
        logging.info("[I] INISIASI SISTEM - SIAP MENERIMA PERINTAH")
        logging.info("--------------------------------------------------")

        # Definisi Palet Warna Antarmuka
        self.c_primary = "#00529C"
        self.c_white   = "#FFFFFF"
        self.c_accent  = "#F37021"
        self.c_success = "#27AE60"
        self.c_danger  = "#C0392B"
        self.c_text    = "#2C3E50"

        # Struktur Data Internal
        self.master_data = []      # Penampung utama hasil ekstraksi
        self.failed_files = []     # Log berkas yang gagal diproses
        self.seen_cache = {}       # Tabel Hash untuk deteksi duplikasi {(ID, Tipe): Berkas}
        self.is_processing = False # Bendera status proses

        self.setup_styles()
        self.create_ui()

    def setup_styles(self):
        """
        Konfigurasi Gaya Widget (Style Configuration).
        
        Menerapkan tema 'clam' dan mendefinisikan atribut visual untuk
        Tabel (Treeview) dan Bilah Progres (Progressbar).
        """
        style = ttk.Style()
        style.theme_use("clam")
        
        # Gaya Header Tabel
        style.configure("Treeview.Heading", background=self.c_primary, foreground="white",
                        font=("Segoe UI", 10, "bold"), relief="flat")
        # Gaya Badan Tabel
        style.configure("Treeview", background="white", fieldbackground="white",
                        foreground=self.c_text, rowheight=28, font=("Segoe UI", 10))
        # Gaya Tabel Galat
        style.configure("Error.Treeview.Heading", background=self.c_danger, foreground="white",
                        font=("Segoe UI", 10, "bold"), relief="flat")
        # Gaya Bilah Progres
        style.configure("Horizontal.TProgressbar", background=self.c_accent, troughcolor="#E0E0E0")

    def create_ui(self):
        """
        Pembangunan Antarmuka Pengguna (GUI Construction).
        
        Menyusun komponen visual: Header, Panel Kontrol Input, Indikator Status,
        dan Tabel Data Utama.
        """
        # --- Bagian Header ---
        header = tk.Frame(self.root, bg=self.c_primary, height=90)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        
        tk.Label(header, text="Otomatisasi Terpadu BRI", bg=self.c_primary, 
                 fg="white", font=("Segoe UI", 20, "bold")).pack(side="left", padx=25, pady=5)
        
        info_frame = tk.Frame(header, bg=self.c_primary)
        info_frame.pack(side="right", padx=25, pady=15)
        tk.Label(info_frame, text="Versi: Logika Prioritas Kolom (UKP > UKO > Kode UKO)", 
                 bg=self.c_primary, fg="#BDC3C7", font=("Segoe UI", 10)).pack(anchor="e")

        # --- Bagian Kontrol Input ---
        controls = tk.Frame(self.root, bg="#F4F5F7")
        controls.pack(fill="x", padx=20, pady=15)
        
        input_frame = tk.Frame(controls, bg="white", highlightbackground="#d9d9d9", highlightthickness=1)
        input_frame.pack(side="left", fill="y", padx=(0, 20))
        
        # Parameter 1: Kata Kunci Baris
        tk.Label(input_frame, text="Kata Kunci Baris (Regex):", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_row = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_row.insert(0, r"bandar.*lampung") 
        self.entry_row.grid(row=0, column=1, padx=10, pady=5)

        # Parameter 2: Kata Kunci Kolom
        tk.Label(input_frame, text="Kata Kunci Kolom:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_col = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_col.insert(0, "Grand Total")
        self.entry_col.grid(row=1, column=1, padx=10, pady=5)

        # Parameter 3: Kode Filter
        tk.Label(input_frame, text="Kode Filter (Regex):", bg="white", font=("Segoe UI", 9, "bold"), fg=self.c_accent).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_code = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_code.insert(0, "8204") 
        self.entry_code.grid(row=2, column=1, padx=10, pady=5)

        # Tombol Operasional
        self.btn_run = tk.Button(controls, text="▶ MULAI PROSES", 
                                 bg=self.c_accent, fg="white",
                                 font=("Segoe UI", 10, "bold"), relief="flat", 
                                 padx=20, pady=25, cursor="hand2", 
                                 command=self.start_thread_process)
        self.btn_run.pack(side="left", fill="y", padx=(0, 10))

        self.btn_reset = tk.Button(controls, text="⟳ ATUR ULANG", 
                                   bg=self.c_danger, fg="white",
                                   font=("Segoe UI", 10, "bold"), relief="flat", 
                                   padx=20, pady=25, cursor="hand2", 
                                   command=self.reset_app)
        self.btn_reset.pack(side="left", fill="y", padx=(0, 10))

        self.btn_export = tk.Button(controls, text="💾 EKSPOR EXCEL", 
                                    bg=self.c_success, fg="white",
                                    font=("Segoe UI", 10, "bold"), relief="flat", 
                                    padx=20, pady=25, cursor="hand2", 
                                    state="disabled", 
                                    command=self.export_to_excel)
        self.btn_export.pack(side="left", fill="y")

        # --- Bagian Status ---
        progress_frame = tk.Frame(self.root, bg="#F4F5F7")
        progress_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.lbl_status = tk.Label(progress_frame, text="Status: Siap", bg="#F4F5F7", fg="#7F8C8D", font=("Segoe UI", 10))
        self.lbl_status.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate", style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=5)

        # --- Bagian Tabel Data ---
        table_frame = tk.Frame(self.root, bg="white")
        table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        sy = ttk.Scrollbar(table_frame, orient="vertical")
        sx = ttk.Scrollbar(table_frame, orient="horizontal")
        
        # DEFINISI KOLOM
        self.cols = ["Nomor Kasus", "Tipe Kasus", "Deskripsi", "Tanggal", "Unit Kerja Pelaksana", "Kanca", "Sumber Berkas"]
        self.tree = ttk.Treeview(table_frame, columns=self.cols, show="headings", 
                                 yscrollcommand=sy.set, xscrollcommand=sx.set)
        
        sy.config(command=self.tree.yview)
        sx.config(command=self.tree.xview)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        for c in self.cols:
            self.tree.heading(c, text=c, anchor="w")
            if c == "Sumber Berkas":
                self.tree.column(c, width=200)
            elif c == "Unit Kerja Pelaksana" or c == "Kanca":
                self.tree.column(c, width=180)
            else:
                self.tree.column(c, width=150)

        self.tree.bind("<Control-c>", self.copy_tree)

    # =========================================================================
    # [ UTILITAS & NORMALISASI DATA ]
    # =========================================================================

    def normalize_val(self, val):
        """
        Normalisasi Nilai Sel (Data Normalization Logic).
        Mengonversi tipe data Excel menjadi String bersih untuk konsistensi.
        """
        if val is None:
            return ""
        
        # Penanganan Khusus Format Tanggal
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
        
        # Penanganan Angka (Menghilangkan .0 pada float)
        if isinstance(val, (float, int)):
            if float(val).is_integer():
                return str(int(val)) 
            else:
                return str(val)
        
        return str(val).strip()

    def get_column_index(self, header_row, keyword):
        """
        Pencarian Indeks Kolom Berdasarkan Nama Header (Case-Insensitive).
        Mengembalikan indeks kolom (0-based) jika ditemukan, atau -1 jika tidak.
        """
        keyword_lower = keyword.lower()
        for idx, cell_val in enumerate(header_row):
            if cell_val and keyword_lower in str(cell_val).lower():
                return idx
        return -1

    def clean_error_msg(self, error_obj):
        """
        Sanitasi Pesan Galat.
        Mengubah pesan teknis menjadi informasi yang mudah dipahami.
        """
        raw_msg = str(error_obj)
        if "ShowDetail" in raw_msg:
            return "Kegagalan Drill-Down: Sel target terkunci atau bukan Pivot Table."
        if "tidak ketemu" in raw_msg:
            return raw_msg 
        if "Exception occurred" in raw_msg or "-2147" in raw_msg:
            match = re.search(r"Microsoft Excel', '(.*?)',", raw_msg)
            if match:
                return f"Galat Excel: {match.group(1)}"
            return "Galat Komunikasi Excel (COM Error)."
        return raw_msg

    # =========================================================================
    # [ LOGIKA PROSES: UTAS & PEKERJA ]
    # =========================================================================

    def start_thread_process(self):
        """
        Inisiasi Proses Latar Belakang (Thread Initialization).
        """
        row_kw = self.entry_row.get()
        col_kw = self.entry_col.get()
        code_kw = self.entry_code.get().strip()
        
        if not code_kw:
             messagebox.showwarning("Peringatan", "Parameter Kode Filter wajib diisi.")
             return

        logging.info(f"[>] MEMULAI BATCH. Params: Baris='{row_kw}', Kolom='{col_kw}', Filter='{code_kw}'")
        
        file_paths = filedialog.askopenfilenames(filetypes=[("Berkas Excel", "*.xlsx;*.xls;*.xlsb")])
        if not file_paths: return

        self.is_processing = True
        self.btn_run.config(state="disabled", bg="#95A5A6")
        self.btn_reset.config(state="disabled")
        self.btn_export.config(state="disabled")
        
        self.progress_bar["maximum"] = len(file_paths)
        self.progress_bar["value"] = 0
        self.failed_files = [] 

        t = threading.Thread(target=self.worker_process, args=(file_paths, row_kw, col_kw, code_kw))
        t.daemon = True
        t.start()

    def worker_process(self, file_paths, row_kw, col_kw, filter_code):
        """
        Logika Utama Pekerja (Worker Main Logic).
        """
        app = None
        try:
            self.update_ui_progress(0, "Menginisialisasi Mesin Excel...")
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            
            total_files = len(file_paths)
            
            for index, path in enumerate(file_paths):
                filename = os.path.basename(path)
                self.update_ui_progress(index, f"Memproses ({index+1}/{total_files}): {filename}")
                logging.info(f"[>] MEMULAI BERKAS: {filename}")
                
                try:
                    self.process_single_file(app, path, row_kw, col_kw, filter_code)
                    logging.info(f"[+] SELESAI BERKAS: {filename}")
                except Exception as e:
                    logging.error(f"[!] GALAT FATAL {filename}: {str(e)}")
                    self.failed_files.append({'file': filename, 'msg': self.clean_error_msg(e)})
                
                self.update_ui_progress(index + 1, f"Selesai: {filename}")

        except Exception as e:
            logging.critical(f"[!] KEGAGALAN SISTEM: {str(e)}", exc_info=True)
            self.failed_files.append({'file': "KESALAHAN SISTEM", 'msg': str(e)})
        finally:
            if app:
                try:
                    app.quit()
                except:
                    pass
            self.root.after(0, self.finish_processing)

    def process_single_file(self, app, path, row_regex, col_keyword, filter_code):
        """
        Pemrosesan Berkas Tunggal.
        Meliputi navigasi sheet, pencarian koordinat, drill-down, dan ekstraksi data dinamis.
        """
        wb = app.books.open(path)
        data_found = False
        sheet_errors = []

        try:
            for sheet in wb.sheets:
                # Filter Lembar Kerja yang Tidak Relevan
                if sheet.name.strip().upper() in ["TABEL", "TABLE", "SHEET1"]:
                    continue

                try:
                    # Pengambilan Data Mentah
                    used_range = sheet.used_range
                    data_val = used_range.value 
                    
                    if not data_val: continue 

                    start_row = used_range.row
                    start_col = used_range.column
                    
                    num_rows = len(data_val)
                    num_cols = len(data_val[0]) if num_rows > 0 else 0

                    target_r = -1
                    target_c = -1
                    
                    # 1. Pindai Baris (Prioritas Kolom)
                    for c in range(num_cols):
                        for r in range(num_rows):
                            val = data_val[r][c]
                            str_val = str(val).strip() if val is not None else ""
                            if re.search(row_regex, str_val, re.IGNORECASE):
                                target_r = start_row + r
                                break
                        if target_r != -1: break
                    
                    # 2. Pindai Kolom (Prioritas Baris)
                    for r in range(num_rows):
                        for c in range(num_cols):
                            val = data_val[r][c]
                            str_val = str(val).strip() if val is not None else ""
                            if col_keyword.lower() in str_val.lower():
                                target_c = start_col + c
                                break 
                        if target_c != -1: break 

                    if target_r == -1 or target_c == -1:
                        continue 

                    # Validasi Sel Target
                    target_cell = sheet.cells(target_r, target_c)
                    if target_cell.value is None:
                        logging.warning(f"[-] Sel target kosong pada lembar {sheet.name}")
                        continue 

                    # Eksekusi Drill Down
                    init_sheet_count = len(wb.sheets)
                    target_cell.api.ShowDetail = True 
                    
                    if len(wb.sheets) <= init_sheet_count:
                        logging.warning(f"[-] Drill down tidak menghasilkan lembar baru pada {sheet.name}")
                        continue

                    # Ekstraksi Data dari Lembar Baru
                    new_sheet = wb.sheets.active 
                    raw_extracted_data = new_sheet.used_range.value
                    
                    self.extract_data_manual(raw_extracted_data, os.path.basename(path), filter_code)
                    data_found = True
                    break 

                except Exception as inner_e:
                    sheet_errors.append(f"{sheet.name}: {str(inner_e)}")
                    continue 

            if not data_found:
                if not sheet_errors:
                    raise Exception(f"Koordinat Kata Kunci tidak ditemukan.")
                else:
                    raise Exception("; ".join(sheet_errors))

        finally:
            wb.close()

    # =========================================================================
    # [ LOGIKA INTI: EKSTRAKSI & DUPLIKASI ]
    # =========================================================================

    def extract_data_manual(self, raw_data, filename, filter_code):
        """
        Logika Ekstraksi Data Manual dengan Pencarian Kolom Dinamis & Bertingkat (Fallback).
        Mekanisme: UKP -> UKO -> Kode UKO.
        """
        if not raw_data or len(raw_data) < 2:
            logging.warning("[-] Data hasil ekstraksi kosong atau hanya header.")
            return

        header_row = raw_data[0]
        data_rows = raw_data[1:]
        
        # --- PENCARIAN INDEKS KOLOM DINAMIS (LOGIKA BERTINGKAT) ---
        
        # 1. Cari Unit Kerja Pelaksana (Prioritas Utama)
        idx_uker = self.get_column_index(header_row, "Unit Kerja Pelaksana")
        
        # 2. Jika tidak ketemu, cari Unit Kerja Operasional
        if idx_uker == -1:
            logging.info("[I] Unit Kerja Pelaksana tidak ditemukan. Mencari Unit Kerja Operasional...")
            idx_uker = self.get_column_index(header_row, "Unit Kerja Operasional")
            
        # 3. Jika tidak ketemu, cari Kode UKO
        if idx_uker == -1:
            logging.info("[I] Unit Kerja Operasional tidak ditemukan. Mencari Kode UKO...")
            idx_uker = self.get_column_index(header_row, "Kode UKO")

        # 4. Cari Kolom Kanca (Untuk kolom tambahan)
        idx_kanca = self.get_column_index(header_row, "Kanca")
        # Opsional: Jika Kanca tidak ada, bisa cari "Branch" atau "Cabang" jika diperlukan
        if idx_kanca == -1:
             idx_kanca = self.get_column_index(header_row, "Cabang")

        logging.info(f"[I] Hasil Pemetaan Kolom: Unit Kerja (Indeks={idx_uker}), Kanca (Indeks={idx_kanca})")

        regex_pattern = re.compile(filter_code, re.IGNORECASE)
        count_found = 0

        for row in data_rows:
            # Pastikan baris memiliki cukup kolom
            # Kita cari indeks maksimum yang dibutuhkan untuk menghindari IndexOutOfRange
            max_idx = max(1, 4, 18, idx_uker, idx_kanca) 
            if len(row) <= max_idx:
                row = row + [None] * (max_idx - len(row) + 1)

            # Ambil Kode (Asumsi Kolom B / Indeks 1 adalah Tipe Kasus/Kode)
            raw_code = row[1]
            clean_code = self.normalize_val(raw_code)
            
            # Pencocokan Pola (Regex Match)
            if regex_pattern.search(clean_code):
                
                # Normalisasi Bidang Data Dasar
                val_a = self.normalize_val(row[0])  # No Kasus
                val_b = clean_code                  # Tipe Kasus
                val_d = self.normalize_val(row[3])  # Deskripsi
                val_e = self.normalize_val(row[4])  # Tanggal
                
                # Ekstraksi Dinamis untuk Unit Kerja (Sesuai Prioritas yang ditemukan)
                val_uker = self.normalize_val(row[idx_uker]) if idx_uker != -1 else ""
                
                # Ekstraksi Dinamis untuk Kanca
                val_kanca = self.normalize_val(row[idx_kanca]) if idx_kanca != -1 else ""
                
                # Cek Duplikasi Global
                unique_key = (val_a, val_b)
                
                if unique_key in self.seen_cache:
                    prev_file = self.seen_cache[unique_key]
                    logging.info(f"[D] DUPLIKASI: {val_a}. Sumber Asal: {prev_file}. Mengabaikan.")
                    continue 
                
                # Simpan Data Valid: [... , Unit Kerja, Kanca, Sumber]
                self.seen_cache[unique_key] = filename
                self.master_data.append([val_a, val_b, val_d, val_e, val_uker, val_kanca, filename])
                count_found += 1

        logging.info(f"[I] Ekstraksi {count_found} baris valid dari {filename}")

    # =========================================================================
    # [ MANAJEMEN UI & EKSPOR ]
    # =========================================================================

    def update_ui_progress(self, value, text):
        self.root.after(0, lambda: self._do_update(value, text))

    def _do_update(self, value, text):
        self.progress_bar["value"] = value
        self.lbl_status.config(text=text)

    def finish_processing(self):
        self.is_processing = False
        self.btn_run.config(state="normal", bg=self.c_accent)
        self.btn_reset.config(state="normal")
        self.refresh_table()
        
        if self.master_data:
            self.btn_export.config(state="normal")

        success_cnt = len(self.master_data)
        error_cnt = len(self.failed_files)
        self.lbl_status.config(text=f"Selesai! Data: {success_cnt}. Galat: {error_cnt}")

        if error_cnt > 0:
            self.show_error_window_gui()
        else:
            messagebox.showinfo("Sukses", f"Proses Selesai.\nTotal Data Unik: {success_cnt}")

    def reset_app(self):
        logging.info("[>] PERMINTAAN RESET APLIKASI")
        if self.is_processing:
            messagebox.showwarning("Peringatan", "Harap tunggu proses selesai.")
            return
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.master_data = []
        self.failed_files = []
        self.seen_cache = {}
        
        self.btn_export.config(state="disabled")
        self.lbl_status.config(text="Status: Siap (Reset Selesai)")
        self.progress_bar["value"] = 0
        messagebox.showinfo("Atur Ulang", "Memori data dan cache duplikasi telah dikosongkan.")

    def show_error_window_gui(self):
        top = tk.Toplevel(self.root)
        top.title("Laporan Kegagalan Berkas")
        top.geometry("900x500") 
        top.configure(bg="#F4F5F7")

        tk.Label(top, text="⚠️ Laporan Galat", bg="#F4F5F7", font=("Segoe UI", 14, "bold"), fg=self.c_danger).pack(pady=(15,5))
        
        frame_table = tk.Frame(top)
        frame_table.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        sy = ttk.Scrollbar(frame_table, orient="vertical")
        sx = ttk.Scrollbar(frame_table, orient="horizontal")

        cols = ["Nama Berkas", "Penyebab Galat"]
        tree_err = ttk.Treeview(frame_table, columns=cols, show="headings", yscrollcommand=sy.set, xscrollcommand=sx.set) 
        tree_err.heading("Nama Berkas", text="Nama Berkas")
        tree_err.heading("Penyebab Galat", text="Penyebab Galat")
        tree_err.column("Nama Berkas", width=300)
        tree_err.column("Penyebab Galat", width=500)
        
        sy.config(command=tree_err.yview)
        sx.config(command=tree_err.xview)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        tree_err.pack(fill="both", expand=True)

        tree_err.tag_configure('err_row', foreground=self.c_danger)
        for item in self.failed_files:
            tree_err.insert("", "end", values=(item['file'], item['msg']), tags=('err_row',))

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for row in self.master_data:
            self.tree.insert("", "end", values=row)

    def export_to_excel(self):
        if not self.master_data: return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Berkas Excel", "*.xlsx")], initialfile=f"Rekap_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not file_path: return
        
        try:
            app = xw.App(visible=False)
            wb = app.books.add()
            sheet = wb.sheets[0]
            
            # Tulis Header & Data
            sheet.range("A1").value = [self.cols] + self.master_data
            
            wb.save(file_path)
            wb.close()
            app.quit()
            messagebox.showinfo("Sukses", f"Berkas berhasil disimpan di:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Gagal Simpan", str(e))

    def copy_tree(self, event):
        sel = self.tree.selection()
        if not sel: return
        res = ""
        for i in sel:
            vals = self.tree.item(i, 'values')
            res += "\t".join(vals) + "\n"
        self.root.clipboard_clear()
        self.root.clipboard_append(res)

if __name__ == "__main__":
    root = tk.Tk()
    app = BRIProSystem(root)
    root.mainloop()