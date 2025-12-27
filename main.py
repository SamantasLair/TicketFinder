import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import xlwings as xw
import os
import threading
import re
import logging
from datetime import datetime

# ==========================================
# KONFIGURASI LOGGING (FORMAT SIMBOLIS)
# ==========================================
# Legenda Log:
# [I] INFO    : Proses berjalan normal
# [+] MATCH   : Data ditemukan/cocok
# [-] SKIP    : Data/Sheet dilewati
# [!] ERROR   : Terjadi kesalahan teknis
# [D] DUPLICATE : Data ganda ditemukan
# [>] PROCESS : Memulai proses/tahapan baru

logging.basicConfig(
    filename='debug_log.txt',
    filemode='w', # Overwrite setiap kali run
    level=logging.DEBUG,
    format='%(asctime)s %(message)s', # Pesan langsung dengan simbol
    datefmt='%H:%M:%S'
)

class BRIProSystem:
    
    # =========================================================================
    # [ INITIALIZATION & UI SETUP ]
    # =========================================================================

    def __init__(self, root):
        """
        Constructor utama untuk inisialisasi aplikasi GUI.
        
        Teknis:
        - Menginisialisasi objek Tkinter root.
        - Menyiapkan variabel container untuk data (list) dan cache duplikat (dict).
        - Memanggil method setup UI dan Styling.
        """
        self.root = root
        self.root.title("Sistem Rekapitulasi Operasional (Clean Code Edition)")
        self.root.geometry("1350x850") 
        self.root.configure(bg="#F4F5F7")

        # Menulis Header Log
        logging.info("--------------------------------------------------")
        logging.info("[I] SYSTEM STARTUP - LOGGING INITIALIZED")
        logging.info("[I] LEGEND: [+]Found, [-]Skip, [!]Error, [D]Duplicate")
        logging.info("--------------------------------------------------")

        # Color Palette
        self.c_primary = "#00529C"
        self.c_white = "#FFFFFF"
        self.c_accent = "#F37021"
        self.c_success = "#27AE60"
        self.c_danger = "#C0392B"
        self.c_text = "#2C3E50"

        # Data Structures
        self.master_data = []      # List untuk menampung hasil akhir
        self.failed_files = []     # List untuk tracking file error
        self.seen_cache = {}       # Dictionary {(CaseNum, CaseType): SourceFile} untuk cek duplikat
        self.is_processing = False

        self.setup_styles()
        self.create_ui()

    # =========================================================================
    
    def setup_styles(self):
        """
        Mengkonfigurasi style global untuk widget Tkinter (Treeview, Progressbar).
        Menggunakan theme 'clam' untuk tampilan modern.
        """
        style = ttk.Style()
        style.theme_use("clam")
        
        # Style Tabel Utama
        style.configure("Treeview.Heading", background=self.c_primary, foreground="white",
                        font=("Segoe UI", 10, "bold"), relief="flat")
        style.configure("Treeview", background="white", fieldbackground="white",
                        foreground=self.c_text, rowheight=28, font=("Segoe UI", 10))
        
        # Style Tabel Error
        style.configure("Error.Treeview.Heading", background=self.c_danger, foreground="white",
                        font=("Segoe UI", 10, "bold"), relief="flat")
        
        # Style Progress Bar
        style.configure("Horizontal.TProgressbar", background=self.c_accent, troughcolor="#E0E0E0")

    # =========================================================================

    def create_ui(self):
        """
        Membangun layout antarmuka pengguna (GUI) menggunakan Grid dan Pack geometry managers.
        Terdiri dari: Header, Control Panel (Input Regex), Progress Bar, dan Data Table.
        """
        # Header Section
        header = tk.Frame(self.root, bg=self.c_primary, height=90)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        
        tk.Label(header, text="BRI Unified Automation", bg=self.c_primary, 
                 fg="white", font=("Segoe UI", 20, "bold")).pack(side="left", padx=25, pady=5)
        
        info_frame = tk.Frame(header, bg=self.c_primary)
        info_frame.pack(side="right", padx=25, pady=15)
        tk.Label(info_frame, text="Raw Data Reading | Duplicate Tracking", 
                 bg=self.c_primary, fg="#BDC3C7", font=("Segoe UI", 10)).pack(anchor="e")

        # Controls Section
        controls = tk.Frame(self.root, bg="#F4F5F7")
        controls.pack(fill="x", padx=20, pady=15)
        
        input_frame = tk.Frame(controls, bg="white", highlightbackground="#d9d9d9", highlightthickness=1)
        input_frame.pack(side="left", fill="y", padx=(0, 20))
        
        # Input 1: Row Keyword
        tk.Label(input_frame, text="Kata Kunci Baris (Regex):", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_row = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_row.insert(0, r"bandar.*lampung") 
        self.entry_row.grid(row=0, column=1, padx=10, pady=5)

        # Input 2: Column Keyword
        tk.Label(input_frame, text="Kata Kunci Kolom:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_col = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_col.insert(0, "Grand Total")
        self.entry_col.grid(row=1, column=1, padx=10, pady=5)

        # Input 3: Filter Code
        tk.Label(input_frame, text="Kode Filter (Regex):", bg="white", font=("Segoe UI", 9, "bold"), fg=self.c_accent).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_code = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_code.insert(0, "8204") 
        self.entry_code.grid(row=2, column=1, padx=10, pady=5)

        # Action Buttons
        self.btn_run = tk.Button(controls, text="▶ START PROCESS", 
                                 bg=self.c_accent, fg="white",
                                 font=("Segoe UI", 10, "bold"), relief="flat", 
                                 padx=20, pady=25, cursor="hand2", 
                                 command=self.start_thread_process)
        self.btn_run.pack(side="left", fill="y", padx=(0, 10))

        self.btn_reset = tk.Button(controls, text="⟳ RESET", 
                                   bg=self.c_danger, fg="white",
                                   font=("Segoe UI", 10, "bold"), relief="flat", 
                                   padx=20, pady=25, cursor="hand2", 
                                   command=self.reset_app)
        self.btn_reset.pack(side="left", fill="y", padx=(0, 10))

        self.btn_export = tk.Button(controls, text="💾 DOWNLOAD EXCEL", 
                                    bg=self.c_success, fg="white",
                                    font=("Segoe UI", 10, "bold"), relief="flat", 
                                    padx=20, pady=25, cursor="hand2", 
                                    state="disabled", 
                                    command=self.export_to_excel)
        self.btn_export.pack(side="left", fill="y")

        # Progress Status
        progress_frame = tk.Frame(self.root, bg="#F4F5F7")
        progress_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.lbl_status = tk.Label(progress_frame, text="Status: Ready", bg="#F4F5F7", fg="#7F8C8D", font=("Segoe UI", 10))
        self.lbl_status.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate", style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=5)

        # Result Table
        table_frame = tk.Frame(self.root, bg="white")
        table_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        sy = ttk.Scrollbar(table_frame, orient="vertical")
        sx = ttk.Scrollbar(table_frame, orient="horizontal")
        
        self.cols = ["Case Number", "Case Type Number", "Deskripsi Case", "Opened Date", "Unit Kerja", "Sumber"]
        self.tree = ttk.Treeview(table_frame, columns=self.cols, show="headings", 
                                 yscrollcommand=sy.set, xscrollcommand=sx.set)
        
        sy.config(command=self.tree.yview)
        sx.config(command=self.tree.xview)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        for c in self.cols:
            self.tree.heading(c, text=c, anchor="w")
            if c == "Sumber":
                self.tree.column(c, width=250)
            else:
                self.tree.column(c, width=180)

        self.tree.bind("<Control-c>", self.copy_tree)

    # =========================================================================
    # [ HELPER FUNCTIONS ]
    # =========================================================================

    def clean_error_msg(self, error_obj):
        """
        Membersihkan pesan error dari Exception Object agar human-readable.
        Menangani error spesifik COM Interop Excel dan Exception buatan sendiri.
        """
        raw_msg = str(error_obj)
        if "ShowDetail" in raw_msg:
            return "Gagal Drill-Down: Cell bukan Pivot Table atau sheet terkunci."
        if "tidak ketemu" in raw_msg:
            return raw_msg 
        if "Exception occurred" in raw_msg or "-2147" in raw_msg:
            match = re.search(r"Microsoft Excel', '(.*?)',", raw_msg)
            if match:
                err_desc = match.group(1)
                if "ShowDetail" in err_desc:
                    return "Gagal Drill-Down: Cell bukan bagian Pivot Table."
                return f"Excel Error: {err_desc}"
            return "Gagal interaksi Excel (COM Error)."
        return raw_msg

    def normalize_val(self, val):
        """
        Normalisasi data mentah dari Excel menjadi String bersih.
        Menangani kasus:
        - Float dengan desimal .0 (8204.0 -> '8204')
        - NoneType -> ''
        - Whitespace cleaning
        """
        if val is None:
            return ""
        if isinstance(val, (float, int)):
            if float(val).is_integer():
                return str(int(val)) # Convert 8204.0 -> "8204"
            else:
                return str(val)
        return str(val).strip()

    # =========================================================================
    # [ CORE LOGIC: THREADING & WORKER ]
    # =========================================================================

    def start_thread_process(self):
        """
        Memulai Thread baru untuk pemrosesan file agar GUI tidak 'Not Responding'.
        Melakukan validasi input sebelum memulai worker.
        """
        row_kw = self.entry_row.get()
        col_kw = self.entry_col.get()
        code_kw = self.entry_code.get().strip()
        
        if not code_kw:
             messagebox.showwarning("Warning", "Kode Filter tidak boleh kosong.")
             return

        logging.info(f"[>] START BATCH. Params: Row='{row_kw}', Col='{col_kw}', Code='{code_kw}'")
        
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsb")])
        if not file_paths: return

        self.is_processing = True
        self.btn_run.config(state="disabled", bg="#95A5A6")
        self.btn_reset.config(state="disabled")
        self.btn_export.config(state="disabled")
        
        # Reset tracking
        self.progress_bar["maximum"] = len(file_paths)
        self.progress_bar["value"] = 0
        self.failed_files = [] 
        # Note: self.master_data dan self.seen_cache TIDAK di-reset agar bisa akumulasi (atau reset manual via tombol)

        # Spawn Thread
        t = threading.Thread(target=self.worker_process, args=(file_paths, row_kw, col_kw, code_kw))
        t.daemon = True
        t.start()

    # =========================================================================

    def worker_process(self, file_paths, row_kw, col_kw, filter_code):
        """
        Fungsi utama yang berjalan di background thread.
        Mengelola lifecycle aplikasi Excel (xlwings) dan iterasi file.
        """
        app = None
        try:
            self.update_ui_progress(0, "Membuka Excel Engine...")
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            
            total_files = len(file_paths)
            
            for index, path in enumerate(file_paths):
                filename = os.path.basename(path)
                self.update_ui_progress(index, f"Processing ({index+1}/{total_files}): {filename}")
                logging.info(f"[>] FILE START: {filename}")
                
                try:
                    self.process_single_file(app, path, row_kw, col_kw, filter_code)
                    logging.info(f"[+] FILE DONE: {filename}")
                except Exception as e:
                    logging.error(f"[!] FATAL ERROR {filename}: {str(e)}")
                    self.failed_files.append({'file': filename, 'msg': self.clean_error_msg(e)})
                
                self.update_ui_progress(index + 1, f"Selesai: {filename}")

        except Exception as e:
            logging.critical(f"[!] SYSTEM CRASH: {str(e)}", exc_info=True)
            self.failed_files.append({'file': "SYSTEM ERROR", 'msg': str(e)})
        finally:
            if app:
                try:
                    app.quit()
                except:
                    pass
            self.root.after(0, self.finish_processing)

    # =========================================================================
    # [ CORE LOGIC: FILE & SHEET PROCESSING ]
    # =========================================================================

    def process_single_file(self, app, path, row_regex, col_keyword, filter_code):
        """
        Memproses satu file Excel.
        1. Loop semua sheet.
        2. Skip sheet blacklist (TABLE/TABEL).
        3. Scanning heuristik (Baris-per-Baris dan Kolom-per-Kolom).
        4. Drill Down (Klik 2x) pada titik temu.
        5. Ekstraksi data mentah dari sheet baru.
        """
        wb = app.books.open(path)
        data_found = False
        sheet_errors = []

        try:
            for sheet in wb.sheets:
                # 1. Sheet Filtering
                if sheet.name.strip().upper() in ["TABEL", "TABLE", "SHEET1"]:
                    logging.info(f"[-] IGNORE SHEET: {sheet.name} (Blacklist)")
                    continue

                try:
                    # 2. Raw Data Fetching (Memory Efficient)
                    used_range = sheet.used_range
                    data_val = used_range.value 
                    
                    if not data_val: continue 

                    start_row = used_range.row
                    start_col = used_range.column
                    
                    num_rows = len(data_val)
                    num_cols = len(data_val[0]) if num_rows > 0 else 0

                    target_r = -1
                    target_c = -1
                    
                    # 3. Scanning: Priority Left-First (Untuk Baris)
                    for c in range(num_cols):
                        for r in range(num_rows):
                            val = data_val[r][c]
                            str_val = str(val).strip() if val is not None else ""
                            if re.search(row_regex, str_val, re.IGNORECASE):
                                target_r = start_row + r
                                break
                        if target_r != -1: break
                    
                    # 4. Scanning: Priority Top-First (Untuk Kolom)
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

                    # 5. Interaction
                    target_cell = sheet.cells(target_r, target_c)
                    if target_cell.value is None:
                        logging.warning(f"[-] Empty Target Cell in {sheet.name}")
                        continue 

                    init_sheet_count = len(wb.sheets)
                    target_cell.api.ShowDetail = True 
                    
                    if len(wb.sheets) <= init_sheet_count:
                        logging.warning(f"[-] Drill down failed in {sheet.name} (No new sheet)")
                        continue

                    # 6. Extraction Strategy
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
                    raise Exception(f"Keyword Row '{row_regex}' atau Col '{col_keyword}' tidak ditemukan.")
                else:
                    raise Exception("; ".join(sheet_errors))

        finally:
            wb.close()

    # =========================================================================
    # [ CORE LOGIC: DATA EXTRACTION & DUPLICATE HANDLING ]
    # =========================================================================

    def extract_data_manual(self, raw_data, filename, filter_code):
        """
        Mengekstrak data dari list mentah (raw_data).
        Melakukan validasi regex, normalisasi data, dan PENGECEKAN DUPLIKAT.
        """
        if not raw_data or len(raw_data) < 2:
            logging.warning("[-] Extracted data is empty")
            return

        data_rows = raw_data[1:] # Skip Header
        regex_pattern = re.compile(filter_code, re.IGNORECASE)
        count_found = 0

        for row in data_rows:
            # Padding row jika kolom kurang dari standard
            if len(row) < 19:
                row = row + [None] * (19 - len(row))

            # Ambil Kode (Index 1 / Kolom B)
            raw_code = row[1]
            clean_code = self.normalize_val(raw_code)
            
            # 1. Regex Matching
            if regex_pattern.search(clean_code):
                
                # 2. Extract Fields
                val_a = self.normalize_val(row[0])  # Case Number
                val_b = clean_code                  # Case Type
                val_d = self.normalize_val(row[3])  # Desc
                val_e = self.normalize_val(row[4])  # Date
                val_s = self.normalize_val(row[18]) # Unit Kerja
                
                # 3. DUPLICATE CHECKING (Global Cache)
                # Kunci unik: Kombinasi Case Number + Case Type
                unique_key = (val_a, val_b)
                
                if unique_key in self.seen_cache:
                    prev_file = self.seen_cache[unique_key]
                    logging.info(f"[D] DUPLICATE: {val_a} (Type {val_b}). Existing in: {prev_file}. Ignoring.")
                    continue # Skip data ini
                
                # 4. Save Valid Data
                self.seen_cache[unique_key] = filename # Simpan ke cache
                self.master_data.append([val_a, val_b, val_d, val_e, val_s, filename])
                count_found += 1
                logging.debug(f"[+] ADD: {val_a} | {val_b}")

        logging.info(f"[I] Extracted {count_found} valid rows from {filename}")

    # =========================================================================
    # [ UI UPDATES & EXPORT ]
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
        self.lbl_status.config(text=f"Selesai! Data: {success_cnt}. Errors: {error_cnt}")

        if error_cnt > 0:
            self.show_error_window_gui()
        else:
            messagebox.showinfo("Sukses", f"Proses Selesai.\nTotal Data Unik: {success_cnt}")

    def reset_app(self):
        logging.info("[>] RESET APPLICATION REQUESTED")
        if self.is_processing:
            messagebox.showwarning("Warning", "Tunggu proses selesai dulu.")
            return
        
        # Clear UI
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Clear Data Memory
        self.master_data = []
        self.failed_files = []
        self.seen_cache = {} # Reset memori duplikat
        
        self.btn_export.config(state="disabled")
        self.lbl_status.config(text="Status: Ready (Reset Done)")
        self.progress_bar["value"] = 0
        messagebox.showinfo("Reset", "Data dan Cache Duplikat telah dibersihkan.")

    def show_error_window_gui(self):
        top = tk.Toplevel(self.root)
        top.title("Laporan File Gagal")
        top.geometry("900x500") 
        top.configure(bg="#F4F5F7")

        tk.Label(top, text="⚠️ Laporan Error", bg="#F4F5F7", font=("Segoe UI", 14, "bold"), fg=self.c_danger).pack(pady=(15,5))
        
        frame_table = tk.Frame(top)
        frame_table.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        sy = ttk.Scrollbar(frame_table, orient="vertical")
        sx = ttk.Scrollbar(frame_table, orient="horizontal")

        cols = ["Nama File", "Penyebab Error"]
        tree_err = ttk.Treeview(frame_table, columns=cols, show="headings", yscrollcommand=sy.set, xscrollcommand=sx.set) 
        tree_err.heading("Nama File", text="Nama File")
        tree_err.heading("Penyebab Error", text="Penyebab Error")
        tree_err.column("Nama File", width=300)
        tree_err.column("Penyebab Error", width=500)
        
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
        
        # Karena kita sudah filter duplikat di 'extract_data_manual', 
        # disini tinggal tampilkan saja.
        for row in self.master_data:
            self.tree.insert("", "end", values=row)

    def export_to_excel(self):
        if not self.master_data: return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=f"Rekap_Data_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not file_path: return
        try:
            df = pd.DataFrame(self.master_data, columns=self.cols)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Sukses", f"File berhasil disimpan di:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Gagal Save", str(e))

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