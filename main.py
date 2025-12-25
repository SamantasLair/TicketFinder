import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import xlwings as xw
import os
import threading
import re
import logging
from datetime import datetime

# --- KONFIGURASI LOGGING ---
logging.basicConfig(
    filename='debug_log.txt',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class BRIProSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Rekapitulasi Operasional (Dynamic Filter)")
        self.root.geometry("1350x850") # Sedikit lebih tinggi untuk input tambahan
        self.root.configure(bg="#F4F5F7")

        logging.info("=== APLIKASI DIMULAI (Versi Dynamic Code) ===")

        self.c_primary = "#00529C"
        self.c_white = "#FFFFFF"
        self.c_accent = "#F37021"
        self.c_success = "#27AE60"
        self.c_danger = "#C0392B"
        self.c_text = "#2C3E50"

        self.master_data = []
        self.failed_files = [] 
        self.is_processing = False

        self.setup_styles()
        self.create_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", background=self.c_primary, foreground="white",
                        font=("Segoe UI", 10, "bold"), relief="flat")
        style.configure("Treeview", background="white", fieldbackground="white",
                        foreground=self.c_text, rowheight=28, font=("Segoe UI", 10))
        style.configure("Error.Treeview.Heading", background=self.c_danger, foreground="white",
                        font=("Segoe UI", 10, "bold"), relief="flat")
        style.configure("Horizontal.TProgressbar", background=self.c_accent, troughcolor="#E0E0E0")

    def create_ui(self):
        header = tk.Frame(self.root, bg=self.c_primary, height=90)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        
        tk.Label(header, text="BRI Unified Automation", bg=self.c_primary, 
                 fg="white", font=("Segoe UI", 20, "bold")).pack(side="left", padx=25, pady=5)
        
        info_frame = tk.Frame(header, bg=self.c_primary)
        info_frame.pack(side="right", padx=25, pady=15)
        tk.Label(info_frame, text="Dynamic Code Search | Left-Scan Logic", 
                 bg=self.c_primary, fg="#BDC3C7", font=("Segoe UI", 10)).pack(anchor="e")

        controls = tk.Frame(self.root, bg="#F4F5F7")
        controls.pack(fill="x", padx=20, pady=15)
        
        input_frame = tk.Frame(controls, bg="white", highlightbackground="#d9d9d9", highlightthickness=1)
        input_frame.pack(side="left", fill="y", padx=(0, 20))
        
        # 1. Input Row Keyword
        tk.Label(input_frame, text="Kata Kunci Baris (Regex):", bg="white", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_row = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_row.insert(0, r"bandar.*lampung") 
        self.entry_row.grid(row=0, column=1, padx=10, pady=5)

        # 2. Input Col Keyword
        tk.Label(input_frame, text="Kata Kunci Kolom:", bg="white", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_col = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_col.insert(0, "Grand Total")
        self.entry_col.grid(row=1, column=1, padx=10, pady=5)

        # 3. Input Filter Code (NEW FEATURE)
        tk.Label(input_frame, text="Kode Filter (Regex):", bg="white", font=("Segoe UI", 9, "bold"), fg=self.c_accent).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_code = tk.Entry(input_frame, width=30, font=("Segoe UI", 10))
        self.entry_code.insert(0, "8204") # Default Value
        self.entry_code.grid(row=2, column=1, padx=10, pady=5)

        self.btn_run = tk.Button(controls, text="▶ START PROCESS", 
                                 bg=self.c_accent, fg="white",
                                 font=("Segoe UI", 10, "bold"), relief="flat", 
                                 padx=20, pady=25, cursor="hand2", # Padding diperbesar agar imbang
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

        progress_frame = tk.Frame(self.root, bg="#F4F5F7")
        progress_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        self.lbl_status = tk.Label(progress_frame, text="Status: Ready", bg="#F4F5F7", fg="#7F8C8D", font=("Segoe UI", 10))
        self.lbl_status.pack(anchor="w")

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate", style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=5)

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

    def clean_error_msg(self, error_obj):
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

    def reset_app(self):
        logging.info("User menekan tombol RESET")
        if self.is_processing:
            messagebox.showwarning("Warning", "Tunggu proses selesai dulu.")
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.master_data = []
        self.failed_files = []
        self.btn_export.config(state="disabled")
        self.lbl_status.config(text="Status: Ready")
        self.progress_bar["value"] = 0
        messagebox.showinfo("Reset", "Aplikasi berhasil di-reset.")

    def start_thread_process(self):
        row_kw = self.entry_row.get()
        col_kw = self.entry_col.get()
        code_kw = self.entry_code.get().strip() # Ambil Kode Filter Baru
        
        if not code_kw:
             messagebox.showwarning("Warning", "Kode Filter tidak boleh kosong.")
             return

        logging.info(f"START. Regex Row: '{row_kw}', Col: '{col_kw}', Code: '{code_kw}'")
        
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsb")])
        if not file_paths: return

        self.is_processing = True
        self.btn_run.config(state="disabled", bg="#95A5A6")
        self.btn_reset.config(state="disabled")
        self.btn_export.config(state="disabled")
        self.progress_bar["maximum"] = len(file_paths)
        self.progress_bar["value"] = 0
        self.failed_files = [] 

        # Pass 'code_kw' ke worker
        t = threading.Thread(target=self.worker_process, args=(file_paths, row_kw, col_kw, code_kw))
        t.daemon = True
        t.start()

    def worker_process(self, file_paths, row_kw, col_kw, filter_code):
        app = None
        try:
            self.update_ui_progress(0, "Membuka Excel Engine...")
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            
            total_files = len(file_paths)
            
            for index, path in enumerate(file_paths):
                filename = os.path.basename(path)
                self.update_ui_progress(index, f"Memproses ({index+1}/{total_files}): {filename}")
                logging.info(f"--- File: {filename} ---")
                
                try:
                    self.process_single_file(app, path, row_kw, col_kw, filter_code)
                    logging.info("SUKSES.")
                except Exception as e:
                    logging.error(f"GAGAL: {str(e)}", exc_info=True)
                    clean_msg = self.clean_error_msg(e)
                    self.failed_files.append({'file': filename, 'msg': clean_msg})
                
                self.update_ui_progress(index + 1, f"Selesai: {filename}")

        except Exception as e:
            logging.critical("CRITICAL ERROR", exc_info=True)
            self.failed_files.append({'file': "SYSTEM", 'msg': str(e)})
        finally:
            if app:
                try:
                    app.quit()
                except:
                    pass
            self.root.after(0, self.finish_processing)

    def process_single_file(self, app, path, row_regex, col_keyword, filter_code):
        wb = app.books.open(path)
        data_found = False
        sheet_errors = []

        try:
            for sheet in wb.sheets:
                if sheet.name.strip().upper() in ["TABEL", "TABLE", "SHEET1"]:
                    logging.warning(f"Skip Blacklist Sheet: {sheet.name}")
                    continue

                try:
                    used_range = sheet.used_range
                    data_val = used_range.value 
                    
                    if not data_val: continue 

                    start_row = used_range.row
                    start_col = used_range.column
                    
                    num_rows = len(data_val)
                    num_cols = len(data_val[0]) if num_rows > 0 else 0

                    target_r = -1
                    target_c = -1
                    
                    # 1. SCAN BARIS (Prioritas Kiri)
                    for c in range(num_cols):
                        for r in range(num_rows):
                            val = data_val[r][c]
                            str_val = str(val).strip()
                            if re.search(row_regex, str_val, re.IGNORECASE):
                                target_r = start_row + r
                                logging.info(f"Ketemu Baris '{str_val}' di (R:{target_r}, C:{start_col+c})")
                                break
                        if target_r != -1: break
                    
                    # 2. SCAN KOLOM (Prioritas Atas)
                    for r in range(num_rows):
                        for c in range(num_cols):
                            val = data_val[r][c]
                            str_val = str(val).strip()
                            if col_keyword.lower() in str_val.lower():
                                target_c = start_col + c
                                logging.info(f"Ketemu Kolom '{str_val}' di (R:{start_row+r}, C:{target_c})")
                                break 
                        if target_c != -1: break 

                    if target_r == -1 or target_c == -1:
                        logging.debug(f"Sheet {sheet.name}: RowIdx={target_r}, ColIdx={target_c}. Next.")
                        continue 

                    target_cell = sheet.cells(target_r, target_c)
                    if target_cell.value is None:
                        logging.warning("Target cell ketemu tapi kosong.")
                        continue 

                    init_sheet_count = len(wb.sheets)
                    target_cell.api.ShowDetail = True 
                    
                    if len(wb.sheets) <= init_sheet_count:
                        logging.warning("Double click tidak memunculkan sheet baru.")
                        continue

                    new_sheet = wb.sheets.active 
                    df = new_sheet.range("A1").options(pd.DataFrame, header=1, index=False, expand='table').value
                    
                    # Pass 'filter_code' ke fungsi ekstraksi
                    self.extract_data(df, os.path.basename(path), filter_code)
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

    def extract_data(self, df, filename, filter_code):
        df = df.fillna("")
        df = df.astype(str)
        
        # Bersihkan data di kolom B (Index 1) dari format ".0"
        col_b_data = df.iloc[:, 1].str.strip().str.replace(r'\.0$', '', regex=True)
        
        # FILTER DINAMIS: Menggunakan input user (filter_code) sebagai Regex
        # regex=True memungkinkan pencarian partial (misal input "8204" cocok dengan "8204-1")
        mask = col_b_data.str.contains(filter_code, na=False, regex=True)
        filtered_df = df[mask]

        if filtered_df.empty:
            logging.info(f"Sheet terbuka, tapi tidak ada kode '{filter_code}'.")
            return

        for _, row in filtered_df.iterrows():
            val_a = str(row.iloc[0]).replace(".0", "").strip() if len(row) > 0 else ""
            val_b = str(row.iloc[1]).replace(".0", "").strip() if len(row) > 1 else ""
            val_d = str(row.iloc[3]).strip() if len(row) > 3 else ""
            val_e = str(row.iloc[4]).strip() if len(row) > 4 else ""
            val_s = str(row.iloc[18]).strip() if len(row) > 18 else ""
            self.master_data.append([val_a, val_b, val_d, val_e, val_s, filename])

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

        success_count = len(self.master_data)
        failed_count = len(self.failed_files)
        self.lbl_status.config(text=f"Selesai! Data: {success_count} baris. Error: {failed_count} file.")

        if failed_count > 0:
            self.show_error_window_gui()
        else:
            messagebox.showinfo("Sukses", "Semua file berhasil diproses 100%!")

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
        if not self.master_data: return
        df = pd.DataFrame(self.master_data, columns=self.cols)
        df = df.drop_duplicates(subset=["Case Number", "Case Type Number"], keep='first')
        self.master_data_clean = df.values.tolist()
        for row in self.master_data_clean:
            self.tree.insert("", "end", values=row)

    def export_to_excel(self):
        if not hasattr(self, 'master_data_clean') or not self.master_data_clean: return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=f"Rekap_Data_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not file_path: return
        try:
            df = pd.DataFrame(self.master_data_clean, columns=self.cols)
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