import pandas as pd
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

KEYWORDS = ["no cng", "no zenput"]

class ExcelTransferApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Transfer System")
        self.root.geometry("700x500")

        self.source_file = ""
        self.dest_file = ""

        self.source_columns = []
        self.dest_columns = []

        self.mappings = {}

        self.build_ui()

    def build_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        tk.Button(frame, text="Select Source Excel", command=self.load_source).grid(row=0, column=0, padx=5)
        tk.Button(frame, text="Select Destination Excel", command=self.load_dest).grid(row=0, column=1, padx=5)

        self.mapping_frame = tk.LabelFrame(self.root, text="Column Mapping")
        self.mapping_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Button(self.root, text="Run Transfer", bg="green", fg="white",
                  command=self.run_transfer).pack(pady=10)

    def load_source(self):
        self.source_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not self.source_file:
            return
        df = pd.read_excel(self.source_file)
        self.source_columns = list(df.columns)
        self.refresh_mapping_ui()

    def load_dest(self):
        self.dest_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not self.dest_file:
            return
        wb = load_workbook(self.dest_file)
        ws = wb.active
        self.dest_columns = [cell.value for cell in ws[1]]
        self.refresh_mapping_ui()

    def refresh_mapping_ui(self):
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()

        if not self.source_columns or not self.dest_columns:
            return

        tk.Label(self.mapping_frame, text="Source Column").grid(row=0, column=0, padx=10)
        tk.Label(self.mapping_frame, text="Destination Column").grid(row=0, column=1, padx=10)

        self.mappings.clear()

        for i, src_col in enumerate(self.source_columns):
            tk.Label(self.mapping_frame, text=src_col).grid(row=i+1, column=0, sticky="w")

            combo = ttk.Combobox(self.mapping_frame, values=self.dest_columns, state="readonly")
            combo.grid(row=i+1, column=1, padx=5)
            self.mappings[src_col] = combo

    def contains_keyword(self, value):
        if pd.isna(value):
            return False
        return any(k in str(value).lower() for k in KEYWORDS)

    def run_transfer(self):
        if not self.source_file or not self.dest_file:
            messagebox.showerror("Error", "Please select both Excel files.")
            return

        source_df = pd.read_excel(self.source_file)
        wb = load_workbook(self.dest_file)
        ws = wb.active
        dest_headers = [cell.value for cell in ws[1]]

        rows_added = 0

        for _, row in source_df.iterrows():

            # 🔑 CRUCIAL CONDITION
            if not (
                self.contains_keyword(row.iloc[3]) or
                self.contains_keyword(row.iloc[4])
            ):
                continue

            new_row = {}
            for src_col, combo in self.mappings.items():
                dest_col = combo.get()
                if dest_col:
                    new_row[dest_col] = row[src_col]

            write_row = ws.max_row + 1
            for col_idx, header in enumerate(dest_headers, start=1):
                if header in new_row:
                    ws.cell(row=write_row, column=col_idx, value=new_row[header])

            rows_added += 1

        wb.save(self.dest_file)
        messagebox.showinfo("Success", f"{rows_added} rows inserted.")

# ==============================
# RUN APP
# ==============================

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTransferApp(root)
    root.mainloop()
