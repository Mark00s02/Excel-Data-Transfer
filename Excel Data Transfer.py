import pandas as pd
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

KEYWORDS = ["no cng", "no zenput"]

PRIMARY = "#4f46e5"      # Indigo
SECONDARY = "#22c55e"    # Green
ACCENT = "#ec4899"       # Pink
BG = "#f8fafc"           # Light gray
CARD = "#ffffff"

class ExcelTransferApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Transfer System")
        self.root.geometry("820x600")
        self.root.configure(bg=BG)

        self.source_file = ""
        self.dest_file = ""

        self.source_columns = []
        self.dest_columns = []

        self.mappings = {}

        self.build_ui()

    def build_ui(self):
        # ===== HEADER =====
        header = tk.Frame(self.root, bg=PRIMARY, height=70)
        header.pack(fill="x")

        tk.Label(
            header,
            text="📊 Excel Data Transfer System",
            bg=PRIMARY,
            fg="white",
            font=("Segoe UI", 20, "bold")
        ).pack(pady=15)

        # ===== MAIN CONTAINER =====
        container = tk.Frame(self.root, bg=BG)
        container.pack(fill="both", expand=True, padx=20, pady=15)

        # ===== FILE SELECTION CARD =====
        file_card = tk.Frame(container, bg=CARD, bd=0, relief="flat")
        file_card.pack(fill="x", pady=10)

        tk.Label(
            file_card,
            text="📂 Select Excel Files",
            bg=CARD,
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", padx=15, pady=10)

        btn_frame = tk.Frame(file_card, bg=CARD)
        btn_frame.pack(padx=15, pady=10)

        tk.Button(
            btn_frame,
            text="Select Source Excel",
            bg=PRIMARY,
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            relief="flat",
            command=self.load_source
        ).grid(row=0, column=0, padx=10)

        tk.Button(
            btn_frame,
            text="Select Destination Excel",
            bg=ACCENT,
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
            relief="flat",
            command=self.load_dest
        ).grid(row=0, column=1, padx=10)

        # ===== MAPPING CARD =====
        self.mapping_frame = tk.LabelFrame(
            container,
            text="🔗 Column Mapping",
            bg=CARD,
            font=("Segoe UI", 12, "bold"),
            labelanchor="n"
        )
        self.mapping_frame.pack(fill="both", expand=True, pady=15)

        # ===== ACTION BUTTON =====
        tk.Button(
            self.root,
            text="🚀 Run Transfer",
            bg=SECONDARY,
            fg="white",
            font=("Segoe UI", 12, "bold"),
            padx=25,
            pady=10,
            relief="flat",
            command=self.run_transfer
        ).pack(pady=15)

        # ===== FOOTER =====
        footer = tk.Label(
            self.root,
            text="Detects: no cng | no zenput • Preserves dropdowns",
            bg=BG,
            fg="gray",
            font=("Segoe UI", 9)
        )
        footer.pack(pady=5)

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

        tk.Label(
            self.mapping_frame,
            text="Source Column",
            bg=CARD,
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, padx=15, pady=8)

        tk.Label(
            self.mapping_frame,
            text="Destination Column",
            bg=CARD,
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=1, padx=15, pady=8)

        self.mappings.clear()

        for i, src_col in enumerate(self.source_columns):
            tk.Label(
                self.mapping_frame,
                text=src_col,
                bg=CARD,
                anchor="w",
                font=("Segoe UI", 10)
            ).grid(row=i + 1, column=0, sticky="w", padx=15, pady=4)

            combo = ttk.Combobox(
                self.mapping_frame,
                values=self.dest_columns,
                state="readonly",
                width=30
            )
            combo.grid(row=i + 1, column=1, padx=15, pady=4)
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


# ==============================
# RUN APP
# ==============================

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTransferApp(root)
    root.mainloop()
