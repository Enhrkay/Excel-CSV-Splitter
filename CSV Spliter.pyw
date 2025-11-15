import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re

def choose_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel or CSV File",
        filetypes=[("Excel or CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
    )
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)
        load_sheets(file_path)

def load_sheets(file_path):
    """Load Excel sheet names if applicable."""
    if file_path.lower().endswith((".xlsx", ".xls")):
        try:
            xls = pd.ExcelFile(file_path)
            sheets = xls.sheet_names
            combo_sheet['values'] = sheets
            combo_sheet.current(0)
            messagebox.showinfo("Sheets Loaded", f"✅ Found {len(sheets)} sheets.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read sheets:\n{e}")
    else:
        combo_sheet['values'] = []
        combo_sheet.set("")
        load_headers(file_path)  # Load headers immediately for CSV

def load_headers(file_path=None):
    """Load headers from the selected sheet or CSV."""
    if not file_path:
        file_path = entry_file.get()

    sheet_name = combo_sheet.get()
    try:
        if file_path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1)
        else:
            try:
                df = pd.read_csv(file_path, nrows=1)
            except:
                df = pd.read_csv(file_path, sep="\t", nrows=1)

        headers = list(df.columns)
        combo_column['values'] = headers
        combo_column.current(0)
        messagebox.showinfo("Headers Loaded", f"✅ Found {len(headers)} columns.")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read file:\n{e}")

def preview_unique():
    file_path = entry_file.get()
    sheet_name = combo_sheet.get()
    column_name = combo_column.get()

    if not file_path:
        messagebox.showwarning("Warning", "Please select a file first.")
        return
    if not column_name:
        messagebox.showwarning("Warning", "Please select a column to preview.")
        return

    try:
        if file_path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        else:
            try:
                df = pd.read_csv(file_path, dtype=str)
            except:
                df = pd.read_csv(file_path, sep="\t", dtype=str)

        df.fillna("", inplace=True)
        unique_values = df[column_name].unique().tolist()

        preview_text.delete("1.0", tk.END)
        preview_text.insert(tk.END, f"Unique values in '{column_name}':\n\n")
        for val in unique_values:
            if val.strip():
                preview_text.insert(tk.END, f"• {val}\n")
        preview_text.insert(tk.END, f"\nTotal unique values: {len(unique_values)}")

    except Exception as e:
        messagebox.showerror("Error", f"Could not preview:\n{e}")

def normalize_phone_columns(df):
    """
    Clean, standardize, and auto-repair all phone/contact columns.
    - Removes non-digit characters.
    - Converts +234... → 0...
    - Keeps leading zero for Nigerian numbers.
    - Auto-repairs short or truncated numbers (e.g. 23512 → 08023512000).
    - Ensures uniform 11-digit format where possible.
    """
    for col in df.columns:
        if re.search(r"phone|mobile|contact", col, re.IGNORECASE):
            cleaned = []
            for val in df[col].astype(str):
                val = str(val).strip()

                # Skip blanks or placeholders
                if val.lower() in ("nan", "none", "", "null"):
                    cleaned.append("")
                    continue

                digits = re.sub(r"[^\d]", "", val)

                # --- Nigerian format cleanup ---
                if digits.startswith("234") and len(digits) >= 11:
                    digits = "0" + digits[-10:]
                elif digits.startswith("0") and len(digits) == 11:
                    pass
                elif len(digits) == 10:
                    digits = "0" + digits
                elif len(digits) > 11:
                    digits = digits[-11:]

                # --- Intelligent repair for short/truncated numbers ---
                elif len(digits) < 10:
                    if digits.startswith("234"):
                        digits = "0" + digits[-10:]
                    elif len(digits) <= 5:
                        digits = "080" + digits.zfill(8)[:8]
                    elif len(digits) <= 8:
                        digits = "080" + digits.zfill(8)
                    else:
                        digits = "0" + digits.zfill(10)

                cleaned.append(digits)
            df[col] = cleaned
    return df

def split_file():
    file_path = entry_file.get()
    sheet_name = combo_sheet.get()
    column_name = combo_column.get()

    if not file_path:
        messagebox.showwarning("Warning", "Please select a file first.")
        return
    if not column_name:
        messagebox.showwarning("Warning", "Please select a column to split by.")
        return
    if not sheet_name and file_path.lower().endswith((".xlsx", ".xls")):
        messagebox.showwarning("Warning", "Please select a sheet to split.")
        return

    try:
        # Read appropriate file
        if file_path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        else:
            try:
                df = pd.read_csv(file_path, dtype=str)
            except:
                df = pd.read_csv(file_path, sep="\t", dtype=str)

        df.fillna("", inplace=True)
        df = normalize_phone_columns(df)

        # Output folder
        output_folder = os.path.join(os.path.dirname(file_path), "split_files")
        os.makedirs(output_folder, exist_ok=True)

        # Use sheet name (or base filename for CSV) + column value
        base_name = sheet_name if sheet_name else os.path.splitext(os.path.basename(file_path))[0]

        unique_values = df[column_name].unique().tolist()
        total = len(unique_values)
        progress["maximum"] = total
        progress["value"] = 0

        for i, value in enumerate(unique_values, 1):
            if not str(value).strip():
                continue
            safe_value = "".join(c for c in str(value) if c.isalnum() or c in (" ", "_", "-")).strip()
            output_file = os.path.join(output_folder, f"{base_name}_{safe_value}.csv")
            df[df[column_name] == value].to_csv(output_file, index=False, encoding="utf-8", quoting=1)

            progress["value"] = i
            root.update_idletasks()

        messagebox.showinfo("Success", f"🎉 {total} files created in:\n{output_folder}")
        progress["value"] = 0

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Excel/CSV Splitter by Column (Sheet + Contact Aware)")
root.geometry("680x680")
root.resizable(False, False)

# === File selection ===
frame_file = tk.Frame(root, pady=10)
frame_file.pack(fill='x', padx=20)
tk.Label(frame_file, text="Excel or CSV File:").pack(anchor='w')
entry_file = tk.Entry(frame_file, width=50)
entry_file.pack(side='left', padx=(0, 10))
tk.Button(frame_file, text="Browse", command=choose_file).pack(side='left')

# === Sheet selection ===
frame_sheet = tk.Frame(root, pady=10)
frame_sheet.pack(fill='x', padx=20)
tk.Label(frame_sheet, text="Select Sheet (for Excel):").pack(anchor='w')
combo_sheet = ttk.Combobox(frame_sheet, state="readonly", width=47)
combo_sheet.pack()
tk.Button(frame_sheet, text="Load Headers", command=load_headers).pack(pady=5)

# === Column selector ===
frame_column = tk.Frame(root, pady=10)
frame_column.pack(fill='x', padx=20)
tk.Label(frame_column, text="Select Column to Split By:").pack(anchor='w')
combo_column = ttk.Combobox(frame_column, state="readonly", width=47)
combo_column.pack()

# === Preview ===
frame_preview = tk.Frame(root, pady=10)
frame_preview.pack(fill='x', padx=20)
tk.Button(frame_preview, text="Preview Unique Values", width=25, command=preview_unique).pack()

# === Preview Text ===
frame_text = tk.Frame(root, pady=5)
frame_text.pack(fill='both', expand=True, padx=20)
preview_text = tk.Text(frame_text, height=10, wrap="word", bg="#F9F9F9")
preview_text.pack(fill='both', expand=True)

# === Progress Bar ===
frame_progress = tk.Frame(root, pady=10)
frame_progress.pack(fill='x', padx=20)
tk.Label(frame_progress, text="Progress:").pack(anchor='w')
progress = ttk.Progressbar(frame_progress, orient="horizontal", mode="determinate", length=560)
progress.pack(fill='x', padx=5, pady=5)

# === Split Button ===
frame_action = tk.Frame(root, pady=20)
frame_action.pack()
tk.Button(frame_action, text="Split File", width=25, bg="#0078D7", fg="white", command=split_file).pack()

# === Footer ===
tk.Label(root, text="Developed by Alim", fg="gray", font=("Arial", 8)).pack(side="bottom", pady=5)

root.mainloop()
