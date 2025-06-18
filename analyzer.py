import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import os

# Optional imports for rare formats
import scipy.io
import pyreadstat
import pyarrow.parquet as pq
import numpy as np

df_global = None  # Global data storage

def load_data(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == '.csv':
            return pd.read_csv(filepath)
        elif ext == '.tsv':
            return pd.read_csv(filepath, sep='\t')
        elif ext in ['.xls', '.xlsx']:
            return pd.read_excel(filepath)
        elif ext == '.json':
            return pd.read_json(filepath)
        elif ext == '.xml':
            return pd.read_xml(filepath)
        elif ext == '.parquet':
            return pd.read_parquet(filepath)
        elif ext == '.orc':
            return pd.read_orc(filepath)
        elif ext == '.h5':
            return pd.read_hdf(filepath)
        elif ext == '.txt':
            return pd.read_csv(filepath)
        elif ext == '.sas7bdat':
            return pd.read_sas(filepath)
        elif ext == '.rdata':
            return pyreadstat.read_rdata(filepath)[0]
        elif ext == '.dta':
            return pd.read_stata(filepath)
        elif ext == '.mat':
            mat = scipy.io.loadmat(filepath)
            return pd.DataFrame({k: v.flatten() for k, v in mat.items() if isinstance(v, np.ndarray)})
        elif ext == '.feather':
            return pd.read_feather(filepath)
        elif ext == '.pickle':
            return pd.read_pickle(filepath)
        elif ext == '.html':
            return pd.read_html(filepath)[0]
        else:
            messagebox.showerror("Unsupported Format", f"File type {ext} is not supported.")
            return None
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return None

def upload_file():
    global df_global
    file_path = filedialog.askopenfilename()
    if not file_path:
        return
    df = load_data(file_path)
    if df is not None:
        df_global = df
        preview_data(df)
        update_dropdowns(df)
        messagebox.showinfo("Success", "File loaded successfully!")

def preview_data(df):
    preview_box.delete(*preview_box.get_children())
    cols = df.columns.tolist()
    preview_box["columns"] = cols[:5]
    for col in cols[:5]:
        preview_box.heading(col, text=col)
    for index, row in df.head(10).iterrows():
        preview_box.insert("", tk.END, values=list(row.values)[:5])

def update_dropdowns(df):
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    col_select1['values'] = numeric_cols
    col_select2['values'] = numeric_cols
    col_select1.set('')
    col_select2.set('')
    plot_type.set('Box Plot')

def analyze():
    global df_global
    if df_global is None:
        messagebox.showerror("Error", "Please upload a file first.")
        return

    col1 = col_select1.get()
    col2 = col_select2.get()
    plot = plot_type.get()

    if plot == "Box Plot":
        df_global[[col1]].plot.box(title=f"Box Plot of {col1}")
    elif plot == "Histogram":
        df_global[[col1]].plot.hist(title=f"Histogram of {col1}", bins=20)
    elif plot == "Line Plot":
        df_global[[col1]].plot(title=f"Line Plot of {col1}")
    elif plot == "Scatter Plot":
        if not col2:
            messagebox.showerror("Missing Input", "Select both X and Y for Scatter plot.")
            return
        df_global.plot.scatter(x=col1, y=col2, title=f"Scatter Plot ({col1} vs {col2})")

    plt.tight_layout()
    plt.show()

def show_stats():
    global df_global
    if df_global is None:
        return
    numeric = df_global.select_dtypes(include='number')
    stats = numeric.describe().T
    stat_window = tk.Toplevel(root)
    stat_window.title("Summary Statistics")

    text = tk.Text(stat_window, wrap="none", font=("Courier", 10))
    text.pack(fill=tk.BOTH, expand=True)
    text.insert(tk.END, stats.to_string())
    text.config(state=tk.DISABLED)

# GUI
root = tk.Tk()
root.title("Advanced Data Analyzer")
root.geometry("850x600")

# Frame
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

btn = tk.Button(frame, text="Upload File", command=upload_file, bg='blue', fg='white')
btn.grid(row=0, column=0, padx=5, pady=5)

tk.Label(frame, text="Plot Type").grid(row=0, column=1)
plot_type = ttk.Combobox(frame, values=["Box Plot", "Histogram", "Line Plot", "Scatter Plot"], state="readonly")
plot_type.grid(row=0, column=2)

tk.Label(frame, text="X / Column").grid(row=0, column=3)
col_select1 = ttk.Combobox(frame, state="readonly")
col_select1.grid(row=0, column=4)

tk.Label(frame, text="Y (if Scatter)").grid(row=0, column=5)
col_select2 = ttk.Combobox(frame, state="readonly")
col_select2.grid(row=0, column=6)

tk.Button(frame, text="Analyze", command=analyze, bg='green', fg='white').grid(row=0, column=7, padx=10)
tk.Button(frame, text="Show Summary", command=show_stats).grid(row=0, column=8)

# Preview Table
preview_box = ttk.Treeview(root)
preview_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

root.mainloop()
