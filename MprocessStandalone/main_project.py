import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from Mprocess1 import process_source_file
from Mpworkingdata import process_working_data
from Perk import add_perk_dates_to_working_sheet

# Initialize the main window
root = tk.Tk()
root.title("Data Processing Application")
root.geometry("700x600")

# Global variables
source_file_path = ""
destination_folder_path = ""
working_file_path = ""
perk_file_path = ""
destination_file_name = ""

# Logging
import logging
logging.basicConfig(
    filename="process_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

def log_message(message, level="info"):
    if level == "info":
        logging.info(message)
    elif level == "error":
        logging.error(message)
    print(message)

# GUI Functions
def browse_source_file():
    global source_file_path
    source_file_path = filedialog.askopenfilename(
        title="Select Source File", filetypes=[("Excel Files", "*.xlsx")]
    )
    if source_file_path:
        lbl_source_file.config(text=f"Source File: {source_file_path}")
    else:
        lbl_source_file.config(text="No file selected.")

def browse_destination_folder():
    global destination_folder_path
    destination_folder_path = filedialog.askdirectory(title="Select Destination Folder")
    if destination_folder_path:
        lbl_destination_folder.config(text=f"Destination Folder: {destination_folder_path}")
    else:
        lbl_destination_folder.config(text="No folder selected.")

def browse_working_file():
    global working_file_path
    working_file_path = filedialog.askopenfilename(
        title="Select Working File", filetypes=[("Excel Files", "*.xlsx")]
    )
    if working_file_path:
        lbl_working_file.config(text=f"Working File: {working_file_path}")
    else:
        lbl_working_file.config(text="No file selected.")

def browse_perk_file():
    global perk_file_path
    perk_file_path = filedialog.askopenfilename(
        title="Select Perk File", filetypes=[("Excel Files", "*.xlsx")]
    )
    if perk_file_path:
        lbl_perk_file.config(text=f"Perk File: {perk_file_path}")
    else:
        lbl_perk_file.config(text="No file selected.")

# Processing Functions
def run_mprocess1():
    if not source_file_path or not destination_folder_path:
        messagebox.showerror("Error", "Please select a source file and destination folder.")
        return
    progress_bar.start()
    threading.Thread(target=process_mprocess1).start()

def process_mprocess1():
    try:
        sheet_name = "Page 1"
        process_source_file(source_file_path, sheet_name, destination_folder_path)
        progress_bar.stop()
        progress_bar["value"] = 100
        messagebox.showinfo("Success", "Mprocess1 completed successfully.")
    except Exception as e:
        progress_bar.stop()
        messagebox.showerror("Error", f"An error occurred: {e}")

def run_mpworkingdata():
    if not working_file_path:
        messagebox.showerror("Error", "Please select the working file.")
        return
    progress_bar.start()
    threading.Thread(target=process_mpworkingdata).start()

def process_mpworkingdata():
    try:
        sheet_name = ent_mp_sheet_name.get().strip()
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty.")
        process_working_data(working_file_path, sheet_name)
        progress_bar.stop()
        progress_bar["value"] = 100
        messagebox.showinfo("Success", "Mpworkingdata completed successfully.")
    except Exception as e:
        progress_bar.stop()
        messagebox.showerror("Error", f"An error occurred: {e}")

def run_perk():
    if not perk_file_path or not working_file_path or not destination_folder_path:
        messagebox.showerror("Error", "Please select all required files and folder.")
        return
    global destination_file_name
    destination_file_name = ent_destination_name.get()
    if not destination_file_name:
        messagebox.showerror("Error", "Please provide a destination file name.")
        return
    progress_bar.start()
    threading.Thread(target=process_perk).start()

def process_perk():
    try:
        add_perk_dates_to_working_sheet(
            working_file_path, perk_file_path, f"{destination_folder_path}/{destination_file_name}.xlsx", "WorkingSheet"
        )
        progress_bar.stop()
        progress_bar["value"] = 100
        messagebox.showinfo("Success", "Perk completed successfully.")
    except Exception as e:
        progress_bar.stop()
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI Layout
frame_source = tk.Frame(root)
frame_source.pack(pady=10)
tk.Button(frame_source, text="Browse Source File", command=browse_source_file).pack(side=tk.LEFT, padx=5)
lbl_source_file = tk.Label(frame_source, text="No source file selected.")
lbl_source_file.pack(side=tk.LEFT)

frame_dest = tk.Frame(root)
frame_dest.pack(pady=10)
tk.Button(frame_dest, text="Browse Destination Folder", command=browse_destination_folder).pack(side=tk.LEFT, padx=5)
lbl_destination_folder = tk.Label(frame_dest, text="No destination folder selected.")
lbl_destination_folder.pack(side=tk.LEFT)

frame_working = tk.Frame(root)
frame_working.pack(pady=10)
tk.Button(frame_working, text="Browse Working File", command=browse_working_file).pack(side=tk.LEFT, padx=5)
lbl_working_file = tk.Label(frame_working, text="No working file selected.")
lbl_working_file.pack(side=tk.LEFT)

frame_mp_sheet = tk.Frame(root)
frame_mp_sheet.pack(pady=10)
tk.Label(frame_mp_sheet, text="Mpworkingdata Sheet Name:").pack(side=tk.LEFT, padx=5)
ent_mp_sheet_name = tk.Entry(frame_mp_sheet)
ent_mp_sheet_name.pack(side=tk.LEFT)

frame_perk = tk.Frame(root)
frame_perk.pack(pady=10)
tk.Button(frame_perk, text="Browse Perk File", command=browse_perk_file).pack(side=tk.LEFT, padx=5)
lbl_perk_file = tk.Label(frame_perk, text="No Perk file selected.")
lbl_perk_file.pack(side=tk.LEFT)

frame_dest_name = tk.Frame(root)
frame_dest_name.pack(pady=10)
tk.Label(frame_dest_name, text="Destination File Name:").pack(side=tk.LEFT, padx=5)
ent_destination_name = tk.Entry(frame_dest_name)
ent_destination_name.pack(side=tk.LEFT)

progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=400)
progress_bar.pack(pady=20)

frame_options = tk.Frame(root)
frame_options.pack(pady=10)
tk.Button(frame_options, text="Run Mprocess1", command=run_mprocess1).pack(side=tk.LEFT, padx=10)
tk.Button(frame_options, text="Run Mpworkingdata", command=run_mpworkingdata).pack(side=tk.LEFT, padx=10)
tk.Button(frame_options, text="Run Perk", command=run_perk).pack(side=tk.LEFT, padx=10)

root.mainloop()
