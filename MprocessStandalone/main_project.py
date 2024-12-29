import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from Mprocess1 import process_source_file
from Mpworkingdata import process_working_data, load_workbook_and_sheet, get_unique_entries_in_column
from Perk import add_perk_dates_to_working_sheet

# Initialize the main window
root = tk.Tk()
root.title("Data Processing Application")
root.geometry("800x600")

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
    lbl_source_file.config(text=f"Source File: {source_file_path}" if source_file_path else "No file selected.")


def browse_destination_folder():
    global destination_folder_path
    destination_folder_path = filedialog.askdirectory(title="Select Destination Folder")
    lbl_destination_folder.config(text=f"Destination Folder: {destination_folder_path}" if destination_folder_path else "No folder selected.")


def browse_working_file():
    global working_file_path
    working_file_path = filedialog.askopenfilename(
        title="Select Working File", filetypes=[("Excel Files", "*.xlsx")]
    )
    lbl_working_file.config(text=f"Working File: {working_file_path}" if working_file_path else "No file selected.")


def browse_perk_file():
    global perk_file_path
    perk_file_path = filedialog.askopenfilename(
        title="Select Perk File", filetypes=[("Excel Files", "*.xlsx")]
    )
    lbl_perk_file.config(text=f"Perk File: {perk_file_path}" if perk_file_path else "No file selected.")


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

    try:
        # Load workbook and fetch unique values from Column G in "Page 2"
        wb, sheet = load_workbook_and_sheet(working_file_path, "Page 2")
        unique_entries = get_unique_entries_in_column(sheet)

        if not unique_entries:
            messagebox.showinfo("Info", "No unique entries found in Column G.")
            return

        # Display selection dialog for unique values
        selection_window = tk.Toplevel(root)
        selection_window.title("Select Unique Values")
        selection_window.geometry("400x300")

        tk.Label(selection_window, text="Select Unique Values from Column G:").pack(pady=10)

        selected_values = []

        def confirm_selection():
            selected_values.clear()
            selected_indices = listbox.curselection()
            selected_values.extend([unique_entries[idx] for idx in selected_indices])
            selection_window.destroy()

            # Start processing
            progress_bar.start()
            threading.Thread(target=process_mpworkingdata, args=(selected_values,)).start()

        listbox = tk.Listbox(selection_window, selectmode="multiple")
        for entry in unique_entries:
            listbox.insert(tk.END, entry)
        listbox.pack(pady=10, fill="both", expand=True)

        tk.Button(selection_window, text="Confirm Selection", command=confirm_selection).pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def process_mpworkingdata(selected_values):
    try:
        process_working_data(working_file_path, "Page 2", selected_values, progress_bar)
        progress_bar.stop()
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
