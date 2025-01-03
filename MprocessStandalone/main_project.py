import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from AGBLProject import process_agbl_project
from Mprocess1 import process_source_file
from Mpworkingdata import process_working_data, load_workbook_and_sheet, get_unique_entries_in_column
from Perk import add_perk_dates_to_working_sheet
from CollectionNotice_2 import process_source_file as process_collection_notice

# Initialize the main window
root = tk.Tk()
root.title("Data Processing Application")
root.geometry("900x700")

notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# Collection Notice Tab
frame_collection = tk.Frame(notebook)
notebook.add(frame_collection, text="Collection Notice")

# Collection Notice 2 Tab
frame_collection2 = tk.Frame(notebook)
notebook.add(frame_collection2, text="Collection Notice 2")

# AGBLProject Tab
frame_agbl = tk.Frame(notebook)
notebook.add(frame_agbl, text="AGBLProject")

# Global variables for Collection Notice
source_file_path = ""
destination_folder_path = ""
working_file_path = ""
perk_file_path = ""
destination_file_name = ""

# Global variables for AGBLProject
agbl_file_path = ""
mi_file_path = ""
agbl_destination_folder = ""

# Global variables for Collection Notice 2
collection2_source_file = ""
collection2_table_file = ""
collection2_destination_folder = ""


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

# Collection Notice Functions

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

def run_mprocess1():
    if not source_file_path or not destination_folder_path:
        messagebox.showerror("Error", "Please select a source file and destination folder.")
        return
    progress_bar_collection.start()
    threading.Thread(target=process_mprocess1).start()

def process_mprocess1():
    try:
        sheet_name = "Page 1"
        process_source_file(source_file_path, sheet_name, destination_folder_path)
        progress_bar_collection.stop()
        messagebox.showinfo("Success", "Mprocess1 completed successfully.")
    except Exception as e:
        progress_bar_collection.stop()
        messagebox.showerror("Error", f"An error occurred: {e}")

def run_mpworkingdata():
    if not working_file_path:
        messagebox.showerror("Error", "Please select the working file.")
        return

    try:
        wb, sheet = load_workbook_and_sheet(working_file_path, "Page 2")
        unique_entries = get_unique_entries_in_column(sheet)

        if not unique_entries:
            messagebox.showinfo("Info", "No unique entries found in Column G.")
            return

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
            progress_bar_collection.start()
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
        process_working_data(working_file_path, "Page 2", selected_values, progress_bar_collection)
        progress_bar_collection.stop()
        messagebox.showinfo("Success", "Mpworkingdata completed successfully.")
    except Exception as e:
        progress_bar_collection.stop()
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
    progress_bar_collection.start()
    threading.Thread(target=process_perk).start()

def process_perk():
    try:
        add_perk_dates_to_working_sheet(
            working_file_path, perk_file_path, f"{destination_folder_path}/{destination_file_name}.xlsx", "WorkingSheet"
        )
        progress_bar_collection.stop()
        messagebox.showinfo("Success", "Perk completed successfully.")
    except Exception as e:
        progress_bar_collection.stop()
        messagebox.showerror("Error", f"An error occurred: {e}")

# AGBLProject Functions

def browse_agbl_file():
    global agbl_file_path
    agbl_file_path = filedialog.askopenfilename(
        title="Select AGBL_MI Source File", filetypes=[("Excel Files", "*.xlsx")]
    )
    lbl_agbl_file.config(text=f"AGBL_MI File: {agbl_file_path}" if agbl_file_path else "No file selected.")

def browse_mi_file():
    global mi_file_path
    mi_file_path = filedialog.askopenfilename(
        title="Select MI Source File", filetypes=[("Excel Files", "*.xlsx")]
    )
    lbl_mi_file.config(text=f"MI File: {mi_file_path}" if mi_file_path else "No file selected.")

def browse_agbl_destination():
    global agbl_destination_folder
    agbl_destination_folder = filedialog.askdirectory(title="Select Destination Folder")
    lbl_agbl_destination.config(text=f"Destination Folder: {agbl_destination_folder}" if agbl_destination_folder else "No folder selected.")


def run_agbl_project():
    if not agbl_file_path or not mi_file_path or not agbl_destination_folder:
        messagebox.showerror("Error", "Please select all required files and folder.")
        return

    try:
        progress_bar_agbl.start()
        process_agbl_project(agbl_file_path, mi_file_path, os.path.join(agbl_destination_folder, "AGBL_MI_updated.xlsx"))
        progress_bar_agbl.stop()
        progress_bar_agbl["value"] = 100
        messagebox.showinfo("Success", "AGBLProject completed successfully.")
    except Exception as e:
        progress_bar_agbl.stop()
        messagebox.showerror("Error", f"An error occurred: {e}")

# Collection Notice 2 Functions
def browse_collection2_source():
    global collection2_source_file
    collection2_source_file = filedialog.askopenfilename(title="Select Source File", filetypes=[("Excel Files", "*.xlsx")])
    lbl_collection2_source.config(text=f"Source File: {collection2_source_file}" if collection2_source_file else "No file selected.")

def browse_collection2_table():
    global collection2_table_file
    collection2_table_file = filedialog.askopenfilename(
        title="Select Table File", filetypes=[("Excel Files", "*.xlsx")]
    )
    lbl_collection2_table.config(
        text=f"Table File: {collection2_table_file}" if collection2_table_file else "No file selected."
    )

def browse_collection2_destination():
    global collection2_destination_folder
    collection2_destination_folder = filedialog.askdirectory(title="Select Destination Folder")
    lbl_collection2_destination.config(
        text=f"Destination Folder: {collection2_destination_folder}" if collection2_destination_folder else "No folder selected."
    )


# def run_collection2():
#     if not collection2_source_file or not collection2_table_file or not collection2_destination_folder:
#         messagebox.showerror("Error", "Please select all required files and destination folder.")
#         return
#         progress_bar_collection2.start()
#         threading.Thread(target=process_collection_notice2).start()
def run_collection2():
    if not collection2_source_file or not collection2_table_file or not collection2_destination_folder:
        messagebox.showerror("Error", "Please select all required files and destination folder.")
        return

    # Start the progress bar
    progress_bar_collection2.start()

    # Run the processing in a separate thread
    threading.Thread(target=process_collection_notice2).start()



def process_collection_notice2():
    try:
        sheet_name = "Page 1"  # Modify this if needed
        table_name = "Table 1"  # Modify this if needed
        process_collection_notice(collection2_source_file, sheet_name, table_name, collection2_destination_folder        )
        progress_bar_collection2.stop()
        messagebox.showinfo("Success", "Collection Notice 2 completed successfully.")
    except Exception as e:
        progress_bar_collection2.stop()
        messagebox.showerror("Error", f"An error occurred: {e}")


# Layout for Collection Notice Tab
frame_source = tk.Frame(frame_collection)
frame_source.pack(pady=10)
tk.Button(frame_source, text="Browse Source File", command=browse_source_file).pack(side=tk.LEFT, padx=5)
lbl_source_file = tk.Label(frame_source, text="No source file selected.")
lbl_source_file.pack(side=tk.LEFT)

frame_dest = tk.Frame(frame_collection)
frame_dest.pack(pady=10)
tk.Button(frame_dest, text="Browse Destination Folder", command=browse_destination_folder).pack(side=tk.LEFT, padx=5)
lbl_destination_folder = tk.Label(frame_dest, text="No destination folder selected.")
lbl_destination_folder.pack(side=tk.LEFT)

frame_working = tk.Frame(frame_collection)
frame_working.pack(pady=10)

tk.Button(frame_working, text="Browse Working File", command=browse_working_file).pack(side=tk.LEFT, padx=5)
lbl_working_file = tk.Label(frame_working, text="No working file selected.")
lbl_working_file.pack(side=tk.LEFT)

frame_perk = tk.Frame(frame_collection)
frame_perk.pack(pady=10)

tk.Button(frame_perk, text="Browse Perk File", command=browse_perk_file).pack(side=tk.LEFT, padx=5)
lbl_perk_file = tk.Label(frame_perk, text="No Perk file selected.")
lbl_perk_file.pack(side=tk.LEFT)

frame_dest_name = tk.Frame(frame_collection)
frame_dest_name.pack(pady=10)

tk.Label(frame_dest_name, text="Destination File Name:").pack(side=tk.LEFT, padx=5)
ent_destination_name = tk.Entry(frame_dest_name)
ent_destination_name.pack(side=tk.LEFT)

progress_bar_collection = ttk.Progressbar(frame_collection, orient="horizontal", mode="determinate", length=400)
progress_bar_collection.pack(pady=20)

tk.Button(frame_collection, text="Run Mprocess1", command=run_mprocess1).pack(pady=10)
tk.Button(frame_collection, text="Run Mpworkingdata", command=run_mpworkingdata).pack(pady=10)
tk.Button(frame_collection, text="Run Perk", command=run_perk).pack(pady=10)

# Layout for Collection Notice 2 Tab
frame_collection2_source = tk.Frame(frame_collection2)
frame_collection2_source.pack(pady=10)
tk.Button(frame_collection2_source, text="Browse Source File", command=browse_collection2_source).pack(side=tk.LEFT, padx=5)
lbl_collection2_source = tk.Label(frame_collection2_source, text="No source file selected.")
lbl_collection2_source.pack(side=tk.LEFT)

frame_collection2_table = tk.Frame(frame_collection2)
frame_collection2_table.pack(pady=10)
tk.Button(frame_collection2_table, text="Browse Table File", command=browse_collection2_table).pack(side=tk.LEFT, padx=5)
lbl_collection2_table = tk.Label(frame_collection2_table, text="No table file selected.")
lbl_collection2_table.pack(side=tk.LEFT)

frame_collection2_dest = tk.Frame(frame_collection2)
frame_collection2_dest.pack(pady=10)
tk.Button(frame_collection2_dest, text="Browse Destination Folder", command=browse_collection2_destination).pack(side=tk.LEFT, padx=5)
lbl_collection2_destination = tk.Label(frame_collection2_dest, text="No destination folder selected.")
lbl_collection2_destination.pack(side=tk.LEFT)

progress_bar_collection2 = ttk.Progressbar(frame_collection2, orient="horizontal", mode="determinate", length=400)
progress_bar_collection2.pack(pady=20)

tk.Button(frame_collection2, text="Run Collection Notice 2", command=run_collection2).pack(pady=10)

# Layout for AGBLProject Tab
frame_agbl_files = tk.Frame(frame_agbl)
frame_agbl_files.pack(pady=10)

tk.Button(frame_agbl_files, text="Browse AGBL_MI File", command=browse_agbl_file).pack(side=tk.LEFT, padx=5)
lbl_agbl_file = tk.Label(frame_agbl_files, text="No AGBL_MI file selected.")
lbl_agbl_file.pack(side=tk.LEFT)

frame_mi_files = tk.Frame(frame_agbl)
frame_mi_files.pack(pady=10)

tk.Button(frame_mi_files, text="Browse MI File", command=browse_mi_file).pack(side=tk.LEFT, padx=5)
lbl_mi_file = tk.Label(frame_mi_files, text="No MI file selected.")
lbl_mi_file.pack(side=tk.LEFT)

frame_agbl_dest = tk.Frame(frame_agbl)
frame_agbl_dest.pack(pady=10)

tk.Button(frame_agbl_dest, text="Browse Destination Folder", command=browse_agbl_destination).pack(side=tk.LEFT, padx=5)
lbl_agbl_destination = tk.Label(frame_agbl_dest, text="No destination folder selected.")
lbl_agbl_destination.pack(side=tk.LEFT)

tk.Button(frame_agbl, text="Run AGBLProject", command=run_agbl_project).pack(pady=20)
progress_bar_agbl = ttk.Progressbar(frame_agbl, orient="horizontal", mode="determinate", length=400)
progress_bar_agbl.pack(pady=20)

root.mainloop()
