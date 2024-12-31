# Python_accountingApplications_AR_AP

this repo will include all the applcation for AR and AP automation for weekly monthly tasks.


# **AGBL and MainProject Application**

## **Overview**
This project is a Python-based application designed to automate the processing and management of Excel workbooks, specifically for AGBL and MI files. It features a user-friendly GUI and powerful backend logic to streamline data workflows.

---

## **Key Features**

### **1. AGBLProject**
- Processes data in the AGBL_MI workbook using data from the MI workbook.
- Matches entries in column A of AGBL_MI with column A of MI and applies the following logic:
  - Collects all values from column F of MI for matches.
  - Checks specific conditions in column G of MI (e.g., "9" at positions 13 and 14).
  - Writes combined results into column O of AGBL_MI as a semicolon-separated string.
- Updates columns J, K, L, M, and N of AGBL_MI:
  - **Column J**: Person Name (from column M of MI).
  - **Column K**: Project Type ("Project").
  - **Column L**: Script execution date.
  - **Column M**: Match count.
  - **Column N**: INVI_GL_NOS (from column G of MI).
- Generates updated Excel files saved in the destination folder.

### **2. MainProject**
- A centralized GUI to manage AGBLProject and other scripts.
- Features:
  - File browsing for source files (AGBL_MI and MI) and destination folder.
  - Progress bar for user feedback.
  - Integration of multiple functionalities (e.g., Perk Processing, Data Matching).

### **3. GUI**
- User-friendly interface built with `tkinter`.
- File browsing and folder selection dialogs.
- Real-time feedback on processing progress and results.

---

## **Setup Instructions**

### **Requirements**
- Python 3.10 or above
- Required Python packages:
  - `openpyxl`
  - `tkinter`
  - `shutil`

### **Installation**
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/projectname.git
   cd projectname

