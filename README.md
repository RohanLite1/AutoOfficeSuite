# Insurance Claim MIS Extractor (Desktop / Localhost App)

A specialized desktop tool for Insurance Surveyors & Loss Assessors to automatically extract data from digital claim PDFs (Bajaj, NIA, Liberty, Sompo, etc.) and append them into your master MIS Excel workbook (`2.mis-2026-2027-.xlsx`).

---

## 🚀 How to Run the App on Localhost

1. **Option A (One-Click Launcher)**:
   * Double-click [run_app.bat](file:///C:/Users/bigma/Desktop/flolder/code/office/run_app.bat).
   * It will automatically launch the server and open your browser at:
     **`http://127.0.0.1:5055`**

2. **Option B (Command Line)**:
   ```cmd
   cd c:\Users\bigma\Desktop\flolder\code\office
   .\.venv\Scripts\python.exe app.py
   ```
   Then open `http://127.0.0.1:5055` in Chrome or Edge.

---

## 📌 Features & Workflow

### 1. Persistent Master Excel
* Go to **"2. Master Excel Settings"** tab.
* Paste or type the full path of your master Excel sheet (e.g., `C:\Users\bigma\Desktop\flolder\2.mis-2026-2027-.xlsx`).
* Click **"Verify & Load Sheets"** and select **`MIS-2026-2027 (2)`**.
* Click **"Save Master Excel Settings"**. 
* **Note**: The app remembers this permanently. You never have to select your file again!

### 2. Drag & Drop PDF Ingestion
* Drop single or batch PDFs into the large drop box.
* The engine automatically identifies the insurance company (e.g. `BAJAJ`, `NIA`, `LIBERTY`, `SOMPO`).

### 3. Bajaj General / Allianz Specific Mapping:
From your **Claim Summary Sheet** (`INTIMATION.pdf`), the engine extracts:
* **Policy Reference Number** $\rightarrow$ **Column L** (`Policy No`)
* **Claim Reference Number** $\rightarrow$ **Column O** (`claim no.`)
* **Policy Holder\\Insured Name** $\rightarrow$ **Column Q** (`Name`)
* **Date of Loss** $\rightarrow$ **Column N** (`date of loss`)
* **Policy Inception Date** $\rightarrow$ **Column M** (`policy start date`)
* **No Claim Bonus (NCB)** $\rightarrow$ **Column J** (`NCB`)
* **Driver Relation with Insured** $\rightarrow$ **Column I** (`DRIVER RELATIO`)
* **Date of survey** $\rightarrow$ **Column P** (Defaults to today's date, e.g. `23-Oct-26`)
* **Ref no.** $\rightarrow$ **Column A** (Auto-increments from the last number in your sheet, e.g. `789`)

### 4. Interactive Review & Verification Grid
* Review all extracted rows before appending.
* Double-click or type in any cell to make manual adjustments or assign surveyors.

### 5. Safe Excel Appending
* Click **"Append Rows to Master Excel"**.
* Automatically detects the last row (e.g., row 789) and writes to row 790.
* Preserves font styling (Calibri 11), borders, alignments, and formulas.
* **File Lock Protection**: Warns you if the file is currently open in Microsoft Excel to prevent access errors.
* **Open Excel Button**: Click the top-bar button to instantly launch the updated Excel file in Microsoft Excel.

---

## 📦 Next Step: Packaging as `.exe`
Once you test and confirm everything works to your satisfaction on localhost, we will run `PyInstaller` to generate a standalone Windows `.exe` desktop application!
