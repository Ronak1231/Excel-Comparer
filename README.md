# 📊Excel File Comparator

An intuitive web application built with **Streamlit** to compare two Excel files and provide a **detailed report** on their structural and data differences.  
This tool helps quickly identify discrepancies between two versions of an Excel file, making it especially useful for **data validation, auditing, and tracking changes** in templates or reports.

---

## 🚀 Key Features

### 🗂️ Sheet Comparison
- Identifies common sheets between two workbooks.  
- Lists sheets that have been **added or removed**.  
- **New!** Option to ignore sheet names and compare the **first sheet** of each file directly.  

### 📑 Column Header Analysis
- **Presence Check**: Detects columns that are present in one file but not the other.  
- **Order Check**: Verifies if the common columns are in the **exact same sequence**.  

### 📊 In-Depth Data Comparison
- **Modified Cells**: Highlights cells where the data has changed between the two files.  
- **New Rows**: Shows entire rows that exist only in the comparison file.  
- **Deleted Rows**: Shows entire rows that were in the reference file but are missing from the comparison file.  

### ⚙️ Powerful & Flexible Settings
- **Primary Key**: Specify a column with unique IDs (e.g., ProductID or Email) for accurate row matching, even if row order changed.  
- **Case-Insensitivity**: Option to treat column headers like `Name` and `name` as identical.  
- **Compare by Position**: Compare the first sheet of each workbook, regardless of their names (e.g., `Report-Week1.xlsx` vs `Report-Week2.xlsx`).  

---

## 🛠️ How to Use

### Prerequisites
- Python 3.7+  
- pip (Python package installer)  

### 1️⃣ Installation
Clone the repository and install dependencies:

```bash
git clone https://github.com/Ronak1231/Excel-File-Comparator.git
cd Excel-File-Comparator

pip install streamlit pandas openpyxl
```

### 2️⃣ Running the Application
Run the following command:

```bash
streamlit run app.py
```

Your default browser will open with the app running.  

### 3️⃣ Performing a Comparison
- Adjust **Settings** in the sidebar:  
  - Check *Compare by position* if sheet names differ.  
  - Leave *Ignore case in column names* checked (recommended).  
  - Enter a *Primary Key Column* for accurate row matching.  
- Upload Files:  
  - Drag and drop your **Reference File**.  
  - Drag and drop your **Comparison File**.  
- Run Comparison:  
  - Click **🚀 Compare Files** and wait for results.  

---

## 📊 Understanding the Results

- **Sheet Name Analysis**: Shows which sheets were compared, added, or removed.  
- **Detailed Analysis** (per sheet):  
  - ✅ *No differences found*: The sheet matches structurally and in data.  
  - ❗ *Differences found*: Shows discrepancies in columns, order, or row data.  

### Inside Differences Section
- **Column Headers**:  
  - Columns only in Reference.  
  - Columns only in Comparison.  
  - Column Order mismatches.  
- **Row Data**:  
  - Modified Cells (old vs new values).  
  - New Rows in comparison file.  
  - Deleted Rows missing from comparison file.  

---

## 💡 Troubleshooting

- **Blank Results Screen**: Usually means one file is corrupt or password-protected. Re-save the Excel files.  
- **No Common Sheets Found**: Check *Compare by position* if sheet names differ.  

---

## 📄 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for more details.

---

## 👨‍💻 Author

**Ronak Bansal**  
[GitHub Profile](https://github.com/Ronak1231)
