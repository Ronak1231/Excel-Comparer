# 📊 Excel File Comparator

A simple and intuitive web application built with **Streamlit** to compare two Excel files, highlighting structural differences in sheet names and column headers. This tool helps ensure consistency between different versions of Excel reports, templates, or data exports.

---

## ✨ Features

- **User-Friendly Interface**: Simple drag-and-drop file uploaders for a seamless user experience.  
- **Sheet Name Analysis**:  
  - Identifies sheets that are common to both files.  
  - Lists sheets that exist only in the reference file.  
  - Lists sheets that exist only in the comparison file.  
- **Column Header Analysis**:  
  - For each common sheet, it performs a detailed comparison of column headers.  
  - Highlights columns present in the reference file but missing from the comparison file.  
  - Identifies new columns added to the comparison file.  
- **Efficient Processing**: Uses **pandas** to read only the necessary metadata (sheet names and headers), ensuring fast performance even with large files.  
- **Clear Results**: Presents the comparison results in a clean, expandable, and easy-to-understand format.  

---

## 🛠️ Technology Stack

- **Python**: The core programming language.  
- **Streamlit**: For building the interactive web application UI.  
- **Pandas**: For powerful and efficient Excel file manipulation.  
- **openpyxl**: Required by Pandas as a backend engine for reading `.xlsx` files.  

---

## 🚀 Getting Started

Follow these instructions to set up and run the project on your local machine.

### Prerequisites

Make sure you have the following installed:

- Python 3.8 or higher  
- pip (Python package installer)  

### Installation & Setup

1. Clone the repository:

```bash
git clone https://github.com/Ronak1231/Excel-Comparer.git
cd Excel-Comparer
```

2. Create a virtual environment (recommended):

**On macOS and Linux:**

```bash
python3 -m venv venv
source venv/bin/activate
```

**On Windows:**

```bash
python -m venv venv
.env\Scriptsctivate
```

3. Install the required packages:  
   Create a file named `requirements.txt` and add the following lines:

```
streamlit
pandas
openpyxl
```

   Then, install them using pip:

```bash
pip install -r requirements.txt
```

### Running the Application

Navigate to the project directory in your terminal and run:

```bash
streamlit run app.py
```

Your default web browser will automatically open with the application running.

---

## 📘 Usage

1. **Upload Files**: Use the file uploaders on the left and right to select your **Reference File** and **Comparison File**. The application supports both `.xlsx` and `.xls` formats.  
2. **Compare**: Click the **🚀 Compare Files** button.  
3. **Review Results**: The results will appear below the button, neatly organized into **Sheet Name Analysis** and **Column Header Analysis**. Use the expanders in the column analysis section to view details for each common sheet.  

---

## 📄 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for more details.

---

## 🙌 Contributing

Contributions, issues, and feature requests are welcome! Feel free to open an issue or submit a pull request.

---

## 👨‍💻 Author

**Ronak Bansal**  
[GitHub Profile](https://github.com/Ronak1231)

