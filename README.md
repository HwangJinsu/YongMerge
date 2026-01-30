# ✨ YongMerge

YongMerge is a 'Mail Merge' automation program that enables you to easily generate **Hangul (HWP)** and **PowerPoint (PPT)** documents using Excel data.

It drastically reduces repetitive document processing time and is optimized for mass-producing data-driven documents such as certificates, report cards, notices, and name tags.

## 🚀 Key Features

*   **📄 Powerful Compatibility**: 
    *   Supports both Hangul (`.hwp`, `.hwpx`) and PowerPoint (`.ppt`, `.pptx`) mail merge functions.
*   **📊 Easy Data Management**: 
    *   Supports importing Excel (`.xlsx`) files.
    *   Allows direct data modification and editing in a spreadsheet format within the program.
    *   Supports adding/deleting rows and columns via right-click, along with Undo/Redo functionality.
*   **🖱️ Intuitive User Experience (UX)**:
    *   **Click-based Field Insertion**: Click field buttons to insert placeholders directly into the document.
    *   **Automated Field Mapping**: Automatically recognizes the `{{FieldName}}` format to substitute data.
*   **🖼️ Automatic Image Insertion**:
    *   Automatically inserts images (photos, signatures, logos, etc.) into specific locations in the document.
    *   Automatic adjustment of image size and ratio (matched to 'Tables' in Hangul and 'Rectangles' in PPT).
*   **💾 Flexible Output Options**:
    *   **Save as Individual Files**: Generates a separate file for each data row.
    *   **Save as Combined File**: Merges all data into a single file.

## 🛠️ System Requirements

*   **OS**: Windows 10 or 11 (Windows environment is required for Hangul/PPT automation).
*   **Required Software**:
    *   Hangul 2010 or higher (for Hangul document automation).
    *   Microsoft PowerPoint (for PPT document automation).

## 📦 How to Run

YongMerge is distributed as a **single executable file**, requiring no separate installation.

1.  Download the distributed `YongMerge.exe` file.
2.  Double-click the downloaded file to run the program.
    *   *If a 'Windows protected your PC' popup appears during the first run, click 'More info' and then click 'Run anyway'.*

---

### 👨‍💻 For Developers (Running from Source Code)

If you wish to modify or run the source code directly, you will need the following environment:

*   **Python**: 3.8 or higher
*   **Install Libraries**: `pip install PyQt5 pywin32 pandas openpyxl Pillow`
*   **Run**: `python main_app.py`

## 📖 Usage Guide

1.  **Launch Program**: Open `YongMerge.exe` to start the application.
2.  **Load Template**: Click 'Select Template File' to choose your template (Hangul or PPT).
3.  **Prepare Data**:
    *   **Upload XLSX**: Click 'Upload XLSX' to load an Excel file containing your data.
    *   **Manual Entry**: Enter names in 'Field Creation' to create columns and input data directly into the table.
4.  **Insert Fields**:
    *   Click the field buttons at the top of the program to insert placeholders (`{{FieldName}}`) at the current cursor position in the document.
    *   In PPT, you can also manually type `{{Name}}`, `{{Address}}`, etc., into text boxes on the slides.
5.  **Insert Images (Optional)**:
    *   Use the 'Add Image' button to load image files; paths will be automatically entered into the 'Image' column.
    *   In PPT, images will be inserted into shapes or boxes containing the `{{Image}}` text placeholder.
6.  **Generate Document**: Click 'Generate Document'. Choose between 'Save as Individual Files' or 'Save as Combined File' to start the process.

## ⚠️ Precautions

*   Please avoid using the mouse or keyboard during the document generation process, as the automation script may take control of them.
*   A security authorization popup may appear during Hangul (HWP) automation; you must click 'Allow All' for the program to function correctly.

## 📄 Open Source Licenses

YongMerge uses the following open-source software and complies with their respective license conditions:

*   **Python 3** (PSF License) - [https://www.python.org/](https://www.python.org/)
*   **PyQt5** (GPL v3) - [https://www.riverbankcomputing.com/software/pyqt/](https://www.riverbankcomputing.com/software/pyqt/)
*   **python-pptx** (MIT License) - [https://python-pptx.readthedocs.io/](https://python-pptx.readthedocs.io/)
*   **Pillow** (HPND License) - [https://python-pillow.org/](https://python-pillow.org/)
*   **pandas** (BSD 3-Clause License) - [https://pandas.pydata.org/](https://pandas.pydata.org/)
*   **PyInstaller** (GPL v2 with exceptions) - [https://pyinstaller.org/](https://pyinstaller.org/)
*   **pywin32 / win32com** (PSF License) - [https://github.com/mhammond/pywin32](https://github.com/mhammond/pywin32)