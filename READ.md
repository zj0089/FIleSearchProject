# File Search Project

This project allows you to search PDF, Word, and Excel files for a list of specific names.

## Requirements

- Python 3.1.41
- The required libraries in `requirements.txt`.

## Setup Instructions

1. Clone the repository:
   ```bash
   git clone <https://github.com/zj0089/FIleSearchProject.git>
   cd FileSearchProject

2. Install dependencies: Install all required packages by running:

   ```bash
   Copy code
   pip install -r requirements.txt

## Usage

1. Run the app by double-clicking on the executable (`app.exe`) in the `dist` folder or in the main root dire7tctory, `FileSearchProject` folder, or by running:
   ```bash
   python app.py
   ```

2. **Browse Folder**: Click the "Browse" button to select the folder with files for searching.
3. **Enter Names**: Enter names manually or load from a `.txt` file by clicking "Load Names from File."
4. **Search Files**: Click "Search Files" to begin the search. Results will be displayed in the app window.

## Creating an Executable (Optional)

To create an executable that can run without installing Python, use PyInstaller:
```bash
pyinstaller --onefile --hidden-import=PyPDF2 --hidden-import=python-docx --hidden-import=pandas app.py
```

This will generate an `app.exe` file in the `dist` directory.