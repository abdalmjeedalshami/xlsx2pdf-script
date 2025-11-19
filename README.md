# Watch and Convert XLSX to PDF

## Description
This project monitors a specified folder (and its subfolders) for new `.xlsx` files and automatically converts them to PDF using Microsoft Excel COM automation. It also processes existing files when the script starts.

## Features
- Watches a folder and subfolders for new Excel files.
- Converts `.xlsx` files to PDF automatically.
- Maintains folder structure in the output directory.
- Handles existing files on startup.

## Folder Structure
```
zz_watch_and_convert/
│
├── watch_and_convert.py    # Main script
├── input/                  # Folder to watch for Excel files
└── output/                 # Folder where PDFs will be saved
```

## Requirements
- Python 3.7+
- Microsoft Excel (desktop version)
- Python packages:
```
watchdog
pywin32
```

Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage
1. Clone the repository:
```bash
git clone https://github.com/yourusername/watch-and-convert.git
```
2. Navigate to the folder:
```bash
cd watch-and-convert
```
3. Run the script:
```bash
python watch_and_convert.py
```

## Important Notes
- Disable **Protected View** in Excel for automation to work:
  - File → Options → Trust Center → Trust Center Settings → Protected View → uncheck all.
- Excel must be installed and licensed.

## Future Improvements
- Add support for LibreOffice conversion (no Excel required).
- Add logging and error handling.
