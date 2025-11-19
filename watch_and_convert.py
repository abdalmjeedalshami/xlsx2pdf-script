import os
import time
import pythoncom
import win32com.client as win32
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

WATCH_FOLDER = r"D:/IMS Document/zz_watch_and_convert/input"
OUTPUT_FOLDER = r"D:/IMS Document/zz_watch_and_convert/output"


def convert_excel_to_pdf(xlsx_path):
    pythoncom.CoInitialize()
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(os.path.abspath(xlsx_path))

        # إنشاء نفس المسار داخل OUTPUT_FOLDER
        relative_path = os.path.relpath(xlsx_path, WATCH_FOLDER)
        folder_part = os.path.dirname(relative_path)

        output_folder = os.path.join(OUTPUT_FOLDER, folder_part)
        os.makedirs(output_folder, exist_ok=True)

        pdf_name = os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf"
        pdf_path = os.path.join(output_folder, pdf_name)

        workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_path))

        workbook.Close(False)
        excel.Quit()

        print(f"✅ Converted: {pdf_path}")
    except Exception as e:
        print(f"❌ Error converting {xlsx_path}: {e}")
    finally:
        try:
            del workbook
            del excel
        except:
            pass
        pythoncom.CoUninitialize()


class Handler(FileSystemEventHandler):
    def process(self, event):
        path = event.src_path

        if event.is_directory:
            return
        if "~$" in path:
            return
        if not path.lower().endswith(".xlsx"):
            return

        print(f"📂 Detected: {path}")
        time.sleep(2)
        convert_excel_to_pdf(path)

    def on_created(self, event):
        self.process(event)

    def on_modified(self, event):
        self.process(event)


def convert_existing_files():
    print("📋 Checking existing subfolders for XLSX files...")
    for root, dirs, files in os.walk(WATCH_FOLDER):
        for f in files:
            if f.lower().endswith(".xlsx") and "~$" not in f:
                full_path = os.path.join(root, f)
                print(f"🔄 Converting existing: {full_path}")
                convert_excel_to_pdf(full_path)


def watch_folder():
    os.makedirs(WATCH_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    convert_existing_files()

    obs = Observer()
    # recursive=True لمراقبة كل المجلدات الفرعية مثل 2025-5-20
    obs.schedule(Handler(), WATCH_FOLDER, recursive=True)
    obs.start()

    print(f"👀 Watching folder & subfolders: {WATCH_FOLDER}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
    obs.join()


if __name__ == "__main__":
    watch_folder()
