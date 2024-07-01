import os
import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert
from concurrent.futures import ThreadPoolExecutor

def convert_doc_to_pdf(doc_path):
    pdf_path = os.path.splitext(doc_path)[0] + '.pdf'
    try:
        convert(doc_path, pdf_path)
        print(f"Converted: {doc_path} -> {pdf_path}")
    except Exception as e:
        print(f"Error converting {doc_path}: {str(e)}")

def select_files():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_paths = filedialog.askopenfilenames(filetypes=[("Word Document", "*.doc;*.docx")])
    return file_paths

def main():
    print("Select the DOC/DOCX files you want to convert to PDF.")
    doc_files = select_files()
    
    if not doc_files:
        print("No files selected. Exiting.")
        return

    print(f"Selected {len(doc_files)} files for conversion.")
    
    with ThreadPoolExecutor() as executor:
        executor.map(convert_doc_to_pdf, doc_files)
    
    print("Conversion process completed.")

if __name__ == "__main__":
    main()
