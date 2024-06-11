import core
import sys
import ctypes
import os
import tkinter as tk
from tkinter.filedialog import askdirectory


def run(dir: str) -> None:
    """Main run function for all processing steps

    :param dir: directory path
    :type dir: str
    """
    output_folder = os.path.abspath(os.path.join(dir, 'PDFs'))

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    p = core.Process(dir, output_folder)
    docx_list = p.find_all_docx()
    pdf_list = p.find_all_pdf()

    # p.check_docx_files(docx_list)
    p.generate_pdfs(docx_list, pdf_list)

if __name__ == '__main__':
    
    try:
        tk.Tk().withdraw()
        dir = askdirectory()
        run(dir)

    except BaseException as b:
        ctypes.windll.user32.MessageBoxW(0, "run() error encountered.", (sys.exc_info()[1]), "Warning!", 16)