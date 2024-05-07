import os
import sys
import ctypes
from docx2pdf import convert

class Process:

    def __init__(self, dir: str, output: str):
        """Class that holds all necessary methods for processing each docx file found.

        :param dir: directory given with docx files
        :type dir: str
        :param output: output path generated with pdf files
        :type output: str
        """
        self.dir = dir
        self.output = output

    def find_all_docx(self) -> list:
        """Finds all DOCX files within the input folder given

        :return: list of docx files
        :rtype: list
        """
        return [os.path.join(self.dir, file_name) for file_name in os.listdir(self.dir) if file_name.endswith(".docx")]
    
    def find_all_pdf(self) -> list:
        """Finds all PDF files within the output folder generated

        :return: list of pdf files
        :rtype: list
        """
        return [os.path.join(self.output, file_name) for file_name in os.listdir(self.output) if file_name.endswith(".pdf")]
    
    def generate_pdfs(self, docx_files: list, pdf_files: list) -> None:
        """Generates PDFs based on the list of found docx files and the
        current number of PDFs already been generated

        :param docx_files: list of docx files found
        :type docx_files: list
        :param pdf_files: list of pdf files found
        :type pdf_files: list
        """

        leftover = []

        for docx in docx_files:
            counter = 0
            for pdf in pdf_files:
                if os.path.splitext(os.path.basename(docx))[0] == os.path.splitext(os.path.basename(pdf))[0]:
                    counter = counter + 1
                else:
                    pass

            if counter == 0:
                leftover.append(docx)
            else:
                pass

        # print(f"Total: {len(leftover)}")

        try:
            for x in range(len(leftover)):
                base = os.path.splitext(os.path.basename(leftover[x]))[0]
                file = base + ".pdf"
                output_file_path = os.path.join(self.output, file)
                convert(leftover[x], output_file_path)
                # print(f"{base} ...Converted.")

        except BaseException as b:
            ctypes.windll.user32.MessageBoxW(0, f"Process() error encountered. {b}", (sys.exc_info()[1]), "Warning!", 16)
