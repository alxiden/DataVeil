import time
import tkinter as tk
import os
import shutil
from tkinter import messagebox
from docx import Document
import openpyxl
import csv
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
import extract_msg
import win32com.client
import fitz


class DataVeil:
    def __init__(self, root):
        self.string_storage = []

        self.root = root
        self.root.title("DataVeil")

        self.title_label = tk.Label(root, text="DataVeil")
        self.title_label.grid(row=0, column=0, columnspan=2)

        self.folder_label = tk.Label(root, text="Folder Location:")
        self.folder_label.grid(row=1, column=0)

        self.folder_entry = tk.Entry(root, width=50)
        self.folder_entry.grid(row=1, column=1)

        self.strings_label = tk.Label(root, text="Strings to Redact:")
        self.strings_label.grid(row=2, column=0)

        self.strings_entry = tk.Entry(root, width=50)
        self.strings_entry.grid(row=2, column=1)

        self.redact_button = tk.Button(root, text="Redact", command=self.files)
        self.redact_button.grid(row=3, column=0, columnspan=2)

    def files(self):
        folder = self.folder_entry.get()
        redacted_folder = os.path.join(folder, "Redacted")
        try:
            self.text_var()
            if not os.path.exists(redacted_folder):
                os.makedirs(redacted_folder)
            files = os.listdir(folder)
            for file in files:
                file_path = os.path.join(folder, file)
                if os.path.isfile(file_path):
                    shutil.copy(file_path, redacted_folder)
            redacted_files = os.listdir(redacted_folder)
            self.fileTypes(redacted_files, redacted_folder)
        except FileNotFoundError:
            self.show_error_popup(f"Folder '{folder}' not found.")
        except Exception as e:
            self.show_error_popup(f"An error occurred: {e}")
        
    def text_var(self):
        self.string_storage = []
        self.strings = self.strings_entry.get().split(',')
        for string in self.strings:
            self.string_storage.append(string.lower())
            self.string_storage.append(string.upper())
            self.string_storage.append(string.capitalize())
            self.string_storage.append(string.title())
            self.string_storage.append(string)
        #print(self.string_storage)

        
    def show_popup(self, message):
        messagebox.showinfo("Redaction Complete", message)
    
    def show_error_popup(self, message):
        messagebox.showerror("Error", message)

    def fileTypes(self, files, folder):
        for file in files:
            file_path = os.path.join(folder, file)
            #print(file_path)
            if file.endswith('.txt'):
                self.redact_txt(file_path)
            elif file.endswith('.csv'):
                self.redact_csv(file_path)
            elif file.endswith('.xlsx'):
                self.redact_xlsx(file_path)
            elif file.endswith('.docx'):
                self.redact_docx(file_path)
            elif file.endswith('.pdf'):
                self.redact_pdf(file_path)
            elif file.endswith('.msg'):
                self.redact_msg(file_path)
            elif file.endswith('.HTML') or file.endswith('.html'):
                self.redact_txt(file_path)
            else:
               self.show_error_popup(f"File type {file} not supported")
        self.show_popup("Redaction complete for all files")

    def redact_txt(self, file):
        #print(file)
        strings_to_redact = self.string_storage
        try:
            with open(file, 'r') as f:
                content = f.read()
            for string in strings_to_redact:
                content = content.replace(string, "Redacted")
            with open(file, 'w') as f:
                f.write(content)
        except Exception as e:
            self.show_error_popup(f"An error occurred while processing the file '{file}': {e}")

    def redact_csv(self, file):
        #print(file)
        strings_to_redact = self.string_storage
        try:
            with open(file, 'r', newline='') as f:
                reader = csv.reader(f)
                rows = list(reader)

            for i, row in enumerate(rows):
                for j, cell in enumerate(row):
                    for string in strings_to_redact:
                        if string in cell:
                            rows[i][j] = cell.replace(string, "Redacted")

            with open(file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerows(rows)
        except Exception as e:
            self.show_error_popup(f"An error occurred while processing the file '{file}': {e}")
    
    def redact_xlsx(self, file):
        #print(file)
        strings_to_redact = self.string_storage
        try:
            wb = openpyxl.load_workbook(file)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            cell_value_str = str(cell.value)
                            for string in strings_to_redact:
                                if string in cell_value_str:
                                    cell_value_str = cell_value_str.replace(string, "Redacted")
                            cell.value = cell_value_str
            wb.save(file)
        except Exception as e:
            self.show_error_popup(f"An error occurred while processing the file '{file}': {e}")
    
    def redact_docx(self, file):
        #print(file)
        strings_to_redact = self.string_storage
        try:
            doc = Document(file)
            for paragraph in doc.paragraphs:
                for string in strings_to_redact:
                    if string in paragraph.text:
                        paragraph.text = paragraph.text.replace(string, "Redacted")
            doc.save(file)
        except ValueError as ve:
            self.show_error_popup(ve)
        except Exception as e:
            self.show_error_popup(f"An error occurred while processing the file '{file}': {e}")

    def redact_pdf(self, file):
        strings_to_redact = self.string_storage
        try:
            doc = fitz.open(file)
            for page in doc:
                for string in strings_to_redact:
                    text_instances = page.search_for(string)
                    for inst in text_instances:
                        page.add_redact_annot(inst, fill=(0, 0, 0))  # Redact with black color
                    page.apply_redactions()
            temp_file = file.replace('.pdf', '_redacted.pdf')
            doc.save(temp_file, garbage=4, deflate=True)  # Apply redactions and save to a new file
            doc.close()
            os.replace(temp_file, file)  # Replace the original file with the redacted file
        except Exception as e:
            self.show_error_popup(f"An error occurred while processing the file '{file}': {e}")
    
    def convert_msg_to_docx(self, msg_file, docx_file):
        msg = extract_msg.Message(msg_file)
        doc = Document()
        doc.add_heading('Email Information', level=1)
        doc.add_paragraph(f"Subject: {msg.subject}")
        doc.add_paragraph(f"From: {msg.sender}")
        doc.add_paragraph(f"To: {msg.to}")
        doc.add_paragraph(f"Date: {msg.date}")
        doc.add_heading('Body', level=1)
        doc.add_paragraph(msg.body)
        doc.save(docx_file)
        msg.close()
        return docx_file
    
    def redact_msg(self, file):
        #print(file)
        docx_file = file.replace('.msg', '.docx')
        try:
            self.convert_msg_to_docx(file, docx_file)
            self.redact_docx(docx_file)
            time.sleep(2)
            os.remove(file)
        except Exception as e:
           self.show_error_popup(f"An error occurred while processing the file '{file}': {e}")


def main():
    root = tk.Tk()
    app = DataVeil(root)
    root.mainloop()

if __name__ == '__main__':
    main()