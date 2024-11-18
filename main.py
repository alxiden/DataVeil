import tkinter as tk
import os
import shutil
from docx import Document
import openpyxl

class DataVeil:
    def __init__(self, root):
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
        redacted_folder = os.path.join(folder, "redacted")
        try:
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
            print(f"Folder '{folder}' not found.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def fileTypes(self, files, folder):
        for file in files:
            file_path = os.path.join(folder, file)
            print(file_path)
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
            else:
                print("File type not supported")

    def redact_txt(self, file):
        #print(file)
        strings_to_redact = self.strings_entry.get().split(',')
        try:
            with open(file, 'r') as f:
                content = f.read()
            for string in strings_to_redact:
                content = content.replace(string, "Redacted")
            with open(file, 'w') as f:
                f.write(content)
        except Exception as e:
            print(f"An error occurred while processing the file '{file}': {e}")

    def redact_csv(self, file):
        print(file)
    
    def redact_xlsx(self, file):
        print(file)
        def redact_xlsx(self, file):
            strings_to_redact = self.strings_entry.get().split(',')
            try:
                wb = openpyxl.load_workbook(file)
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows():
                        for cell in row:
                            if cell.value and isinstance(cell.value, str):
                                for string in strings_to_redact:
                                    if string in cell.value:
                                        cell.value = cell.value.replace(string, "Redacted")
                wb.save(file)
            except Exception as e:
                print(f"An error occurred while processing the file '{file}': {e}")
    
    def redact_docx(self, file):
        #print(file)
        strings_to_redact = self.strings_entry.get().split(',')
        try:
            doc = Document(file)
            for paragraph in doc.paragraphs:
                for string in strings_to_redact:
                    if string in paragraph.text:
                        paragraph.text = paragraph.text.replace(string, "Redacted")
            doc.save(file)
        except ValueError as ve:
            print(ve)
        except Exception as e:
            print(f"An error occurred while processing the file '{file}': {e}")

    def redact_pdf(self, file):
        print(file)

def main():
    root = tk.Tk()
    app = DataVeil(root)
    root.mainloop()

if __name__ == '__main__':
    main()