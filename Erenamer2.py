import os
import re
import shutil
import pandas as pd
import subprocess
import sys
from tkinter import Tk, filedialog, messagebox, Label, Button, Entry, Listbox, END

try:
    import openpyxl
except ImportError:
    print("openpyxl is not installed. Installing now...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        import openpyxl
        print("openpyxl installed successfully!")
    except Exception as e:
        print(f"Failed to install openpyxl: {e}")
        sys.exit(1)

class FileExtractionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Extraction Tool")

        # Variables
        self.from_folders = []
        self.to_folder = ""
        self.excel_path = self.load_last_excel_path()
        self.name_dict = {}
        self.files_copied_on_time = 0
        self.files_copied_too_late = 0

        # UI Components
        self.create_widgets()

    def create_widgets(self):
        # Extract FROM
        Label(self.root, text="Extract FROM:").grid(row=0, column=0, sticky="w")
        self.from_listbox = Listbox(self.root, height=5, width=50)
        self.from_listbox.grid(row=1, column=0, columnspan=3, sticky="w")
        Button(self.root, text="+", command=self.add_from_folder).grid(row=1, column=3, sticky="w")
        Button(self.root, text="Clear", command=self.clear_from_folders).grid(row=1, column=4, sticky="w")

        # Extract TO
        Label(self.root, text="Extract TO:").grid(row=2, column=0, sticky="w")
        self.to_entry = Entry(self.root, width=50)
        self.to_entry.grid(row=3, column=0, columnspan=3, sticky="w")
        Button(self.root, text="Browse", command=self.set_to_folder).grid(row=3, column=3, sticky="w")

        # Excel Filepath
        Label(self.root, text="Source (Excel filepath):").grid(row=4, column=0, sticky="w")
        self.excel_entry = Entry(self.root, width=50)
        self.excel_entry.insert(0, self.excel_path)
        self.excel_entry.grid(row=5, column=0, columnspan=3, sticky="w")
        Button(self.root, text="Browse", command=self.set_excel_path).grid(row=5, column=3, sticky="w")

        # Start Button
        Button(self.root, text="Start", command=self.start_processing).grid(row=6, column=0, columnspan=4)

    def add_from_folder(self):
        folder = filedialog.askdirectory(title="Select a folder to extract FROM")
        if folder:
            self.from_folders.append(folder)
            self.from_listbox.insert(END, folder)

    def clear_from_folders(self):
        self.from_folders = []
        self.from_listbox.delete(0, END)

    def set_to_folder(self):
        folder = filedialog.askdirectory(title="Select a folder to extract TO")
        if folder:
            self.to_folder = folder
            self.to_entry.delete(0, END)
            self.to_entry.insert(0, folder)

    def set_excel_path(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path = file_path
            self.excel_entry.delete(0, END)
            self.excel_entry.insert(0, file_path)
            self.save_last_excel_path(file_path)

    def extract_number_from_folder(self, folder_name):
        match = re.search(r"\b\d{6}\b", folder_name)
        return match.group() if match else None

    def process_files_in_folder(self, folder_path, folder_name, is_telaat=False):
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                new_name = f"{folder_name}_{file}"
                if is_telaat:
                    telaat_folder = os.path.join(self.to_folder, "TE LAAT")
                    os.makedirs(telaat_folder, exist_ok=True)
                    dst_path = os.path.join(telaat_folder, new_name)
                    self.files_copied_too_late += 1
                else:
                    dst_path = os.path.join(self.to_folder, new_name)
                    self.files_copied_on_time += 1

                # Handle duplicate file names
                base, ext = os.path.splitext(dst_path)
                counter = 2
                while os.path.exists(dst_path):
                    dst_path = f"{base}({counter}){ext}"
                    counter += 1

                os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                shutil.copy(file_path, dst_path)

    def process_folder(self, folder_path, is_telaat=False):
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isdir(item_path):
                if item.upper() == "TE LAAT":
                    telaat_folder = os.path.join(self.to_folder, "TE LAAT")
                    os.makedirs(telaat_folder, exist_ok=True)
                    self.process_folder(item_path, is_telaat=True)
                else:
                    folder_number = self.extract_number_from_folder(item)
                    if folder_number and folder_number in self.name_dict:
                        self.process_files_in_folder(item_path, self.name_dict[folder_number], is_telaat=is_telaat)
                    else:
                        self.process_folder(item_path, is_telaat=is_telaat)

    def start_processing(self):
        if not self.from_folders:
            messagebox.showerror("Error", "You must select at least one source folder.")
            return

        if not self.to_folder:
            messagebox.showerror("Error", "You must select a destination folder.")
            return

        if not os.path.exists(self.excel_path):
            messagebox.showerror("Error", f"The Excel file does not exist at: {self.excel_path}")
            return

        try:
            df = pd.read_excel(self.excel_path, usecols=[0, 1], header=0)
            self.name_dict = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return

        self.files_copied_on_time = 0
        self.files_copied_too_late = 0

        for root_folder in self.from_folders:
            self.process_folder(root_folder)

        messagebox.showinfo("Done", f"Files have been extracted and renamed.\n{self.files_copied_on_time} on time, {self.files_copied_too_late} too late.")

    def save_last_excel_path(self, path):
        try:
            documents_path = os.path.join(os.path.expanduser("~"), "Documents")
            file_path = os.path.join(documents_path, "ELORENAMERFILEPATH.txt")
            with open(file_path, "w") as file:
                file.write(path)
        except Exception as e:
            print(f"Error saving last Excel path: {e}")

    def load_last_excel_path(self):
        try:
            documents_path = os.path.join(os.path.expanduser("~"), "Documents")
            file_path = os.path.join(documents_path, "ELORENAMERFILEPATH.txt")
            if os.path.exists(file_path):
                with open(file_path, "r") as file:
                    return file.read().strip()
        except Exception as e:
            print(f"Error loading last Excel path: {e}")
        return ""

if __name__ == "__main__":
    root = Tk()
    app = FileExtractionApp(root)
    root.mainloop()
