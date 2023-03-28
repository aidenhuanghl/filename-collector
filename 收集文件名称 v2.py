
import os
import openpyxl
import tkinter as tk
from tkinter import filedialog


def get_filenames_in_folder(folder_path):
    file_names = []

    # Use os.listdir to list all files and directories in the folder
    for entry in os.listdir(folder_path):
        entry_path = os.path.join(folder_path, entry)

        # Check if the entry is a file using os.path.isfile
        if os.path.isfile(entry_path):
            file_names.append(entry)

    return file_names

#folder_path = input("Please enter the folder path: ")
#folder_path = '/Users/huanghailong/Documents'


# Print the list of file names
#print(file_names)

def write_filenames_to_excel(file_names, excel_file_path):
    # Create a new Excel workbook and add a worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write file names to the first column of the worksheet
    for index, file_name in enumerate(file_names, start=1):
        sheet.cell(row=index, column=1).value = file_name

    # Save the workbook to the specified Excel file path
    workbook.save(excel_file_path)

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window

    print("Please select the folder:")
    folder_path = filedialog.askdirectory()
    print(f"Selected folder: {folder_path}")
    file_names = get_filenames_in_folder(folder_path)
    #file_names = get_filenames_in_folder(folder_path)

    print("Please select the output Excel file path:")
    excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    print(f"Selected output file: {excel_file_path}")

    write_filenames_to_excel(file_names, excel_file_path)
    print(f"File names written to {excel_file_path}")

if __name__ == "__main__":
        main()

excel_file_path = input("Please enter the output Excel file path: ")
    #excel_file_path = '/Users/huanghailong/Documents/textxlsx.xlsx'

write_filenames_to_excel(file_names, excel_file_path)

print(f"File names written to {excel_file_path}")

