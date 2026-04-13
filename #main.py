import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from openpyxl import load_workbook  
# standard openpyxl code:

print("Welcome to the Excel Data Processor!")
print("Enter the name of the sheet/Control Scheme you want to process (e.g. raM_Dvc_Valve(2.4):")


def list_names(file_path, sheet_name=None):
    try:
        # Attempt to load the workbook
        workbook = load_workbook(filename=file_path, read_only=True)

        if sheet_name not in workbook.sheetnames:
            messagebox.showerror("Sheet Not Found", f"Sheet '{sheet_name}' not found in the workbook.")


            print(f"Sheet '{sheet_name}' not found in the workbook.")
            print("Available sheets:")
            for sheet in workbook.sheetnames:
                print(f"- {sheet}")
            return []

        sheet = workbook[sheet_name]

        names = []
        # In case of valves, the names are in column A, starting from row 4.
        # This will be changed to work off color in the future.
        for row in sheet.iter_rows(min_row=4, values_only=True):
            name = row[0]

            if name is None:
                continue

            if isinstance(name, str) and name.startswith("$"):
                continue

            names.append(name)
        return names

    except FileNotFoundError:
        print(f"File '{file_path}' not found. Please check the file path and try again.")
        return []
    except Exception as e:
        print(f"An error occurred: {e}")
        return


#OPEN FILE FUNCTION:
def open_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        list_names(file_path)
    else:
        messagebox.showwarning("No File Selected", "Please select an Excel file to process.")


    names = list_names(file_path)

    listbox.delete(0, tk.END)

    if names:
        for name in names:
            listbox.insert(tk.END, name)
    else:
        listbox.insert(tk.END, "No names found.")








#TKINTER GUI

root = tk.Tk()
root.title("Excel Data Processor")
root.geometry("400x400")


open_button = tk.Button(root, text="Open Excel File", command=open_file)
open_button.pack(pady=20)

list_button = tk.Button(root, text="List Names", command=list_names)
list_button.pack(pady=20)

listbox = tk.Listbox(root, width=50, height=15)
listbox.pack(pady=20)




root.mainloop()

    