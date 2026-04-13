from openpyxl import load_workbook  # type: ignore
# standard openpyxl code:

print("Welcome to the Excel Data Processor!")
print("Enter the name of the sheet/Control Scheme you want to process (e.g. raM_Dvc_Valve(2.4):")


def list_names(file_path):
    try:
        # Attempt to load the workbook
        workbook = load_workbook(filename=file_path, read_only=True)
        sheet_name = input("Sheet Name: ")

        if sheet_name not in workbook.sheetnames:
            print(f"Sheet '{sheet_name}' not found in the workbook.")
            print("Available sheets:")
            for sheet in workbook.sheetnames:
                print(f"- {sheet}")
            return

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

        if len(names) > 0:
            print(f"Names found in the sheet '{sheet_name}':")
            for name in names:
                print(name)
        else:
            print(f"No names found in the sheet '{sheet_name}'.")

    except FileNotFoundError:
        print(f"File '{file_path}' not found. Please check the file path and try again.")
    except Exception as e:
        print(f"An error occurred: {e}")
    
if __name__ == "__main__":
    file_path = input("Enter the path to the Excel file: ")
    list_names(file_path)