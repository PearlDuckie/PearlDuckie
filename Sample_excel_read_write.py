from openpyxl import Workbook, load_workbook
# Writing to an Excel file
def write_excel(filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = "Name"
    ws['B1'] = "Age"
    ws.append(["Alice", 24])
    ws.append(["Bob", 30])
    wb.save(filename)
    print(f"Data written to {filename}")

# Reading from an Excel file
def read_excel(filename):
    wb = load_workbook(filename)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        print(row)

if __name__ == "__main__":
    file = "sample.xlsx"
    #write_excel(file)
    print("Reading data from the Excel file:")
    read_excel(file)
