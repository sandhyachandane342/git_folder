import csv
import openpyxl
def read_file(file_path):
    file_extension = file_path.split('.')[-1].lower()
    if file_extension == 'txt':
        with open(file_path, 'r') as file:
             data = file.read()
             print("Text file content:")
             print(data)
    if file_extension == 'csv':
           with open(file_path,'r') as file:
                  reader = csv.reader(file)
                  data = list(reader)
                  print("CSV file content:")
                  for row in data:
                      print(row)
                  column_choice = int(input("Enter the column number you want to read:"))
                  if column_choice < len(data[0]):
                      print(f"Selected column {column_choice}:")
                      column_data = [row[column_choice] for row in data]
                      print(column_data)
                  else:
                      print("Invalid column number.")
                  
    elif file_extension == 'xlsx':
        workbook = openpyxl.load_workbook(file_path, 'r')
        sheet = workbook.active
        data = []
        for row in sheet.iter_rows(values_only = True):
             data.append(row)
        print("Content of the data")
        for row in data:
             print(row)
        column_choice = int(input("Enter the column number you want to read:"))
        if column_choice < len(data[0]):
            print(f"Selected column {column_choice}:")
            column_data = [row[column_choice] for row in data]
            print(column_data)
      
    else:
        print("Unsupported file type.")
file_path = input("Enter the file path:")
read_file(file_path)
