#pip3 install openpyxl
import openpyxl

#specify the files location (or path)
excel_files = ['/home/hunter/Documents/Workspace/Python_code/datasets/SampleData2.xlsx']
#create an empty list to append values later on
values = []
#loop through the files in the "excel_files" list
for file in excel_files:
    wb = openpyxl.load_workbook(file)
    #locate worksheet
    worksheet = wb["SalesOrders"]
    #compute average in G46
    worksheet['G46'] = '=AVERAGE(G3:G45)'
    #save the file
    wb.save(file)
    #locate cell and get its value
    cell_value = worksheet['E36'].value
    #append value to "values" list
    values.append(cell_value)

    #print totals
    print(cell_value)

    