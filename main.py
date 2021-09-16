import os, shutil, xlrd

# Global Variables
spreadsheet_file = 'test.xlsx'

names = []
numbers = []


def findNameCol(sheet):
    for i in range(sheet.ncols):
        if "NAME" in sheet.cell_value(0, i).upper(): # Programmatically finds column heading 'name' in the spreadsheet, without explicitly asking the user
            name_col = i
    
    return name_col


def findNumbersCol(sheet):
    for i in range(sheet.ncols):
        if "PHONE NO." in sheet.cell_value(0, i).upper(): # Programmatically finds column heading 'name' in the spreadsheet, without explicitly asking the user
            numbers_col = i
    
    return numbers_col


# Gets names from the xlsx file
def getNames():
    
    workbook = xlrd.open_workbook(spreadsheet_file)
    sheet = workbook.sheet_by_index(0) # Getting first sheet of the xlsx workbook

    for i in range(sheet.nrows):
        names.append(sheet.cell_value(i, findNameCol(sheet))) # Extracts names from the column and inserts it into the array
        
    del names[0] # Deleting the column heading

    return names



def getNumbers():
    
    workbook = xlrd.open_workbook(spreadsheet_file)
    sheet = workbook.sheet_by_index(0) # Getting first sheet of the xlsx workbook

    for i in range(sheet.nrows):
        numbers.append(sheet.cell_value(i, findNumbersCol(sheet))) # Extracts names from the column and inserts it into the array
        
    del numbers[0] # Deleting the column heading

    return numbers
        





workbook = xlrd.open_workbook(spreadsheet_file)
sheet = workbook.sheet_by_index(0) # Getting first sheet of the xlsx workbook

file = 'myfile.txt'

# with open(os.path.join(path, file), 'w') as fp:
#         pass 

with open(file, "w") as file:
    file.write("")

f = open("myfile.txt", "w")

nms = getNames()
nums = getNumbers()

for i in range (sheet.nrows - 1):
    f.write("BEGIN:VCARD\n")
    f.write("FN:" + nms[i] + "\n")
    f.write("TEL:" + str(int(nums[i])) + "\n")
    f.write("END:VCARD\n\n\n")

f.close()
os.rename("myfile.txt", "contacts.vcf")
