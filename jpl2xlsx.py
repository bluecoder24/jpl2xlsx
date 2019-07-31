"""jpl2xlsx - Console application to migrate .jpl files (d.capture batch) to .xlsx

    Author: Jonathan Haist <jonathan.haist@t-online.de>

"""
import openpyxl
from openpyxl import Workbook
import os
exit = 0

def dirsearch(path):
    """ 
       Looks for all .jpl files in the path directory
    """
    for root, dirs, filenames in os.walk(path):
        for filename in filenames:
            if os.path.splitext(filename)[-1] == ".jpl":
                yield os.path.join(root, filename)

def jpl2xlsx(source=os.getcwd(), target="export.xlsx"):
    """Function which does the magic
        
        Args:
            source (String)     Directory of the .jpl file(s), default = execution directory
            target (String)     Path & file name of the .xlsx which is created, default = execution directory, file name = export.xlsx

        Returns:
            msg (String)        Error or hopefully success message 
    
    """ 
    

    #for row in ws.iter_rows(min_row=1, max_col=200, max_row=9999):
        #for cell in row:
            #print(cell)

    

    files = list(dirsearch(source))  
    line_counter = 0
    data = []
    #wb = Workbook()
    #ws = wb.active
    #ws.title = "Data"

    #Reading out the titles and writing them into the data as first data_set
    with open(files[0]) as infile_header:
        data_set = []
        data_set.append("ID")
        for line in infile_header:
            line_counter += 1
            if line_counter >= 9:
                content = line[:-1].split("=")
                title = content[0].strip()
                data_set.append(title)
        line_counter = 0
        data.append(data_set)
    
    #Reading out the values and writing them into the data as data_set
    for file in files:
        with open(file) as infile:
            basename = os.path.basename(file).split(".")
            data_set = []
            for line2 in infile:
                line_counter += 1
                if line_counter >= 9:
                    content = line2[:-1].split("=")
                    content2 = content[1].split("\"")
                    value = content2[1]
                    data_set.append(value)
                elif line_counter == 1:
                    data_set.append(basename[0])
            line_counter = 0    
            #print("----------")
            data.append(data_set)

    print(data)
    #wb.save(target)

print("--- jpl2xlsx ---\n\nConsole application to migrate .jpl files (d.capture batch) to .xlsx\nVersion: 1.0\nAuthor: Jonathan Haist <jonathan.haist@t-online.de>")
print("\n- Commands -\n")
print("go\tRun the migration and export in the execution directory")
print("goc\tRun the migration and export in a custom directory")
print("ex\tExit the program\n")

while exit == 0:
    cmd = input("Please enter a command: ").lower()
    if cmd == "go":
        jpl2xlsx()
    elif cmd == "goc":
        u_source = input("Enter the source directory: ")
        u_target = input("Enter the target directory: ")
        jpl2xlsx(u_source, u_target)
    elif cmd == "ex":
        exit = 1
    else:
        print("Please enter a valid command!")