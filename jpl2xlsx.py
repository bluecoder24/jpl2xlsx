"""jpl2xlsx - Console application to migrate .jpl files (d.capture batch) to .xlsx

    Author: Jonathan Haist <jonathan.haist@t-online.de>

"""
import openpyxl
from openpyxl import Workbook
import os
exit = 0

def export(target, data):
    """Exports the data to an Excel-file

        Args:
            target (String)     File name or path of the Excel to create
            data (list)         List containing the data to export    

        Returns:
            None

    """
    
    print("Writing worksheet...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    set_id = 0
    value_id = 0
    for row in ws.iter_rows(min_row=1, max_col=200, max_row=1000):
        if((len(data)-1) < set_id):
            continue
        else:
            data_set = data[set_id]
            set_id += 1
            for cell in row:
                if((len(data_set)-1) < value_id):
                    continue
                else:   
                    value = data_set[value_id]
                    value_id += 1
                    cell.value = value
            value_id = 0

    print("Saving...")
    wb.save(target)
    print("Saved to " + target)

def dirsearch(path):
    """Looks for all .jpl files in the path directory

        Args:
            path (String)   Path of the directory

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
            None
    
    """ 

    files = list(dirsearch(source))  
    line_counter = 0
    data = []

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
            data.append(data_set)

    export(target, data)

print("--- jpl2xlsx ---\n\nConsole application to migrate .jpl files (d.capture batch) to .xlsx\nVersion: 1.0\nAuthor: Jonathan Haist <jonathan.haist@t-online.de>")
print("\n- Commands -\n")
print("go\tRead the .jpl-files from the execution directory and save the excel file there.")
print("goc\tRead the .jpl-files from a custom directory and save the excel file there.")
print("ex\tExit the program\n")

while exit == 0:
    cmd = input("Please enter a command: ").lower()
    if cmd == "go":
        jpl2xlsx()
    elif cmd == "goc":
        u_source = input("Enter the source directory: ")
        jpl2xlsx(u_source)
    elif cmd == "ex":
        exit = 1
    else:
        print("Please enter a valid command!")