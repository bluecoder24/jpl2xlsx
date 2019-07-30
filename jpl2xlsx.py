"""jpl2xlsx - Console application to migrate .jpl files (d.capture batch) to .xlsx

    Author: Jonathan Haist <jonathan.haist@t-online.de>

"""
import xlrd
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
    files = list(dirsearch(source))  

    print(files)

    line_counter = 0

    for file in files:
        with open(file) as infile:
            for line in infile:
                line_counter += 1
                if line_counter >= 9:
                    print(line)
                elif line_counter == 1: 
                    basename = os.path.basename(file).split(".")
                    print(basename[0] + "\n")
                
            print("----------")

print("--- jpl2xlsx ---\n\nConsole application to migrate .jpl files (d.capture batch) to .xlsx\nVersion: 1.0\nAuthor: Jonathan Haist <jonathan.haist@t-online.de>")
print("\n- Commands -\n")
print("go\tRun the migration in the execution directory")
print("goc\tRun the migration in a custom directory")
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