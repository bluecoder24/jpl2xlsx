"""jpl2xlsx - Console application to migrate .jpl files (d.capture batch) to .xlsx

    Author: Jonathan Haist <jonathan.haist@t-online.de>

"""
import xlrd
exit = 0

def jpl2xlsx(mode=0, source="", target="export.xlsx"):
    """Function which does the magic
        
        Args:
            mode (int)          How migration is done (0 = all .jpl files in a directory into a new .xlsx, 1 = only single file into new .xlsx)
            source (String)     Path of the .jpl file(s), can be a file or a directory, default = execution directory
            target (String)     Path & file name of the .xlsx which is created, default = execution directory, file name = export.xlsx

        Returns:
            msg (String)        Error or hopefully success message 
    
    """ 
    msg = "Tragically an error occured! Error Code: 42"

    return msg

def cc():
    """Function to configurate the arguments for the magic function

    Args:
        none

    Returns:
        none
    
    """
def dc():
   """Function to display the current configuration

    Args:
        none

    Returns:
        mode (int)          How migration is done (0 = all .jpl files in a directory into a new .xlsx, 1 = only single file into new .xlsx)
        source (String)     Path of the .jpl file(s), can be a file or a directory, default = execution directory
        target (String)     Path & file name of the .xlsx which is created, default = execution directory, file name = export.xlsx
    
    """ 

print("--- jpl2xlsx ---\n\nConsole application to migrate .jpl files (d.capture batch) to .xlsx\nVersion: 1.0\nAuthor: Jonathan Haist <jonathan.haist@t-online.de>")
print("\n- Commands -\n")
print("go\tRun the migration with the current (or default) configuration")
print("dc\tDisplay the current configuration")
print("cc\tChange the configuration\n")
print("ex\tExit the program\n")

while exit == 0:
    cmd = input("Please enter a command: ").lower()
    if cmd == "go":
        print(jpl2xlsx(dc()))
    elif cmd == "dc":
        dc()
    elif cmd == "cc":
        cc()
    elif cmd == "ex":
        raise SystemExit
    else:
        print("Please enter a valid command!")