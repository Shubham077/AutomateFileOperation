import os

from FileOperationsReport import createlogfile
from AutomateExcel import openfile,editfile,printfile,savefile,closefile, unprotectfile,verifyedit


currentpath = "C:\\Users\\Admin\\Desktop\\Shubham"
os.chdir(currentpath)
createlogfile()
files_in_directory = os.listdir(currentpath)                # got the list of all files
list=['o','e','s','c','v']

switch = {                  # switch case implemented using dictionary
    'o' : openfile,
    'e' : editfile,
    'p' : printfile,
    's' : savefile,
    'c' : closefile,
    'u' : unprotectfile,
    'v' : verifyedit
}

for file in files_in_directory:
    file_from_fullpath = os.path.join(currentpath, file)                # concatenating "path+filename" because python needs it in that format
    file_number = files_in_directory.index(file) + 1                # get the number if file we are operating on
    for operations in list:
        func = switch.get(operations, lambda: "Invalid operation")
        func(file_from_fullpath, file_number)
