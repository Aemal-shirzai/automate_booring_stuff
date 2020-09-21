import sys
import openpyxl
import pyinputplus as pyip

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Full Information"

#user selected files will be written here
user_files = []


def add_file():
    """This function takes filename from user and add it to userfiles list"""
    file_name = pyip.inputStr("File Name: ")
    user_files.append(file_name)


def start_processing():
    """
    This function loop thourght userfiles list and write it to the excel file.
    If the file is not found it is skiped.
    At the end it clear user select list.
    """
    global user_files
    #if the user_files list is empty
    if not user_files:
        print("======= Warning: You need atleast one file ============")
    else:
        # set initial col value to 1 and its incremented evry time one correct
        # file is proccessed
        col = 1
        #loop through user selected files
        for file_name in user_files:
            # set row initial value to 1 and it is value is incremented when
            # one files all lines are written.
            # its value is reseted when the every file is proccessed completely
            row = 1
            try:
                f = open(file_name, mode="r", encoding="utf8")
                for line in f:
                    sheet.cell(row=row, column=col, value=line.strip("\n"))
                    row +=1
                col += 1
                print(f"\n====== {file_name} success ====== \n")
            except FileNotFoundError as error:
                print(error)
        #save file
        workbook.save("Full_info.xlsx")
        print("\n========== End ===========\n")
    # clear the selected files
    user_files = []


def stop():
    """This function  exit the system"""
    sys.exit()

# all choices availibe
choices = {"Add File":add_file, "Start Processing": start_processing, "Exit":stop}

while True:
    print(f"\n====== {len(user_files)} file selected ==========")
    choice = pyip.inputMenu(["Add File", "Start Processing", "Exit"],
        numbered=True)
    
    if choice in choices:
        choices[choice]()
    
