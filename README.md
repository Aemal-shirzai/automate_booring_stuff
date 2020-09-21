# Automate Booring Staff 

### Challange Question from chapter 13 
### 1- Blank Row Inserter
A program blankRowInserter.py that takes two integers  as command line arguments. Let’s call the first integer N and the second integer M. 
Starting at row N, the program should insert M blank rows into the spreadsheet.

#### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)
2. pyinputplus: Python module for validating user inputs. [install](https://pypi.org/project/PyInputPlus/)

#### Code Description:
1. Open the excel file, select active sheet, find max rows and max_cols
2. Take input from user to enter 
- starting_empty_row = Row number from which he/she wants to add empty rows. Note: this number can not be greather than max_rows
- num_of_emtpy_row = How many blank rows he/she want to add
3. Check if the starting_empty_row number is greather than max_rows or not
- If greater: Then exits the program because its does not sence to add empty rows at end
- If not greather: The keep the operations going
4. Create new_sheet to store our new formated data
5. now we have to copy data from our main sheet to new_sheet 
6. copy every row from main sheet to new_sheet
7. But if the starting_empty_row reaches then whe should leave  (num_empty_row) rows blank and add our data at the row number (row + num_empty_row)
8. And Finnaly save our changes     
 	
### 2- Multiplication Table Maker
A program multiplicationTable.py that takes a number N from the command line and creates an N×N multiplication table in an Excel spreadsheet.

#### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)
2. pyinputplus: Python module for validating user inputs. [install](https://pypi.org/project/PyInputPlus/)

#### Code Description:
1. create new workbook object, and select the active sheet
2. Take input from user to enter 
- num = number of multiplication 
3. The First for is the parent loop which every time it completes it set one row header, one column header and all multiplication for one row and all columns
4. The inner loop is responsilbe for adding multiplication for one row and all that row's columns
5. Finally save the workbook

### 3-Spreadsheet Cell Inverter
A program to invert the row and column of the cells in the spreadsheet. For example, the value at row 5, column 3, will be at row 3, column 5, (and vice versa). 

#### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)

### 4- Text Files to Spreadsheet
A program to read in the contents of several text files (Entered by user) and insert those contents into a spreadsheet, with one line of text per row. The lines of the first text file will be in the cells of column A, the lines of the second text file will be in the cells of column B, and so on.

#### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)
2. pyinputplus: Python module for validating user inputs. [install](https://pypi.org/project/PyInputPlus/)

#### Code Description
1. import our packages
2. create blank workbook
3. while its true ask user for inputs
- To Add file
- To start processing the added files
- To exit the program
4. then whe check if the entered input by user is a valid one or not and if it is valid we call the desired function for that
5. Functions:
- add_file: This function ask user for file name and append it to the user_files list.
- stop: This function exit the program and clear every thing 
- start_processing: This function takes user_files list and proccess the list. if the list is empty then print and error.or else it check if file exists or not if it exists then it loops thourgh file and read every file lines and write it to the excel file. And finally we save our excel file. **If File is not found it is escaped**.

### 4- Spreadsheet to Text Files
A program that open a spreadsheet and write the cells of column A into one text file, the cells of column B into another text file, and so on.

### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)
	
#### Code Description
1. This program open an excel file and loop through every row and column
2. First row's each column value will be used as file name
3. In inner loop it will open the desired file and and the cell value to the file

