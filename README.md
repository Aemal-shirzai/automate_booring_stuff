# Automate Booring Staff 

### Challange Question from chapter 13 
### 1- Blank Row Inserter
A program blankRowInserter.py that takes two integers  as command line arguments. Let’s call the first integer N and the second integer M. 
Starting at row N, the program should insert M blank rows into the spreadsheet.

#### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)
2. pyinputplus: Python module for validating user inputs. [install](https://pypi.org/project/PyInputPlus/)

#### Code Description:
- Open the excel file, select active sheet, find max rows and max_cols
- Take input from user to enter 
	1. starting_empty_row = Row number from which he/she wants to add empty rows. Note: this number can not be greather than max_rows
	2. num_of_emtpy_row = How many blank rows he/she want to add
- Check if the starting_empty_row number is greather than max_rows or not
	1. If greater: Then exits the program because its does not sence to add empty rows at end
	2. If not greather: The keep the operations going
- Create new_sheet to store our new formated data
- now we have to copy data from our main sheet to new_sheet 
- copy every row from main sheet to new_sheet
- But if the starting_empty_row reaches then whe should leave  (num_empty_row) rows blank and add our data at the row number (row + num_empty_row)
- And Finnaly save our changes     
 	
### 2- Multiplication Table Maker
A program multiplicationTable.py that takes a number N from the command line and creates an N×N multiplication table in an Excel spreadsheet.

#### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)
2. pyinputplus: Python module for validating user inputs. [install](https://pypi.org/project/PyInputPlus/)

#### Code Description:
- create new workbook object, and select the active sheet
- Take input from user to enter 
	1. num = number of multiplication 
- The First for is the parent loop which every time it completes it set one row header, one column header and all multiplication for one row and all columns
- The inner loop is responsilbe for adding multiplication for one row and all that row's columns
- Finally save the workbook

### 3-Spreadsheet Cell Inverter
A program to invert the row and column of the cells in the spreadsheet. For example, the value at row 5, column 3, will be at row 3, column 5, (and vice versa). 

#### Packages to install
1. openpyxl: Python module for working with excel files. [install](https://openpyxl.readthedocs.io/en/stable/)


 	
