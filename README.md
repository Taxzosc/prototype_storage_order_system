# prototype_storage_order_system
prototype of a ~~storage and~~ work order system, based on a currently existing program in use at my work that i want to improve in the long term. this project uses openpyxl to handle the excel file, and tkinter for the GUI.

why this is just a prototype:
this current version is only meant as a prototype for me to tinker with and learn how i can manipulate an excel file with programming, and how to make usable guis. the full program will have images of the products entered in the work order system, a way to add products to the database, editing previous orders, handle returns, a way to see the whole history of the excel sheet, and more.

to use this program:
This program works with an excel file (.xslx) with two sheets: a sheet containing a database of products, and a sheet containing the history of product movements. 
You can use your own xlsx file or use the provided .xlsx file
The provided .xslx file in /files that contains two sheets: Database and History.
the first line of each sheet are the titles:
Database:
Product Number, Product Description, Product amount

and

History:
Date, Product Number, Product Description, Amount Ordered, New Amount, Type of order, Order Number

the only package needed to install is Openpyxl
pip install openpyxl

To start using this program, Populate the Database with Products, then go into the Settings.py and change the field TYPE_OF_ORDER to what type of order you are processing. you can leave it as "Order" if wished, it does not affect the program.


or if you are using your own excel file:

go into settings.py and change the variables
EXCEL_FILE = "Warehouse" to EXCEL_FILE = "Insert your own excel file name here.xlsx" 
DATABASE = "Database" to DATABASE = "Insert your database sheet name here"
HISTORY = "History" to HISTORY = "Insert your history sheet name here"

in settings you can also change the amount of fields in the program by changing MAX_ORDER_ROWS

when you populate your Database, make sure to not have product numbers starting with 0's(i believe excel does not allow this natually)

then you run main.py
when you run main.py a gui will pop up, all entry fields are set to 0, it does not matter if you write any number beyond the 0, either excel or tinkter(i suspect excel) does not care about 0's at the beginning of a number, so 0001 becomes 1.

enter a order number into the top field, if the order number has been used before, when you click "find descriptions" the program will check whether the number has been used before. if it has been used there will be text saying "invalid/order number already exists", else if it is valid, there will be text saying "Valid/order number does not already exist"

in the products fields enter the product number and amount, then click "find description". if a description field is blank that means the product is invalid. you can still submit the form but the invalid section will not be entered into the system.
you do not have to fill all 25 fields. the program will only work with the valid fields.
the program does not care if the database amounts will go into the negatives.

if for any reason you need, you can reset the form by clicking "reset form".

when the fields are done and acceptable click submit. then the program will add all the valid fields to the history, save the excel file, then the form resets.


this can be useful for people or small buisness that manually deals with warehouse management, however this is just a personal project to expand my knowledge with a improved version of this coming sometime in the future.




dependencies: 
openpyxl
pip install openpyxl


tkinter - https://tkdocs.com/tutorial/
tkinter usually comes installed with python on windows, but if you need to install it follow the guides over at https://tkdocs.com/tutorial/install.html

re
comes with all python versions.
