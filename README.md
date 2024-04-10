# Inventory-Manager
A simple Python program for managing inventory.

Required Python Packages:

https://pypi.org/project/sv-ttk/

pip install sv-ttk

https://pypi.org/project/openpyxl/

pip install openpyxl

Including the standard built in packages:

- tkinter
- sqlite3
- os
- csv

Opening the program will generate the required sqlite database file in the same directory.

You can generate a CSV import template that can be used to import all the data into the database from a CSV format.

To create an exe for windows use PyInstaller:

PyInstaller --onefile --noconsole --collect-data sv_ttk Inventory_Manager.py
