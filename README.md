# Inventory Project

This project processes an inventory spreadsheet to perform various calculations and save the results to a new file.

## Features

- Calculate the number of products per supplier
- Compute the total inventory value per supplier
- Identify products with inventory less than 10
- Add a new column for total inventory value

### Prerequisites

Ensure you have Python and `openpyxl` installed:

pip install openpyxl 

## Running the Script
    1. Place your inventory.xlsx file in the project directory.
    2. Run the script:
         python main.py

### Input Format
The script expects an Excel file (inventory.xlsx) with these columns in Sheet1:

Column A (1): Product Number
Column B (2): Inventory
Column C (3): Price
Column D (4): Supplier

### Output
A new Excel file inventory_with_total_value.xlsx is generated with a new column E (5) that stores the Total Inventory Value for each product.
