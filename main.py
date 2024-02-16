import sqlite3

import pandas as pd

# Function to create tables in the database
def create_tables():
    conn = sqlite3.connect('mydatabase.db')
    cursor = conn.cursor()

    # Create Customers table
    cursor.execute('''CREATE TABLE IF NOT EXISTS Customers
                      (CustomerID INTEGER PRIMARY KEY AUTOINCREMENT,
                      Name TEXT,
                      Email TEXT)''')

    # Create Orders table
    cursor.execute('''CREATE TABLE IF NOT EXISTS Orders
                      (OrderID INTEGER PRIMARY KEY AUTOINCREMENT,
                      CustomerID INTEGER,
                      Product TEXT,
                      Quantity INTEGER,
                      FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID))''')

    # Create Payments table
    cursor.execute('''CREATE TABLE IF NOT EXISTS Payments
                      (PaymentID INTEGER PRIMARY KEY AUTOINCREMENT,
                      CustomerID INTEGER,
                      Amount REAL,
                      FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID))''')

    conn.commit()
    conn.close()

# Function to extract data from Excel file and insert into tables
def extract_data_from_excel(filename):
    conn = sqlite3.connect('mydatabase.db')

    # Read the Excel file
    excel_data = pd.read_excel(filename, sheet_name=None)

    # Iterate through each sheet
    for sheet_name, data in excel_data.items():
        # Insert data into respective tables
        if sheet_name == 'Customers':
            data.to_sql('Customers', conn, if_exists='replace', index=False)
        elif sheet_name == 'Orders':
            data.to_sql('Orders', conn, if_exists='replace', index=False)
        elif sheet_name == 'Payments':
            data.to_sql('Payments', conn, if_exists='replace', index=False)

    conn.close()

# Main program
def main():
    # Create tables in the database
    create_tables()

    # Specify the path to the Excel file
    excel_file = 'C:\\Users\\username\\PycharmProjects\\Projectname\\excelfilename.xlsx'

    # Extract data from the Excel file and insert into tables
    extract_data_from_excel(excel_file)

# Run the main program
main()
