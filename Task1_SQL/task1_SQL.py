import os
import pandas as pd
import mysql.connector
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

# MySQL database connection details
conn = mysql.connector.connect(
    host="localhost",
    database="db1",                          # change according to user Database name
    user="root",                             # MySQL userName (change according to user's Username)
    password="root"                          # MySQL password (change according to user's password)
)

# Converting Departments.xlsx(excel file) to Database table deparments(MySQL)
def convert_Depart_excel_to_database(file_path):
    # Read the Excel file into a pandas DataFrame
    excel_data = pd.read_excel(file_path, sheet_name=None)

    # Create a database connection
    cursor = conn.cursor()

    for sheet_name, sheet_data in excel_data.items():
        # Prepare the table name based on the sheet name
        table_name = sheet_name.lower().replace(' ', '_')

        # Create the table
        create_table_query = f"CREATE TABLE IF NOT EXISTS `{table_name}` ("
        columns = sheet_data.columns.tolist()
        primary_key = columns[0]  # Assuming the first column is the primary key
        create_table_query += f"`{columns[0]}` int, "
        create_table_query += f"`{columns[1]}` VARCHAR(255), "

        create_table_query += f"PRIMARY KEY (`{primary_key}`)"
        create_table_query += ")"

        cursor.execute(create_table_query)

        # Insert data into the table
        insert_query = f"INSERT INTO `{table_name}` VALUES ("

        for _, row in sheet_data.iterrows():
            insert_query_values = ""

            for value in row:
                if pd.isnull(value):
                    insert_query_values += "NULL, "
                elif isinstance(value, str):
                    insert_query_values += f"'{value}', "
                else:
                    insert_query_values += f"{value}, "

            insert_query_values = insert_query_values[:-2]  # Remove the trailing comma and space
            full_insert_query = insert_query + insert_query_values + ")"
            cursor.execute(full_insert_query)

    # Commit the changes and close the database connection
    conn.commit()


def convert_Emp_excel_to_database(file_path):
    # Read the Excel file into a pandas DataFrame
    excel_data = pd.read_excel(file_path, sheet_name=None)

    # Create a database connection
    cursor = conn.cursor()

    for sheet_name, sheet_data in excel_data.items():
        # Prepare the table name based on the sheet name
        table_name = sheet_name.lower().replace(' ', '_')

        # Create the table
        create_table_query = f"CREATE TABLE IF NOT EXISTS `{table_name}` ("
        columns = sheet_data.columns.tolist()
        primary_key = columns[0]  # The first column is the primary key
        foreign_key = columns[2].replace(' ', '_')
        create_table_query += f"`{columns[0]}` int, "
        create_table_query += f"`{columns[1]}` VARCHAR(255), "
        create_table_query += f"`{columns[2].replace(' ', '_')}` int, "

        create_table_query += f"PRIMARY KEY (`{primary_key}`),"
        create_table_query += f"FOREIGN KEY (`{foreign_key}`) REFERENCES Departments(ID)"
        create_table_query += ")"

        cursor.execute(create_table_query)

        # Insert data into the table
        insert_query = f"INSERT INTO `{table_name}` VALUES ("

        for _, row in sheet_data.iterrows():
            insert_query_values = ""

            for value in row:
                if pd.isnull(value):
                    insert_query_values += "NULL, "
                elif isinstance(value, str):
                    insert_query_values += f"'{value}', "
                else:
                    insert_query_values += f"{value}, "

            insert_query_values = insert_query_values[:-2]  # Remove the trailing comma and space
            full_insert_query = insert_query + insert_query_values + ")"
            cursor.execute(full_insert_query)

    # Commit the changes and close the database connection
    conn.commit()



def convert_Sala_excel_to_database(file_path):
    # Read the Excel file into a pandas DataFrame
    excel_data = pd.read_excel(file_path, sheet_name=None)

    # Create a database connection
    cursor = conn.cursor()

    for sheet_name, sheet_data in excel_data.items():

        # Prepare the table name based on the sheet name
        table_name = sheet_name.lower().replace(' ', '_')

        # Create the table
        create_table_query = f"CREATE TABLE IF NOT EXISTS `{table_name}` ("
        columns = sheet_data.columns.tolist()

        foreign_key = columns[0]
        create_table_query += f"`{columns[0]}` int, "
        create_table_query += f"`{columns[1].replace(' ', '_').replace('(','').replace(')','')}` VARCHAR(255), "
        create_table_query += f"`{columns[2].replace(' ', '_').replace('(','').replace(')','')}` int, "

        create_table_query += f"FOREIGN KEY (`{foreign_key}`) REFERENCES Employees(ID)"
        create_table_query += ")"

        cursor.execute(create_table_query)

        # Insert data into the table
        insert_query = f"INSERT INTO `{table_name}` VALUES ("

        for _, row in sheet_data.iterrows():
            insert_query_values = ""

            for value in row:
                if pd.isnull(value):
                    insert_query_values += "NULL, "
                elif isinstance(value, str):
                    insert_query_values += f"'{value}', "
                else:
                    insert_query_values += f"{value}, "

            insert_query_values = insert_query_values[:-2]  # Remove the trailing comma and space
            full_insert_query = insert_query + insert_query_values + ")"
            cursor.execute(full_insert_query)

    # Commit the changes and close the database connection
    conn.commit()


# Example usage
convert_Depart_excel_to_database('Departments.xlsx')
convert_Emp_excel_to_database('Employees.xlsx')
convert_Sala_excel_to_database('Salaries.xlsx')
print("Excel data successfully imported into the MySQL database.")



def fetch_data_from_mysql():
    # Create a database connection
   

    # Fetch data from MySQL using a SQL query
    sql_query = """
        SELECT  d.NAME AS DEPT_NAME, AVG(s.AMT_USD) AS AVG_MONTHLY_SALARY
        FROM departments d
        JOIN employees e ON d.ID = e.DEPT_ID
        JOIN salaries s ON e.ID = s.EMP_ID
        GROUP BY d.ID, d.NAME
        ORDER BY AVG_MONTHLY_SALARY DESC
        LIMIT 3;
    """
    df = pd.read_sql(sql_query, conn)

    # Close the database connection
    conn.close()

    return df

def export_to_excel(dataframe, folder_path, file_path):
    # Create a new Excel workbook
    workbook = Workbook()
    sheet = workbook.active

    # Write the DataFrame to the Excel sheet
    # Write column names to the Excel sheet
    header_font = Font(bold=True)
    for row in dataframe_to_rows(dataframe, header=True, index=False):
        sheet.append(row)

    for cell in sheet[1]:
        cell.font = header_font

    # Create the folder if it doesn't exist
    os.makedirs(folder_path, exist_ok=True)

    # Save the workbook to the specified file path
    file_path = os.path.join(folder_path, file_path)

    # Save the workbook to the specified file path
    workbook.save(file_path)

# Fetch data from MySQL
data = fetch_data_from_mysql()

# Export data to Excel
export_to_excel(data, "./Ouput","report.xlsx")