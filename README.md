# OnlineSales.Ai_assignmnet
<h2>Task_1 SQL</h2>

Use each worksheet as a table in a relational database and write an SQL query that generates the output report.

Output (Report)
Fetch top 3 departments along with their name and average monthly salary. Below is the format of the report.

![Report](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/4785608a-6acd-44b5-9f15-43039a82498a)


Below is SQL query for the above task


        SELECT  d.NAME AS DEPT_NAME, AVG(s.AMT_USD) AS AVG_MONTHLY_SALARY

        FROM departments d
        
        JOIN employees e ON d.ID = e.DEPT_ID
        
        JOIN salaries s ON e.ID = s.EMP_ID
        
        GROUP BY d.ID, d.NAME
        
        ORDER BY AVG_MONTHLY_SALARY DESC
        
        LIMIT 3;
        
        
But, I wrote a python script which will convert worksheet.xlsx file into relational MySQL database tables
and insert data into tables
and then sricpt will perfome above SQL qeury on database/tables to Fetch top 3 departments along with their name and average monthly salary
and then genrated result table will be converted into report.xlsx file and report.xlsx will be saved into Ouput folder in the same directory as script

<h2>How to run the Script?</h2>

<h4>Requriements</h4>
 1) Python installed on your System

 2) MySQL

<h4>Instruction</h4>

1) Download/clone this repository to your local computer

2) Go to Task1_SQL folder there you will find "task1_SQL.py" python file

![fileSQL](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/e82bd6b2-cf0c-413d-a04b-01519d0b0a50)

3) Open task1_SQL.py with Code Editor/IDE 
4) Change MySQL database Name , Username, Password according to local System.

![mySQL](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/4071b26a-c23e-4d5a-bace-a56ea1248895)

5) Save the changes and do not change anything another than that.
6) Now open CMD as Administrator excute following command to install pandas and openpyxl libary
   
    a) pip install pandas
    
    b) pip install openpyxl
    
    c) pip install mysql-connector-python (To connect with the database(MySQL))
    
    *Don't forget to create database in your MySQL.

7) after sucessfully installing liabries close the CMD.
8) Go to Task1_SQL folder where "task1_SQL.py" python file is and open CMD or PowerShell in that folder using shift + right click

![powerShell](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/576801bb-b534-4dd9-9604-d70ac59b7b95)

9) In PowerShell or CMD Enter following command to run the python script and this genrate tables in your database and using
   SQL query will will gernate report.xlsx file into Ouput folder inside same folder.
   
   Command -> python task1_SQL.py
   
   press Enter key
   
   ![command](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/a3377d52-adf9-4b89-9609-6a092747b917)

10) it will create folder called Output at same folder as script in which you will find report.xlsx file

![result](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/25c86a77-ceb6-4cbc-996e-1ce3f764d476)

Newly created Ouput folder

![output](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/d02ef4d2-21e1-4ffe-a81b-2a4b08f7e2cc)

Report.xlxs file

![reportxx](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/5fecac08-43f2-4100-8c10-f56e85759ec0)

<h2>Task 2 Script</h2>

<h4>Task</h4>

Use each worksheet as a CSV file and write a script (Bash or Python) that generates the same report. Data is to be read from the CSV files not from a database.

<h4>Solution</h4>

To convert CSV file into data, made uses of pandas and openpyxl.workbook to genrate result into .xlsx file.

Using pandas libary, manage to CSV data into python readable data, by using Merge and Groupby function from liabry

established relations between sheets and genrated desired output

<h3>How to run the Script?</h3>

<h4>Instruction</h4>

 1) Download/clone the repository into local system.
 
 2) Go to Task2_CSV folder where you will find task2_CSV.py file

![Csv_py](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/3e0f4f1d-4e97-4d63-ab73-a4faa7b88246)

 3) Open CMD or PowerShell in that folder (perss shift + right click to open powerShell)

![rightclick](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/fe9c73a9-be91-4e88-95c5-6c63073b3fe8)


4) Use the following command to run or excute the script 

command -> python task2_CSV.py

and hit Enter key

![pycommand](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/d542e813-ce00-4d4c-95cc-9928ec7c0459)

Ater sucessful excecution you will get this message.

![afterCommand](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/4e1e19b3-b7f1-4774-a08f-076a1c63d00b)

And Output folder will be created in your Task2_CVS folder like this

![Outputtask2](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/6bf74f71-4833-4c6e-af0f-aa44f1994beb)

Inside this output folder you will find report.xlsx file

![reporttask2](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/f862fa3f-9c98-4aca-9b0f-7e81083a22dc)
