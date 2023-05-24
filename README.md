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


