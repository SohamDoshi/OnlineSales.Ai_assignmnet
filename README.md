# OnlineSales.Ai_assignmnet
Task_1 SQL

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

How to run the Script?

Requriements
 1) Python installed on your System
 2) MySQL

Instruction

1) Download/clone this repository to your local computer

2) Go to Task1_SQL folder there you will find "task1_SQL.py" python file

![fileSQL](https://github.com/SohamDoshi/OnlineSales.Ai_assignmnet/assets/106314995/e82bd6b2-cf0c-413d-a04b-01519d0b0a50)
