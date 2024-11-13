# Emp Avg Sal by Dept
# Employee Count Bye Dept
# Sum of Sal by Dept
# Emp Count by Zip Code

import pyodbc as odbc
import csv

conn_str = (
        r'Driver=SQL Server;' +
        r'Server=WIN-N7UGT6C69B8;' +
        r'Database=NYCTaxi_Sample;' +
        r'Trusted_Connection=yes;'
)

def Excercise1(conn_str, sqlcommandrun, output_filename):
    conn = odbc.connect(conn_str)

    with conn.cursor() as zzz:
        zzz.execute(sqlcommandrun)

    # Fetch all rows
        columns = [desc[0] for desc in zzz.description]
        rows = zzz.fetchall()
        for row in rows:
            print(row)

    with open(output_filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(columns)
        for row in rows:
            writer.writerow(row)

    print(f"Data saved to {output_filename}")
    conn.close()


query1 = """
SELECT d.Dept_Name, AVG(e.Emp_Sal) AS AvgSalary, SUM(e.Emp_Sal) AS SumSalary 
FROM EMPLOYEE e 
JOIN EMP_DEPARTMENT d ON e.EDept_Id = d.Dept_Id 
GROUP BY d.Dept_Name
"""

query2 = """
SELECT d.Dept_Name, e.EDept_Id, COUNT(*) AS EmployeeCount  
FROM EMPLOYEE e 
JOIN EMP_DEPARTMENT d ON e.EDept_Id = d.Dept_Id 
GROUP BY d.Dept_Name, e.EDept_Id 
ORDER BY EmployeeCount DESC
"""

query3 = """
SELECT d.Dept_Name, e.Emp_Zip, COUNT(*) AS EmployeeCountByZip 
FROM EMPLOYEE e 
JOIN EMP_DEPARTMENT d ON e.EDept_Id = d.Dept_Id 
GROUP BY d.Dept_name, e.Emp_Zip  
ORDER BY EmployeeCountByZip DESC
"""

Excercise1(conn_str, query1, "AvgSalary_SumSalary_ByDept.csv")
Excercise1(conn_str, query2, "EmployeeCount_ByDept.csv")
Excercise1(conn_str, query3, "EmployeeCount_ByZip.csv")


