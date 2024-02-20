# go-mssql-insert-from-excel
Go script to read data from Excel file and insert into SQLExpress, Dream Report Manual Data table

Data in Excel follows this template:

|Tag|Date (text)|NumValue|
|---|---|---|
|Tag1|2024-02-19|11|
|Tag2|2024-02-20|12|

The script will prompt the user for a few details such as:
- user (sa) - make sure you enable the user in SSMS Security
- password
- hostname where SQLExpress is - make sure you enable TCP/IP for SQLExpress in SQL Server Configuration Manager
- database name
- table name - the script is made to insert data in `Manual_Data_md` tables created by Dream Report or AVEVA Reports for Operations
- file name

For close inspection, each sheet in inserted one at the time with user permission.


To build for Windows:
```
env GOOS=windows GOARCH=amd64 go build
```
