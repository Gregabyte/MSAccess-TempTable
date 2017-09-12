MSAccess-TempTable
========================

About
-----

Custom functions to create a temporary file to store temporary Access Tables.  This allows your frontend to use temporary tables without bloating the frontend's file size.

Installing the Custom Scripts
-----------------------------

For the purposes of these instructions, assume your database is called `C:\Temp\myDb.accdb`.

1. Load `mod_Temp_Tables.bas` into a new module in your database with that exact name.
 1. Go to the VBA editor (CTRL-G) and select "File" > "Import File..."
    (or you can just drag and drop the file from windows explorer into the vba editor module list).
 2. Select the `mod_Temp_Tables.bas` file.
 3. Save the file (CTRL-S).

Usage
-----
bVariable = UpdateTempTable(TableName, TableQueryOrSQL, [InCurrentDB], [ValidMinutes], [PK Field])

TableName: Required, string. This is the name of the table that will be created by the function.

TableQueryOrSQL: This argument will accept the name of a table, query, or a SQL SELECT query string (one that contains the term “SELECT” as the first word of the query and DOES NOT include the term “INTO”).  You cannot use a make-table or an append query as the source for your temp table.

InCurrentDB: Optional, boolean, default = False.  The value of this argument determines whether the table gets created in the front-end (True) or in an external database (False).  If False, the external database will be created on the same path as the front-end and have the same file name as the front-end but have the additional “_temp” at the end of the file name.  
Example front-end filename: C:\Temp\myDb.accdb
Example temp table filename: C:\Temp\myDb_Temp.accdb

ValidMinutes: Optional, integer, default = 0.  Determines how many minutes the temporary table is valid for.  A value of zero indicates that the table should overwritten if it already exists.  Any other value will cause the function to check the DateCreated property of the table if it already exists and if the number of minutes since the table was created is less than the ValidMinutes argument, the function exits without updating the table.

PK Field: Optional, Default = NULL.  Will accept a single field and will use an “ALTER COLUMN” DDL query to set this field as the Primary key for the temp table.

Contributing
============

Original functions were created by Dale Fye.  You can find his article at https://www.experts-exchange.com/articles/9753/Creating-and-using-Temporary-Tables-in-Microsoft-Access.html.
Pull requests, issue reports etc welcomed.