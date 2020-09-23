Attribute VB_Name = "ADORoutines"
'
'   This module is for general purpose creation of a MYSQL database.
'
'   Items of note:
'
'       These routines were developed with the intent of helping someone
'       who has little or no experience with ADO as well as database
'       programming.  Hopefully it will provide enough information to
'       get them started down the road to fully understanding how to
'       use ADO and MYSQL.  I fully encourage anyone who uses this
'       to investigate further and delve into the MYSQL and ADO documentation
'       as what is presented here only scratches the surface of what
'       you can accomplish.
'
'       Development and testing was done using MS XP Professional.  It has
'       not been tested on any other operating system.
'
'       Do not use a column/field name of "dummy".  This name is used
'       to create a temporary column/field name when creating the table.
'
'       The logic in these routines dictate that when you create a table
'       the creation of all columns/fields will follow before creating
'       a new table.
'
'       Errors are not handled in the routine(s).  A true value will be passed
'       back from the function(s) on sucess and false if failed.  It is up to
'       the calling routine to dictate how to handle the error.
'
'       These routines will only work with MYSQL database.  If you need the
'       this capability for MS SQL Server contact me and I will send them to
'       you.
'
'       You will need to have MYSQL installed on your local computer.  The username must have the
'       correct priveledges to connect to the Variable Defined MasterDb and also have the
'       privledges to create a new database.
'
'       If you do not have MYSQL database you can obtain it for free at www.mysql.com.
'       MYSQL is an open source project (thank you much guys and gals) and if you can
'       donate a little cash or time I am sure they could more than use it.  You will need to
'       download the following items from the MYSQL web site www.mysql.com:
'
'       Latest version of MYSQL Database
'       Latest version of MYSQL Driver
'       Not required but sure makes life easier if you download MYSQL Control Center
'
'       Please note that the "Provider" Driver used within this project is 3.51.  If you use
'       a different version you must change the CN.ConnectionString to reflect the new driver.
'
'       <Jack Rizzo>jrizzo@allianceatmservices.com
'
Option Explicit
Global Const HostName = "localhost"        'server name to connect to
Global Const MasterDb = "MySql"            'master database to connect to on database create
Global Const LogName = "yourloginname"     'the MYSQL server login name YOU WILL NEED TO CHANGE
Global Const Pword = "yourpassword"        'the MYSQL login server password YOU WILL NEED TO CHANGE
Global CN As ADODB.Connection              'adodb connection variable
Global RS As ADODB.Recordset               'adodb recordset variable
Global TripVar As Boolean                  'switch used to delete dummy column/field

Public Function AdoCreateDatabase(dBname As String) As Boolean
'
'   Create the argument database
'
'
AdoCreateDatabase = True
On Error GoTo CreateError
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider=MSDASQL; DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & HostName & "; DATABASE=" & MasterDb & _
   "; UID=" & LogName & "; PWD=" & Pword & "; OPTION=3"
CN.Open
CN.Execute "CREATE DATABASE " & dBname
CN.Close
Set CN = Nothing
Exit Function
CreateError:
AdoCreateDatabase = False
Set CN = Nothing
End Function
Public Function AdoCreateTable(dBname As String, TableName As String) As Boolean
'
'  Create a table in the database
'
'  ado CREATE TABLE requires at least one column/field be specified on create
'  we make a dummy column/field value that will be deleted before exit
'
Dim SqlStmt As String
AdoCreateTable = True
On Error GoTo TableError
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider=MSDASQL; DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & HostName & "; DATABASE=" & dBname & _
   "; UID=" & LogName & "; PWD=" & Pword & "; OPTION=3"
CN.Open
SqlStmt = "CREATE TABLE " & TableName & "(dummy varchar(10))"
CN.Execute SqlStmt
Set CN = Nothing
TripVar = False
Exit Function
TableError:
AdoCreateTable = False
Set CN = Nothing
End Function
Public Function AdoCreateField(dBname As String, TableName As String, FieldVar As String) As Boolean
'
'   Create a column/field in the table
'
'   Note that FieldVar argument is passed in with appropriate ado syntax, i.e.:
'
'   LastName varchar(35)  would create a variable character field max of 35 characters
'
Dim SqlStmt As String
AdoCreateField = True
On Error GoTo FieldError
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider=MSDASQL; DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & HostName & "; DATABASE=" & dBname & _
   "; UID=" & LogName & "; PWD=" & Pword & "; OPTION=3"
CN.Open
CN.Execute "ALTER TABLE " & TableName & " ADD " & FieldVar
'
'  if this is the first column/field created then delete the dummy
'  created when the table was made.
'
If Not TripVar Then
   CN.Execute "ALTER TABLE " & TableName & " DROP dummy"
   TripVar = True
End If
CN.Close
Set CN = Nothing
Exit Function
FieldError:
AdoCreateField = False
Set CN = Nothing
End Function
