VERSION 5.00
Begin VB.Form frmAddress 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF00&
      Caption         =   "All"
      Height          =   255
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   25
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   24
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   23
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   22
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   21
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   20
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   19
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   18
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   17
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   16
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   15
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   14
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   13
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   12
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   11
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   10
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   9
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   8
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   7
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   6
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   5
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   4
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   3
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "A"
      Height          =   255
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Clear Record"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   9
      Left            =   4920
      TabIndex        =   23
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Delete Record"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Update Record"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Add New Record"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Height          =   4710
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   17
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   16
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   6
      Left            =   4920
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   13
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   12
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   10
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Pager:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   22
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Cell:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "FAX:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Phone:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "ZIP:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "State:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "City:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Address:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "First Name:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Last Name:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'    DO NOT RUN BEFORE READING PROGRAM NOTES!!!!!!!!!!!!!!
'
'    Program Notes:
'
'    This program requires that Active Data Objects (ADO) be selected from the
'    reference list for the project.  (Project->References->Microsoft Active Data Objects x.x, Library where x.x
'    referes to the version number.  This project was built with 2.7)
'
'    This project was built and tested with VB 6.0 (Service Pack 6) on Windows XP Professional.
'
'    The audience for this program is those that want to or are just beginning to use
'    ADO and MYSQL Server.  If you are an experience ADO/MYSQL programmer this will not
'    be you cup of tea.  The intent is to quickly get a beginner with little or no
'    database experience off and running, with the hope that they fully explore both
'    ADO and MYSQL.
'
'    Before running this program, please read the notes in the ADORoutines module as
'    if you don't have MYSQL installed this will fail.  Further documentation is
'    presented there as well.  There are also some global variable defined there that
'    are used in the form code.  You will have to change the login(LogName), password(Pword) and HostName
'    global variables to those set for your installed MYSQL database values.
'
'    If you are a beginner and the ADO Object set does not make sense to you I suggest
'    you invest the time and money to buy an ADO Book or take any one of the many Free
'    ADO tutorials on the net.  In any event, to really know how to use ADO you need to
'    understand how the objet model works.  An understanding of SQL syntax is also needed.
'
Option Explicit
Public ButtonIndex As String
Private Sub Form_Load()
Dim i As Integer
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.BackColor = vbRed
For i = 0 To 25
   Command5(i).Caption = Chr(65 + i)
Next i
Me.Show
DoEvents
On Error Resume Next
Set CN = New ADODB.Connection
CN.ConnectionString = "Provider=MSDASQL; DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & HostName & "; DATABASE=Address; UID=jack; PWD=gandalf; OPTION=3; PORT=3306"
CN.Open
If Err <> 0 Then
  Set CN = Nothing
  MakeDatabase
  Set CN = New ADODB.Connection
  CN.ConnectionString = "Provider=MSDASQL; DRIVER={MySQL ODBC 3.51 Driver}; SERVER=localhost; DATABASE=Address; UID=jack; PWD=gandalf; OPTION=3; PORT=3306"
  CN.Open
End If
GetNames
End Sub

Private Sub Command1_Click()
Dim Lname As String
Dim Fname As String
Dim i As Integer
Lname = Text1(0).Text
Fname = Text1(1).Text
If ValidateRecords() Then
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.LockType = adLockPessimistic
    RS.Open "SELECT * FROM addressbook WHERE LastName='" & Lname & " ' AND FirstName='" & Fname & "'", CN
    If RS.BOF Then
        RS.AddNew
        RS!LastName = Text1(0).Text
        RS!FirstName = Text1(1).Text
        RS!Address = Text1(2).Text
        RS!City = Text1(3).Text
        RS!State = Text1(4).Text
        RS!ZIP = Text1(5).Text
        RS!Phone = Text1(6).Text
        RS!FAX = Text1(7).Text
        RS!Cell = Text1(8).Text
        RS!Pager = Text1(9).Text
        RS.Update
        RS.Close
        If Command6.BackColor = vbRed Or ButtonIndex = UCase(Left(Text1(0).Text, 1)) Then
           List1.AddItem Text1(0).Text & ", " & Text1(1).Text
        End If
        Command4_Click
   Else
        MsgBox "Duplicate First and Last Name! Record Not Added."
        RS.Close
   End If
Else
    MsgBox "You must provide a Last Name, First Name and Phone Number!"
End If
End Sub
Private Function ValidateRecords() As Boolean
Dim i As Integer
If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(6).Text = "" Then
   ValidateRecords = False
   Exit Function
End If
For i = 2 To 5
   If Text1(i).Text = "" Then
      Text1(i).Text = "None"
   End If
Next i
For i = 7 To 9
   If Text1(i).Text = "" Then
      Text1(i).Text = "None"
   End If
Next i
ValidateRecords = True
End Function

Private Sub Command2_Click()
Dim Lname As String
Dim Fname As String
Dim i As Integer
Lname = Text1(0).Text
Fname = Text1(1).Text
If ValidateRecords() Then
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.LockType = adLockPessimistic
    RS.Open "SELECT * FROM addressbook WHERE LastName='" & Lname & " ' AND FirstName='" & Fname & "'", CN
    If RS.BOF Then
        MsgBox "Unable to find record in database to update!"
    Else
        RS!LastName = Text1(0).Text
        RS!FirstName = Text1(1).Text
        RS!Address = Text1(2).Text
        RS!City = Text1(3).Text
        RS!State = Text1(4).Text
        RS!ZIP = Text1(5).Text
        RS!Phone = Text1(6).Text
        RS!FAX = Text1(7).Text
        RS!Cell = Text1(8).Text
        RS!Pager = Text1(9).Text
        RS.Update
        Command4_Click
        RS.Close
    End If
Else
   MsgBox "You must enter Last Name, First Name and a phone number at a minimum!"
End If
       
End Sub

Private Sub Command3_Click()
Dim Lname As String
Dim Fname As String
Dim i As Integer
Lname = Text1(0).Text
Fname = Text1(1).Text
If ValidateRecords() Then
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenForwardOnly
    RS.LockType = adLockPessimistic
    RS.Open "SELECT * FROM addressbook WHERE LastName='" & Lname & " ' AND FirstName='" & Fname & "'", CN
    If RS.BOF Then
       MsgBox "Unable to find record to delete in the database!"
    Else
       For i = 0 To List1.ListCount - 1
          If List1.List(i) = Lname & ", " & Fname Then
             List1.RemoveItem i
             Exit For
          End If
       Next i
       RS.Delete
       Command4_Click
    End If
End If
End Sub

Private Sub Command4_Click()
Dim i As Integer
For i = 0 To 9
   Text1(i).Text = ""
Next i
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Text1(0).SetFocus
End Sub

Private Sub Command5_Click(Index As Integer)
Dim i As Integer
For i = 0 To 25
   Command5(i).BackColor = &HFFFF00
Next i
Command6.BackColor = &HFFFF00
Command5(Index).BackColor = vbRed
ButtonIndex = Command5(Index).Caption
List1.Clear
Command4_Click
Set RS = New ADODB.Recordset
RS.CursorType = adOpenForwardOnly
RS.LockType = adLockPessimistic
RS.Open "SELECT * FROM addressbook WHERE LastName LIKE '" & Command5(Index).Caption & "%'", CN
If Not RS.BOF Then
   While Not RS.EOF
      List1.AddItem RS!LastName & ", " & RS!FirstName
      RS.MoveNext
   Wend
End If
RS.Close
End Sub

Private Sub Command6_Click()
Dim i As Integer
For i = 0 To 25
   Command5(i).BackColor = &HFFFF00
Next i
Command4_Click
Command6.BackColor = vbRed
ButtonIndex = "*"
GetNames
End Sub

Private Sub MakeDatabase()
   If Not AdoCreateDatabase("Address") Then
      MsgBox "Unable to create new database!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateTable("Address", "AddressBook") Then
      MsgBox "Can not Create Address Table!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "LastName varchar(35)") Then
      MsgBox "Can not create field LastName!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "FirstName varchar(35)") Then
      MsgBox "Can not create field FirstName!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "Address varchar(45)") Then
      MsgBox "Can not create field Address!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "City varchar(25)") Then
      MsgBox "Can not create field LastName!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "State varchar(25)") Then
      MsgBox "Can not create field LastName!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "ZIP varchar(35)") Then
      MsgBox "Can not create field Zip!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "Phone varchar(15)") Then
      MsgBox "Can not create field LastName!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "FAX varchar(15)") Then
      MsgBox "Can not create field FAX!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "Cell varchar(15)") Then
      MsgBox "Can not create field Cell!"
      Set CN = Nothing
      End
   End If
   If Not AdoCreateField("Address", "AddressBook", "Pager varchar(15)") Then
      MsgBox "Can not create field Pager!"
      Set CN = Nothing
      End
   End If
End Sub
Private Sub GetNames()
List1.Clear
Set RS = New ADODB.Recordset
RS.CursorType = adOpenForwardOnly
RS.LockType = adLockPessimistic
RS.Open "SELECT * FROM addressbook ORDER BY LastName", CN
If Not RS.BOF Then
   While Not RS.EOF
      List1.AddItem RS!LastName & ", " & RS!FirstName
      RS.MoveNext
   Wend
End If
RS.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
RS.Close
CN.Close
Set RS = Nothing
Set CN = Nothing
End
End Sub

Private Sub List1_DblClick()
Dim Lname As String
Dim Fname As String
Dim i As Integer
i = InStr(1, List1.List(List1.ListIndex), ", ", vbTextCompare)
If i = 0 Then
   Exit Sub
End If
Lname = Left(List1.List(List1.ListIndex), i - 1)
Fname = Mid(List1.List(List1.ListIndex), i + 2)
Set RS = New ADODB.Recordset
RS.CursorType = adOpenForwardOnly
RS.LockType = adLockPessimistic
RS.Open "SELECT * FROM addressbook WHERE LastName='" & Lname & "' AND FirstName='" & Fname & "'", CN
If Not RS.BOF Then
   Text1(0).Text = RS!LastName
   Text1(1).Text = RS!FirstName
   Text1(2).Text = RS!Address
   Text1(3).Text = RS!City
   Text1(4).Text = RS!State
   Text1(5).Text = RS!ZIP
   Text1(6).Text = RS!Phone
   Text1(7).Text = RS!FAX
   Text1(8).Text = RS!Cell
   Text1(9).Text = RS!Pager
End If
RS.Close
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 11 Then
   KeyAscii = 0
   If Index = 9 Then
      Command1.SetFocus
   Else
      Text1(Index + 1).SetFocus
   End If
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Index = 0 Then
   Command1.Enabled = True
   Command2.Enabled = True
   Command3.Enabled = True
   Command4.Enabled = True
End If
End Sub
