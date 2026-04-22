VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFC0C0&
   Caption         =   "General store management system"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17805
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   17805
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "customer form.frx":0000
      Height          =   2415
      Left            =   840
      TabIndex        =   23
      Top             =   6120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   4260
      _Version        =   393216
      Appearance      =   0
      DefColWidth     =   133
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   9720
      TabIndex        =   15
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print collective"
         Height          =   390
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print selective"
         Height          =   390
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   510
         Left            =   720
         TabIndex        =   19
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   510
         Left            =   720
         TabIndex        =   17
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Value"
         Height          =   375
         Left            =   1200
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Search by"
         Height          =   390
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      Height          =   390
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delete"
      Height          =   390
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Save"
      Height          =   390
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Clear"
      Height          =   390
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update"
      Height          =   390
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add New"
      Height          =   390
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3945
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   510
      Left            =   6600
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   510
      Left            =   6600
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   510
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5760
      Top             =   7080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=PRJ2531B/PRJ2531B;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=PRJ2531B/PRJ2531B;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM CUSTOMER"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Address"
      Height          =   390
      Left            =   5040
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Phone no"
      Height          =   390
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Name"
      Height          =   390
      Left            =   5040
      TabIndex        =   3
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "ID"
      Height          =   390
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "CUSTOMER DETAIL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   3885
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs


Private Sub Combo1_Click()
 Call CONN

If Combo1.Text = "ID" Then
    R.Open "select CUST_ID from customer", C
ElseIf Combo1.Text = "Name" Then
    R.Open "select CUST_NM from customer", C
ElseIf Combo1.Text = "Phone no" Then
    R.Open "select PHNO from customer", C
End If

Combo2.Clear

Do While Not R.EOF
    Combo2.AddItem R(0)
    R.MoveNext
Loop

R.Close

       
End Sub


Private Sub Command1_Click()
Dim CC As String
CC = "CUS00"
SQL = "SELECT COUNT(CUST_ID) FROM CUSTOMER"
Set R = C.Execute(SQL)
Label3.Caption = CC & R.Fields(0) + 1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
If Label3.Caption = "" Then
MsgBox "please select customer id to update", vbExclamation
Exit Sub
End If
SQL = "UPDATE CUSTOMER SET CUST_NM ='" + Text1.Text + "',PHNO=" + Text2.Text + ",ADDR='" + Text3.Text + "'WHERE CUST_ID='" + Combo2.Text + "'"
Set R = C.Execute(SQL)
MsgBox "RECORD UPDATED", vbInformation
Label3.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo2.SetFocus
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Label3.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text1.SetFocus
End Sub

Private Sub Command4_Click()
If Label3.Caption = "" Or Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "PLEASE ENTER ALL TEXT"
Else
SQL = "INSERT INTO CUSTOMER VALUES('" & Label3.Caption & "','" & Text2.Text & "','" & Text1.Text & "','" & Text3.Text & "')"

C.Execute SQL

MsgBox "RECORD SAVED"
Command4.Enabled = True
Adodc1.Refresh
End If
End Sub

Private Sub Command5_Click()
If Combo2.Text = "" Then
MsgBox "PLEASE SELECT CUSTOMER ID", vbExclamation
Exit Sub
End If
SQL = "DELETE FROM CUSTOMER WHERE CUST_ID='" + Combo2.Text + "'"
Set R = C.Execute(SQL)
MsgBox "RECORD DELETED", vbInformation
Label3.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo2.Text = ""
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Dim RESPONSE As VbMsgBoxResult
RESPONSE = MsgBox("ARE YOU SURE WANT TO EXIT?", vbYesNo + vbQuestion, "CONFIRM EXIT")
If RESPONSE = vbYes Then
Unload Me
End If
End Sub

Private Sub Command7_Click()
Call CONN
If Combo1.Text = "ID" Then
SQL = "select * from customer where CUST_ID like '%" & Combo2.Text & "%'"
ElseIf Combo1.Text = "Name" Then
SQL = "select * from customer where CUST_NM like '%" & Combo2.Text & "%'"
ElseIf Combo1.Text = "Phone no" Then
SQL = "select * from customer where PHNO like '%" & Combo2.Text & "%'"
End If
R.CursorLocation = adUseClient
R.Open SQL, C, adOpenStatic, adLockReadOnly
If R.EOF Then
MsgBox "This customer is not registered yet", vbInformation
Set DataGrid1.DataSource = Nothing
Exit Sub
End If
Set DataGrid1.DataSource = R
Label3.Caption = R("cust_id")
Text2.Text = R("cust_nm")
Text1.Text = R("phno")
Text3.Text = R("addr")
Adodc1.Refresh
End Sub

Private Sub Command8_Click()
If Combo1.Text = "ID" Then
Unload DataReport5
If DataEnvironment1.rsCommand5.State = 1 Then
DataEnvironment1.rsCommand5.Close
End If
DataEnvironment1.Commands("command5").CommandText = "select * from CUSTOMER where CUST_ID='" & Combo2.Text & "'"

DataEnvironment1.rsCommand5.Open
Set DataReport1.DataSource = DataEnvironment1
DataReport5.DataMember = "command5"
DataReport5.Show
Exit Sub
End If
If Combo1.Text = "Name" Then
Unload DataReport5
If DataEnvironment1.rsCommand5.State = 1 Then
DataEnvironment1.rsCommand5.Close
End If
DataEnvironment1.Commands("command5").CommandText = "select * from CUSTOMER where CUST_NM='" & Combo2.Text & "'"

DataEnvironment1.rsCommand5.Open
Set DataReport1.DataSource = DataEnvironment1
DataReport5.DataMember = "command5"
DataReport5.Show
Exit Sub
End If
If Combo1.Text = "Phone no" Then
Unload DataReport5
If DataEnvironment1.rsCommand5.State = 1 Then
DataEnvironment1.rsCommand5.Close
End If
DataEnvironment1.Commands("command5").CommandText = "select * from customer where Phno='" & Combo2.Text & "'"
DataEnvironment1.rsCommand5.Open
Set DataReport1.DataSource = DataEnvironment1
DataReport5.DataMember = "command5"
DataReport5.Show
Exit Sub
End If
End Sub

Private Sub Command9_Click()
DataReport6.Show
End Sub

Private Sub Form_Load()
CONN
Combo1.AddItem "ID"
Combo1.AddItem "Name"
Combo1.AddItem "Phone no"
Command4.Enabled = True
SQL = "SELECT CUST_ID FROM CUSTOMER"
Set R = C.Execute(SQL)
Do While Not R.EOF
'Combo1.AddItem R.Fields("CUST_ID")
R.MoveNext
Loop
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    ' Only numbers + backspace
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If

    ' Limit 10 digits
    If Len(Text1.Text) >= 10 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Or _
       KeyAscii = 32 Or _
       KeyAscii = 8 Then
        ' allow
    Else
        KeyAscii = 0
    End If
End Sub

