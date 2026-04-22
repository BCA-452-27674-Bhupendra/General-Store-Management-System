VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "General Store Management System"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15225
   Begin VB.ComboBox combo2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3120
      TabIndex        =   30
      Text            =   "combo2"
      Top             =   2760
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "EMPDET.frx":0000
      Height          =   2655
      Left            =   240
      TabIndex        =   25
      Top             =   6120
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5880
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      RecordSource    =   "SELECT * FROM EMPLOYEE"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add new"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   11400
      TabIndex        =   20
      Top             =   600
      Width           =   3255
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print collective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print selective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   32
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select value"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   21
         Top             =   1200
         Width           =   1320
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   13
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   12
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   11
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   2415
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   11520
      TabIndex        =   28
      Top             =   1080
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16761024
      Appearance      =   1
      StartOfWeek     =   126287873
      CurrentDate     =   46030
   End
   Begin VB.Label Label12 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   24
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6000
      TabIndex        =   8
      Top             =   3600
      Width           =   1755
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experience"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6000
      TabIndex        =   7
      Top             =   2760
      Width           =   1470
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6000
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hire Date"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6000
      TabIndex        =   5
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Role"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   2010
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE DETAIL"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   3765
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo3_Click()
    If Combo3.Text = "" Then Exit Sub
Combo1.Enabled = True

    Call CONN
    Combo1.Clear
    Set R = New ADODB.Recordset

    If Combo3.Text = "ID" Then
        R.Open "select emp_id from employee", C
        While Not R.EOF
        Combo1.AddItem R.Fields(0).Value
        R.MoveNext
        Wend
        R.Close
        Exit Sub
        End If
        
        

    If Combo3.Text = "Name" Then
    Text1.Visible = True
    Exit Sub
    End If
    
       ' R.Open "select emp_nm from employee", C
 If Combo3.Text = "Phno" Then

        R.Open "select phno from employee", C
        While Not R.EOF
        Combo1.AddItem R.Fields(0).Value
        R.MoveNext
        Wend
        R.Close
        Exit Sub
        End If
        

    If Combo3.Text = "Role" Then
    Combo1.Clear
    Combo1.AddItem "Manager"
    Combo1.AddItem "Salesman"
    Combo1.AddItem "Delivery boy"
    Combo1.AddItem "Purchase assistant"
    Combo1.AddItem "Cashier"
    Combo1.AddItem "Store Keeper"
    Exit Sub
    End If
    End Sub
Private Sub Command7_Click()
If Combo3.Text = "Role" Then
Unload DataReport7
If DataEnvironment1.rsCommand7.State = 1 Then
DataEnvironment1.rsCommand7.Close
End If
DataEnvironment1.Commands("command7").CommandText = "select * from employee where role='" & Combo1.Text & "'"
DataEnvironment1.rsCommand7.Open
Set DataReport7.DataSource = DataEnvironment1
DataReport7.DataMember = "command7"
DataReport7.Show
Exit Sub
End If
If Combo3.Text = "Phno" Then
Unload DataReport7
If DataEnvironment1.rsCommand7.State = 1 Then
DataEnvironment1.rsCommand7.Close
End If
DataEnvironment1.Commands("command7").CommandText = "select * from employee where Phno='" & Combo1.Text & "'"
DataEnvironment1.rsCommand7.Open
Set DataReport7.DataSource = DataEnvironment1
DataReport7.DataMember = "command7"
DataReport7.Show
Exit Sub
End If
If Combo3.Text = "ID" Then
If Combo1.Text = "" Then
MsgBox "please select id", vbExclamation
Exit Sub
End If
If DataEnvironment1.rsCommand7.State = adStateOpen Then
DataEnvironment1.rsCommand7.Close
End If
DataEnvironment1.Commands("command7").CommandText = "select * from employee where emp_id='" & Combo1.Text & "'"
DataEnvironment1.rsCommand7.Open
DataReport7.Show
Exit Sub
End If


If Combo3.Text = "Name" Then
If Text1.Text = "" Then
MsgBox "please enter Name", vbExclamation
Exit Sub
End If
If DataEnvironment1.rsCommand7.State = adStateOpen Then
DataEnvironment1.rsCommand7.Close
End If
DataEnvironment1.Commands("command7").CommandText = "select * from employee where UPPER(emp_nm) like '%" & UCase(Text1.Text) & "%'"

DataEnvironment1.rsCommand7.Open
DataReport7.Show
Exit Sub
End If
End Sub

Private Sub Command8_Click()
DataReport8.Show
End Sub

Private Sub Command9_Click()
'Text1.Visible = True
Call CONN
Set R = New ADODB.Recordset

If Combo3.Text = "ID" Then
    SQL = "SELECT * FROM EMPLOYEE WHERE EMP_ID='" & Combo1.Text & "'"
ElseIf Combo3.Text = "Name" Then
    SQL = "SELECT * FROM EMPLOYEE WHERE UPPER(EMP_NM) LIKE '%" & UCase(Text1.Text) & "%'"
    
    ElseIf Combo3.Text = "Role" Then
    SQL = "SELECT * FROM EMPLOYEE WHERE ROLE LIKE '%" & Combo1.Text & "%'"
ElseIf Combo3.Text = "Phno" Then
    SQL = "SELECT * FROM EMPLOYEE WHERE PHNO LIKE '%" & Combo1.Text & "%'"
End If

R.CursorLocation = adUseClient
R.Open SQL, C, adOpenStatic, adLockReadOnly

If R.EOF Then
    MsgBox "This employee is not registered yet", vbInformation
    Set DataGrid1.DataSource = Nothing
    Exit Sub
End If

Set DataGrid1.DataSource = R

Label12.Caption = R.Fields(0)
Text2.Text = R.Fields(1)
Combo2.Text = R.Fields(2)
Text4.Text = R.Fields(3)
Text5.Text = R.Fields(4)
Text6.Text = R.Fields(5)
Text7.Text = R.Fields(6)
Text8.Text = R.Fields(7)
Adodc1.Refresh
End Sub



Private Sub Form_Load()
Text1.Visible = False
Combo3.Clear
Combo3.AddItem "ID"
Combo3.AddItem "Name"
Combo3.AddItem "Role"
Combo3.AddItem "Phno"
MonthView1.MaxDate = Date
Combo2.AddItem "Manager"
Combo2.AddItem "Cashier"
Combo2.AddItem "Salesman"
Combo2.AddItem "Store keeper"
Combo2.AddItem "Purchase assistant"
Combo2.AddItem "Delivery boy"
Combo2.AddItem "Accountant"
Command3.Enabled = True
Command2.Enabled = True
Combo1.Enabled = False


    MonthView1.Visible = False
    CONN
    Adodc1.RecordSource = "SELECT * FROM EMPLOYEE ORDER BY EMP_ID ASC"
    Adodc1.Refresh
    SQL = "select EMP_ID FROM EMPLOYEE"
    Set R = C.Execute(SQL)
    Do While Not R.EOF
    Combo1.AddItem R.Fields("EMP_ID")
    R.MoveNext
    Loop
End Sub

Private Sub Command6_Click()
    Dim CC As String
    CONN
    CC = "EMP"
    SQL = "SELECT COUNT(EMP_ID) FROM EMPLOYEE"
    Set R = C.Execute(SQL)
    Text2.Text = ""
    Combo2.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Label12.Caption = CC & R.Fields(0) + 1
    
    
    
    Text2.SetFocus
End Sub

Private Sub Command1_Click()
               On Error GoTo abc
        If Text2.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text8.Text = "" Then
    MsgBox "please fill all required fields", vbExclamation
    Exit Sub
    End If
    If MonthView1.Value > Date Then
    MsgBox "future date is not allowed. please select current date.", vbExclamation
    MonthView1.SetFocus
    Exit Sub
    End If
    SQL = "INSERT INTO EMPLOYEE VALUES('" + Label12.Caption + "','" + Text2.Text + "','" + Combo2.Text + "'," + Text4.Text + ",'" + Format(MonthView1.Value, "dd/mmm/yyyy") + "','" + Text6.Text + "','" + Text7.Text + "','" + Text8.Text + "')"
    Set R = C.Execute(SQL)
    MsgBox "record saved successfully", vbInformation
    Label12.Caption = ""
    Text2.Text = ""
    Combo2.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text2.SetFocus
    Adodc1.Refresh
    Exit Sub
abc:
    MsgBox "ENTER 5000 TO 50000"
    Text6.Text = ""
    Text6.SetFocus
    End Sub

Private Sub Command2_Click()
    If Label12.Caption = "" Then
        MsgBox "Please select employee ID to update", vbExclamation
        Exit Sub
    End If
    CONN
    SQL = "UPDATE EMPLOYEE SET EMP_NM='" + Text2.Text + "',ROLE='" + Combo1.Text + "',PHNO=" + Text4.Text + ",H_DT='" + Format(MonthView1.Value, "dd/mmm/yyyy") + "',SAL=" + Text6.Text + ",EXP='" + Text7.Text + "',QAL='" + Text8.Text + "' WHERE EMP_ID='" + Combo1.Text + "'"
    
     
   Set R = C.Execute(SQL)
   
   

    MsgBox "Record updated successfully", vbInformation
     Label12.Caption = ""
    Text2.Text = ""
    Combo2.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text2.SetFocus
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
    If Label12.Caption = "" Then
        MsgBox "No Record Selected to Delete", vbExclamation
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    CONN
    C.Execute "DELETE FROM EMPLOYEE WHERE EMP_ID = '" & Label12.Caption & "'"
    MsgBox "Record deleted successfully", vbInformation
    Text2.Text = ""
    Combo2.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Label12.Caption = ""
    Text2.SetFocus
End Sub

Private Sub Command4_Click()
    Text2.Text = ""
    Combo2.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text2.SetFocus
    Call loademployee
    
    End Sub
Private Sub loademployee()
Set R = New ADODB.Recordset
R.CursorLocation = adUseClient
R.Open "select * from employee", C, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = R
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text5.Text = Format(MonthView1.Value, "dd/mmm/yyyy")
MonthView1.Visible = False
MonthView1.ZOrder 0
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
If KeyAscii <> 8 Then
KeyAscii = 0
Exit Sub
End If
End If
If Len(Text4.Text) >= 10 And KeyAscii <> 8 Then
KeyAscii = 0
End If
End Sub
Private Sub Text4_LostFocus()
If Len(Text4.Text) <> 10 Then
MsgBox "Phone Enter all required fields", vbExclamation
Text4.SetFocus
End If
End Sub
Private Sub Text5_GotFocus()
MonthView1.Visible = True
MonthView1.Left = Text5.Left + Text5.Width + 100
MonthView1.Top = Text5.Top
MonthView1.ZOrder 0
End Sub

