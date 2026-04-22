VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   " "
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   14805
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "supplierDETAIL.frx":0000
      Height          =   2295
      Left            =   360
      TabIndex        =   39
      Top             =   6720
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   4048
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
      Left            =   4320
      Top             =   7560
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   "Provider=MSDAORA.1;User ID=PRJ2531B/PRJ2531B;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=PRJ2531B/PRJ2531B;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM SUP_DET"
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
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   38
      Text            =   "Text4"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   36
      Text            =   "Text12"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   34
      Text            =   "Text11"
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   32
      Text            =   "Text6"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   30
      Text            =   "Text10"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   28
      Text            =   "Text9"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   26
      Text            =   "Text8"
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   24
      Text            =   "Text7"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   22
      Text            =   "Text5"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   21
      Text            =   "Text3"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   9960
      TabIndex        =   9
      Top             =   1320
      Width           =   4695
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print collective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print selective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2520
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
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
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Search by"
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
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command6 
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
      Height          =   330
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
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
      Height          =   330
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
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
      Height          =   330
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Height          =   330
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
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
      Height          =   330
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   330
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact no"
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
      Left            =   240
      TabIndex        =   37
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IFSC"
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
      Left            =   5280
      TabIndex        =   35
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C no."
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
      Left            =   5280
      TabIndex        =   33
      Top             =   3240
      Width           =   1020
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ac.Holder name"
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
      Left            =   5280
      TabIndex        =   31
      Top             =   2520
      Width           =   2190
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branch name"
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
      Left            =   5280
      TabIndex        =   29
      Top             =   1680
      Width           =   1770
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank name"
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
      Left            =   5280
      TabIndex        =   27
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   240
      TabIndex        =   25
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gst Number"
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
      Left            =   240
      TabIndex        =   23
      Top             =   4560
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Supplier Phno"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   2025
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person "
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
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   2100
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company name"
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
      Left            =   240
      TabIndex        =   17
      Top             =   3840
      Width           =   2145
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Supplier ID"
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "SUPPLIER DETAIL"
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
      Left            =   6630
      TabIndex        =   0
      Top             =   120
      Width           =   3510
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public C As ADODB.Connection
Public R As ADODB.Recordset
Public SQL As String
Private mMode As String   ' STARTUP / ADD / EDIT

'========================================================
' CONNECTION
'========================================================
Public Sub CONN()
    On Error GoTo ErrHandler

    If C Is Nothing Then
        Set C = New ADODB.Connection
        C.Open "Provider=MSDAORA.1;User ID=PRJ2531B;Password=PRJ2531B;Persist Security Info=False"
    ElseIf C.State = 0 Then
        C.Open "Provider=MSDAORA.1;User ID=PRJ2531B;Password=PRJ2531B;Persist Security Info=False"
    End If
    Exit Sub

ErrHandler:
    MsgBox "Database connection error: " & Err.Description, vbCritical, "Connection Error"
End Sub

Private Sub Command8_Click()
    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Select value", vbExclamation
        Exit Sub
    End If

    If Combo1.Text = "ID" Then
        DataEnvironment1.Commands("Command9").CommandText = _
        "SELECT * FROM SUP_DET WHERE SUP_ID='" & Combo2.Text & "'"

    ElseIf Combo1.Text = "Name" Then
        DataEnvironment1.Commands("Command9").CommandText = _
        "SELECT * FROM SUP_DET WHERE SUP_NM='" & Combo2.Text & "'"
    End If

    If DataEnvironment1.rsCommand9.State = 1 Then
        DataEnvironment1.rsCommand9.Close
    End If

    DataEnvironment1.Commands("COMMAND9").Execute


    DataReport2.Show

End Sub

Private Sub Command9_Click()
DataReport4.Show
End Sub

'========================================================
' FORM LOAD
'========================================================
Private Sub Form_Load()
    On Error GoTo ErrHandler

    Call CONN

    Combo1.Clear
    Combo1.AddItem "ID"
    Combo2.Clear
    Combo2.Enabled = False

    Call LoadData
    Call ClearData
    Call SetStartupMode

    Exit Sub

ErrHandler:
    MsgBox "Form load error: " & Err.Description, vbCritical, "System Error"
End Sub

'========================================================
' MODE HANDLING
'========================================================
Private Sub SetStartupMode()
    mMode = "STARTUP"

    Call EnableEntryControls(False)
    Call EnableSearchControls(True)

    Command1.Enabled = True   ' Add New
    Command2.Enabled = False  ' Update
    Command3.Enabled = True   ' Clear
    Command4.Enabled = False  ' Save
    Command5.Enabled = False  ' Delete
    Command6.Enabled = True   ' Exit
    Command7.Enabled = False  ' Search
End Sub

Private Sub SetAddMode()
    mMode = "ADD"

    Call EnableEntryControls(True)
    Call EnableSearchControls(False)

    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = True
    Command7.Enabled = False

End Sub

Private Sub SetEditMode()
    mMode = "EDIT"

    Call EnableEntryControls(True)
    Call EnableSearchControls(True)

    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = False
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True

End Sub

Private Sub EnableEntryControls(ByVal bFlag As Boolean)
    Text2.Enabled = bFlag   ' Company Name
    Text3.Enabled = bFlag   ' Supplier Phone
    Text4.Enabled = bFlag   ' Account No
    Text5.Enabled = bFlag   ' Contact Person
    Text6.Enabled = bFlag   ' Account Holder
    Text7.Enabled = bFlag   ' GST Number
    Text8.Enabled = bFlag   ' Address
    Text9.Enabled = bFlag   ' Branch Name
    Text10.Enabled = bFlag  ' Bank Name
    Text11.Enabled = bFlag  ' Contact No
    Text12.Enabled = bFlag  ' IFSC
End Sub

Private Sub EnableSearchControls(ByVal bFlag As Boolean)
    Combo1.Enabled = bFlag
    Combo2.Enabled = bFlag
End Sub

'========================================================
' LOAD GRID
'========================================================
Private Sub LoadData()
    On Error GoTo ErrHandler

    If C Is Nothing Then Call CONN
    If C.State = 0 Then Call CONN

    If Not R Is Nothing Then
        If R.State = 1 Then R.Close
    End If

    Set R = New ADODB.Recordset
    R.CursorLocation = adUseClient
    R.Open "SELECT SUP_ID, SUP_PH, CON_PER, CON_NO, CMP_NM, ADDR, BNK_NM, BR_NM FROM SUP_DET ORDER BY SUP_ID ASC", _
           C, adOpenStatic, adLockReadOnly

    Set DataGrid1.DataSource = R
    Exit Sub

ErrHandler:
    MsgBox "Grid load error: " & Err.Description, vbCritical, "System Error"
End Sub

'========================================================
' HELPERS
'========================================================
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Trim(s), "'", "''")
End Function

Private Function NzField(ByVal v As Variant) As String
    If IsNull(v) Then
        NzField = ""
    Else
        NzField = Trim(CStr(v))
    End If
End Function

Private Function IsDigitsOnly(ByVal s As String) As Boolean
    Dim i As Integer
    s = Trim(s)

    If s = "" Then
        IsDigitsOnly = False
        Exit Function
    End If

    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next i

    IsDigitsOnly = True
End Function

Private Function ProperCaseText(ByVal s As String) As String
    s = Trim(s)

    If s = "" Then
        ProperCaseText = ""
    Else
        ProperCaseText = StrConv(s, vbProperCase)
    End If
End Function

Private Function IsValidIFSC(ByVal s As String) As Boolean
    s = Trim(s)

    If s = "" Then
        IsValidIFSC = True
    ElseIf Len(s) <> 11 Then
        IsValidIFSC = False
    Else
        IsValidIFSC = True
    End If
End Function

Private Function IsValidGST(ByVal s As String) As Boolean
    s = Trim(s)

    If s = "" Then
        IsValidGST = True
    ElseIf Len(s) <> 15 Then
        IsValidGST = False
    Else
        IsValidGST = True
    End If
End Function

Private Function RecordExists(ByVal pSupID As String) As Boolean
    Dim rsChk As New ADODB.Recordset

    On Error GoTo ErrHandler
    RecordExists = False

    rsChk.Open "SELECT SUP_ID FROM SUP_DET WHERE SUP_ID='" & Esc(pSupID) & "'", C, adOpenStatic, adLockReadOnly
    If Not rsChk.EOF Then RecordExists = True
    rsChk.Close
    Set rsChk = Nothing
    Exit Function

ErrHandler:
    RecordExists = False
End Function

Private Function ValidateData() As Boolean
    ValidateData = False

    If Trim(Label3.Caption) = "" Then
        MsgBox "Supplier ID is missing.", vbExclamation, "Validation Error"
        Exit Function
    End If

    
   

    If Trim(Text3.Text) = "" Then
        MsgBox "Please enter Supplier Phone Number.", vbExclamation, "Validation Error"
        Text3.SetFocus
        Exit Function
    End If

    If Not IsDigitsOnly(Text3.Text) Or Len(Trim(Text3.Text)) <> 10 Then
        MsgBox "Supplier Phone Number must be exactly 10 digits.", vbExclamation, "Validation Error"
        Text3.SetFocus
        Exit Function
    End If

    If Trim(Text11.Text) <> "" Then
        If Not IsDigitsOnly(Text11.Text) Or Len(Trim(Text11.Text)) <> 10 Then
            MsgBox "Contact Number must be exactly 10 digits.", vbExclamation, "Validation Error"
            Text11.SetFocus
            Exit Function
        End If
    End If

    If Trim(Text2.Text) <> "" And Len(Trim(Text2.Text)) > 40 Then
        MsgBox "Company Name cannot exceed 40 characters.", vbExclamation, "Validation Error"
        Text2.SetFocus
        Exit Function
    End If

    If Trim(Text5.Text) <> "" And Len(Trim(Text5.Text)) > 30 Then
        MsgBox "Contact Person cannot exceed 30 characters.", vbExclamation, "Validation Error"
        Text5.SetFocus
        Exit Function
    End If

    If Trim(Text6.Text) <> "" And Len(Trim(Text6.Text)) > 30 Then
        MsgBox "Account Holder Name cannot exceed 30 characters.", vbExclamation, "Validation Error"
        Text6.SetFocus
        Exit Function
    End If

    If Trim(Text7.Text) <> "" Then
        If Not IsValidGST(UCase(Trim(Text7.Text))) Then
            MsgBox "GST Number must be 15 characters.", vbExclamation, "Validation Error"
            Text7.SetFocus
            Exit Function
        End If
    End If

    If Trim(Text8.Text) <> "" And Len(Trim(Text8.Text)) > 100 Then
        MsgBox "Address cannot exceed 100 characters.", vbExclamation, "Validation Error"
        Text8.SetFocus
        Exit Function
    End If

    If Trim(Text9.Text) <> "" And Len(Trim(Text9.Text)) > 40 Then
        MsgBox "Branch Name cannot exceed 40 characters.", vbExclamation, "Validation Error"
        Text9.SetFocus
        Exit Function
    End If

    If Trim(Text10.Text) <> "" And Len(Trim(Text10.Text)) > 40 Then
        MsgBox "Bank Name cannot exceed 40 characters.", vbExclamation, "Validation Error"
        Text10.SetFocus
        Exit Function
    End If

    If Trim(Text4.Text) <> "" Then
        If Not IsDigitsOnly(Text4.Text) Then
            MsgBox "Account Number must contain digits only.", vbExclamation, "Validation Error"
            Text4.SetFocus
            Exit Function
        End If
        If Len(Trim(Text4.Text)) < 6 Or Len(Trim(Text4.Text)) > 18 Then
            MsgBox "Account Number should be between 6 and 18 digits.", vbExclamation, "Validation Error"
            Text4.SetFocus
            Exit Function
        End If
    End If

    If Trim(Text12.Text) <> "" Then
        If IsValidIFSC(UCase(Trim(Text12.Text))) = False Then
            MsgBox "IFSC Code must be 11 characters.", vbExclamation, "Validation Error"
            Text12.SetFocus
            Exit Function
        End If
    End If

    ValidateData = True
End Function

'========================================================
' GENERATE ID
'========================================================
Private Sub GenerateID()
    On Error GoTo ErrHandler

    Dim rs As New ADODB.Recordset
    Dim num As Long

    If C Is Nothing Then Call CONN
    If C.State = 0 Then Call CONN

    rs.Open "SELECT NVL(MAX(TO_NUMBER(SUBSTR(SUP_ID,4))),0) FROM SUP_DET", C, adOpenStatic, adLockReadOnly
    num = rs.Fields(0).Value + 1

    Label3.Caption = "SUP" & Format(num, "000")
    rs.Close
    Set rs = Nothing
    Exit Sub

ErrHandler:
    MsgBox "ID generation error: " & Err.Description, vbCritical, "System Error"
End Sub

'========================================================
' ADD NEW  - COMMAND1
'========================================================
Private Sub Command1_Click()
    Call ClearData
    Call GenerateID
    Call SetAddMode
End Sub

'========================================================
' UPDATE - COMMAND2
'========================================================
Private Sub Command2_Click()
    On Error GoTo UpErr

    If mMode <> "EDIT" Then
        MsgBox "Please search a record first.", vbExclamation, "Update Error"
        Exit Sub
    End If

    If Trim(Label3.Caption) = "" Then
        MsgBox "Supplier ID is missing.", vbExclamation, "Update Error"
        Exit Sub
    End If

    If ValidateData = False Then Exit Sub

    SQL = "UPDATE SUP_DET SET " & _
          "SUP_PH=" & Trim(Text3.Text) & "," & _
          "CON_PER='" & Esc(Text5.Text) & "'," & _
          "CON_NO=" & IIf(Trim(Text11.Text) = "", "NULL", Trim(Text11.Text)) & "," & _
          "CMP_NM='" & Esc(Text2.Text) & "'," & _
          "ADDR='" & Esc(Text8.Text) & "'," & _
          "BNK_NM='" & Esc(Text10.Text) & "'," & _
          "BR_NM='" & Esc(Text9.Text) & "'," & _
          "AC_HLDR='" & Esc(Text6.Text) & "'," & _
          "AC_NO=" & IIf(Trim(Text4.Text) = "", "NULL", Trim(Text4.Text)) & "," & _
          "IFSC='" & Esc(UCase(Trim(Text12.Text))) & "'," & _
          "GST_NO='" & Esc(UCase(Trim(Text7.Text))) & "' " & _
          "WHERE SUP_ID='" & Esc(Label3.Caption) & "'"

    C.Execute SQL

    MsgBox "Supplier details updated successfully.", vbInformation, "Update Success"

    Call LoadData
    Call ClearData
    Call SetStartupMode
    Exit Sub

UpErr:
    MsgBox "Update error: " & Err.Description, vbCritical, "System Error"
End Sub

'========================================================
' CLEAR - COMMAND3
'========================================================
Private Sub Command3_Click()
    Call ClearData
    Call SetStartupMode
End Sub

'========================================================
' SAVE - COMMAND4
'========================================================
Private Sub Command4_Click()
    On Error GoTo SaveErr

    If mMode <> "ADD" Then
        MsgBox "Please click Add New first.", vbExclamation, "Save Error"
        Exit Sub
    End If

    If ValidateData = False Then Exit Sub

    SQL = "INSERT INTO SUP_DET VALUES(" & _
          "'" & Esc(Label3.Caption) & "'," & _
          Trim(Text3.Text) & "," & _
          "'" & Esc(Text5.Text) & "'," & _
          IIf(Trim(Text11.Text) = "", "NULL", Trim(Text11.Text)) & "," & _
          "'" & Esc(Text2.Text) & "'," & _
          "'" & Esc(Text8.Text) & "'," & _
          "'" & Esc(Text10.Text) & "'," & _
          "'" & Esc(Text9.Text) & "'," & _
          "'" & Esc(Text6.Text) & "'," & _
          IIf(Trim(Text4.Text) = "", "NULL", Trim(Text4.Text)) & "," & _
          "'" & Esc(UCase(Trim(Text12.Text))) & "'," & _
          "'" & Esc(UCase(Trim(Text7.Text))) & "')"

    C.Execute SQL

    MsgBox "Supplier details saved successfully.", vbInformation, "Save Success"

    Call LoadData
    Call ClearData
    Call SetStartupMode
    Exit Sub

SaveErr:
    MsgBox "Save error: " & Err.Description, vbCritical, "System Error"
End Sub

'========================================================
' DELETE - COMMAND5
'========================================================
Private Sub Command5_Click()
    On Error GoTo DelErr

    If mMode <> "EDIT" Then
        MsgBox "Please search a record first.", vbExclamation, "Delete Error"
        Exit Sub
    End If

    If Trim(Label3.Caption) = "" Then
        MsgBox "Supplier ID is missing.", vbExclamation, "Delete Error"
        Exit Sub
    End If

    If MsgBox("Do you want to delete this supplier record?", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
        C.Execute "DELETE FROM SUP_DET WHERE SUP_ID='" & Esc(Label3.Caption) & "'"
        MsgBox "Supplier record deleted successfully.", vbInformation, "Delete Success"
    End If

    Call LoadData
    Call ClearData
    Call SetStartupMode
    Exit Sub

DelErr:
    MsgBox "Delete error: " & Err.Description, vbCritical, "System Error"
End Sub

'========================================================
' EXIT - COMMAND6
'========================================================
Private Sub Command6_Click()
    Unload Me
End Sub

'========================================================
' SEARCH - COMMAND7
'========================================================
Private Sub Combo1_Click()
    On Error GoTo ErrHandler

    Dim rsS As New ADODB.Recordset

    If Combo1.Text = "" Then Exit Sub

    Combo2.Clear
    Combo2.Enabled = True
    Command7.Enabled = False

    If C Is Nothing Then Call CONN
    If C.State = 0 Then Call CONN

    If Combo1.Text = "ID" Then
        SQL = "SELECT SUP_ID FROM SUP_DET ORDER BY SUP_ID ASC"
    End If

    rsS.Open SQL, C, adOpenStatic, adLockReadOnly

    Do While Not rsS.EOF
        Combo2.AddItem rsS.Fields(0).Value
        rsS.MoveNext
    Loop

    rsS.Close
    Set rsS = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Search list load error: " & Err.Description, vbCritical, "System Error"
End Sub

Private Sub Combo2_Click()
    If Trim(Combo2.Text) <> "" Then
        Command7.Enabled = True
    End If
End Sub

Private Sub Command7_Click()
    On Error GoTo ErrHandler

    If Combo1.Text = "" Or Combo2.Text = "" Then
        MsgBox "Please select search criteria.", vbExclamation, "Search Validation"
        Exit Sub
    End If

    If C Is Nothing Then Call CONN
    If C.State = 0 Then Call CONN

    If Not R Is Nothing Then
        If R.State = 1 Then R.Close
    End If

    Set R = New ADODB.Recordset
    R.CursorLocation = adUseClient

    If Combo1.Text = "ID" Then
        SQL = "SELECT * FROM SUP_DET WHERE SUP_ID='" & Esc(Combo2.Text) & "'"
        
    End If

    R.Open SQL, C, adOpenStatic, adLockReadOnly

    If R.EOF Then
        MsgBox "Record not found.", vbExclamation, "Search Result"
        Exit Sub
    End If

    Set DataGrid1.DataSource = R

    Label3.Caption = NzField(R("SUP_ID"))
    Text3.Text = NzField(R("SUP_PH"))
    Text5.Text = NzField(R("CON_PER"))
    Text11.Text = NzField(R("CON_NO"))
    Text2.Text = NzField(R("CMP_NM"))
    Text8.Text = NzField(R("ADDR"))
    Text10.Text = NzField(R("BNK_NM"))
    Text9.Text = NzField(R("BR_NM"))
    Text6.Text = NzField(R("AC_HLDR"))
    Text4.Text = NzField(R("AC_NO"))
    Text12.Text = NzField(R("IFSC"))
    Text7.Text = NzField(R("GST_NO"))

    Call SetEditMode
    MsgBox "Record found successfully.", vbInformation, "Search Success"
    Exit Sub

ErrHandler:
    MsgBox "Search error: " & Err.Description, vbCritical, "System Error"
End Sub

'========================================================
' CLEAR DATA
'========================================================
Private Sub ClearData()
    Label3.Caption = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""

    Combo1.ListIndex = -1
    Combo2.Clear
    Combo2.Text = ""
    Combo2.Enabled = False
    Command7.Enabled = False
End Sub

'========================================================
' TEXT FORMAT
'=======================================================

Private Sub Text2_LostFocus()
    Text2.Text = ProperCaseText(Text2.Text)
End Sub

Private Sub Text5_LostFocus()
    Text5.Text = ProperCaseText(Text5.Text)
End Sub

Private Sub Text6_LostFocus()
    Text6.Text = ProperCaseText(Text6.Text)
End Sub

Private Sub Text8_LostFocus()
    Text8.Text = ProperCaseText(Text8.Text)
End Sub

Private Sub Text9_LostFocus()
    Text9.Text = ProperCaseText(Text9.Text)
End Sub

Private Sub Text10_LostFocus()
    Text10.Text = ProperCaseText(Text10.Text)
End Sub

Private Sub Text7_LostFocus()
    Text7.Text = UCase(Trim(Text7.Text))
End Sub

Private Sub Text12_LostFocus()
    Text12.Text = UCase(Trim(Text12.Text))
End Sub

'========================================================
' KEY PRESS RESTRICTIONS
'========================================================


Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text7.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 38 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text5.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If Len(Text3.Text) >= 10 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text11.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If Len(Text11.Text) >= 10 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text8.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If Len(Text7.Text) >= 15 Then
        KeyAscii = 0
        Exit Sub
    End If

    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Or _
       (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    End If

    KeyAscii = 0
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text10.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text9.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text6.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text4.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Or KeyAscii = 32 Then Exit Sub
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text12.SetFocus
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If Len(Text4.Text) >= 18 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Command4.Enabled = True Then
            Command4.SetFocus
        ElseIf Command2.Enabled = True Then
            Command2.SetFocus
        End If
        Exit Sub
    End If

    If KeyAscii = 8 Then Exit Sub
    If Len(Text12.Text) >= 11 Then
        KeyAscii = 0
        Exit Sub
    End If

    If (KeyAscii >= 65 And KeyAscii <= 90) Or _
       (KeyAscii >= 97 And KeyAscii <= 122) Or _
       (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    End If

    KeyAscii = 0
End Sub

