VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00FFC0C0&
   Caption         =   "GENERAL STORE MANAGEMENT SYSTEM"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14910
   LinkTopic       =   "Form7"
   ScaleHeight     =   8775
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo6 
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
      Left            =   2160
      TabIndex        =   39
      Text            =   "Combo6"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   12240
      TabIndex        =   37
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   12240
      TabIndex        =   36
      Top             =   2880
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   12240
      TabIndex        =   35
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   152174593
      CurrentDate     =   46084
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12240
      TabIndex        =   34
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   152174593
      CurrentDate     =   46084
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   7200
      TabIndex        =   33
      Text            =   "Combo5"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add new"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":0000
      Height          =   2295
      Left            =   120
      TabIndex        =   20
      Top             =   6360
      Width           =   14655
      _ExtentX        =   25850
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   7800
      TabIndex        =   14
      Top             =   3720
      Width           =   6255
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
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   3600
         TabIndex        =   27
         Text            =   "Combo4"
         Top             =   600
         Width           =   2055
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
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print Collective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print Selective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Value"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3840
         TabIndex        =   19
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Search By "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   720
         TabIndex        =   18
         Top             =   120
         Width           =   1410
      End
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2160
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   2400
      Width           =   2055
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
      Left            =   2160
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   7200
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   7200
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   7200
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   2160
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4560
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   "Provider=MSDAORA.1;Password=PRJ2531B;User ID=PRJ2531B;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=PRJ2531B;User ID=PRJ2531B;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM PUR_ORD_DETAIL"
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
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Date"
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
      Left            =   9600
      TabIndex        =   32
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dues"
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
      Left            =   9600
      TabIndex        =   31
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Payment"
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
      Left            =   9600
      TabIndex        =   30
      Top             =   3000
      Width           =   2445
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
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
      Left            =   9600
      TabIndex        =   29
      Top             =   1560
      Width           =   1890
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode of payment"
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
      Left            =   4560
      TabIndex        =   28
      Top             =   3120
      Width           =   2400
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE DETAIL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   3810
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1605
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grand total"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4560
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4560
      TabIndex        =   4
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   1230
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1320
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   1350
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public C As ADODB.Connection
Public R As ADODB.Recordset
Public SQL As String
Public Function CONN()
Set C = New ADODB.Connection
C.Open "Provider=MSDAORA.1;User ID=PRJ2531B/PRJ2531B;Persist Security Info=False"
Set R = New ADODB.Recordset
End Function

Private Sub Combo3_Click()
If Combo3.Text = "" Then Exit Sub
Combo4.Enabled = True

    Call CONN
    Combo4.Clear
    Set R = New ADODB.Recordset

    If Combo3.Text = "Order ID" Then
        R.Open "select oid from product_order_detail", C
        While Not R.EOF
        Combo4.AddItem R.Fields(0).Value
        R.MoveNext
        Wend
        R.Close
        Exit Sub
        End If
        If Combo3.Text = "Supplier ID" Then

        R.Open "select sid from supplier", C
        While Not R.EOF
        Combo4.AddItem R.Fields(0).Value
        R.MoveNext
        Wend
        R.Close
        Exit Sub
        End If
        If Combo3.Text = "Product ID" Then

        R.Open "select pr_id from product", C
        While Not R.EOF
        Combo4.AddItem R.Fields(0).Value
        R.MoveNext
        Wend
        R.Close
        Exit Sub
        End If
         If Combo3.Text = "Order Date" Then

        R.Open "select ord_date from pur_ord_detail", C
        While Not R.EOF
        Combo4.AddItem R.Fields(0).Value
        R.MoveNext
        Wend
        R.Close
        Exit Sub
        End If
    
    
    
End Sub

Private Sub Command3_Click()
If Combo6.Text = "" Then
MsgBox "Please select record to delete", vbInformation
Exit Sub
End If

If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then

SQL = "DELETE FROM PUR_ORD_DETAIL WHERE OID='" & Combo6.Text & "'"

Set R = C.Execute(SQL)

MsgBox "Record Deleted Successfully", vbInformation
Combo1.Text = ""
Combo6.Text = ""
Combo2.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo5.Text = ""
Text6.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Text5.Text = ""
Text1.SetFocus
End If
Adodc1.Refresh
End Sub

Private Sub Command5_Click()
If Combo1.Text = "" Or Combo2.Text = "" Or Combo6.Text = "" Or _
   Combo5.Text = "" Or DTPicker1.Value = Date Or DTPicker2.Value = Date Then
MsgBox "please fill all required fields", vbExclamation
Exit Sub
End If
SQL = "INSERT INTO PUR_ORD_DETAIL VALUES('" & Combo1.Text & "','" & Combo6.Text & "','" & Combo2.Text & "'," & Val(Text1.Text) & "," & Val(Text2.Text) & "," & Val(Text3.Text) & "," & Val(Text4.Text) & ",'" & Combo5.Text & "',TO_DATE('" & Format(DTPicker1.Value, "dd-mm-yyyy") & "','DD-MM-YYYY')," & Val(Text5.Text) & "," & Val(Text6.Text) & ",TO_DATE('" & Format(DTPicker2.Value, "dd-mm-yyyy") & "','DD-MM-YYYY'))"

Set R = C.Execute(SQL)
MsgBox "record saved successfully", vbInformation
Combo1.Text = ""
Combo6.Text = ""
Combo2.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo5.Text = ""
Text6.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Text5.Text = ""
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
If Combo6.Text = "" Then
MsgBox "Select record to update", vbInformation
Exit Sub
End If

SQL = "UPDATE PUR_ORD_DETAIL SET " & _
"SID='" & Combo1.Text & "'," & _
"PR_ID='" & Combo2.Text & "'," & _
"QTY=" & Val(Text1.Text) & "," & _
"AMT=" & Val(Text2.Text) & "," & _
"DISC=" & Val(Text3.Text) & "," & _
"GRND_TOL=" & Val(Text4.Text) & "," & _
"MOP='" & Combo5.Text & "'," & _
"DEL_DATE=TO_DATE('" & Format(DTPicker1.Value, "dd-mm-yyyy") & "','DD-MM-YYYY')," & _
"ADVA_PAYM=" & Val(Text5.Text) & "," & _
"DUES=" & Val(Text6.Text) & "," & _
"ORD_DATE=TO_DATE('" & Format(DTPicker2.Value, "dd-mm-yyyy") & "','DD-MM-YYYY')" & _
" WHERE OID='" & Combo6.Text & "'"

Set R = C.Execute(SQL)



MsgBox "Record Updated Successfully", vbInformation
Adodc1.Refresh


End Sub

Private Sub Command7_Click()
Combo1.Text = ""
Combo6.Text = ""
Combo2.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo5.Text = ""
Text6.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Text5.Text = ""
End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Form_Load()
CONN
Combo4.Enabled = False
Combo3.AddItem "Order ID"
Combo3.AddItem "Supplier ID"
Combo3.AddItem "Product ID"
Combo3.AddItem "Order Date"
Combo5.AddItem "Cash"
Combo5.AddItem "Card"
Combo5.AddItem "Cheque"
SQL = "SELECT PR_ID FROM PRODUCT"
Set R = C.Execute(SQL)
Do While Not R.EOF
Combo2.AddItem R.Fields(0)
R.MoveNext
Loop
SQL = "SELECT SID FROM SUPPLIER"
Set R = C.Execute(SQL)
Do While Not R.EOF
Combo1.AddItem R.Fields(0)
R.MoveNext
Loop
SQL = "SELECT OID FROM PRODUCT_ORDER_DETAIL"
Set R = C.Execute(SQL)
Do While Not R.EOF
Combo6.AddItem R.Fields(0)
R.MoveNext
Loop
End Sub

