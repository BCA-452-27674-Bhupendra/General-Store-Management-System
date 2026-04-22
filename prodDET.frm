VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFC0C0&
   Caption         =   "GENERAL STORE MANAGEMENT SYSTEM"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15255
   LinkTopic       =   "form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   15255
   Begin VB.CommandButton Command10 
      Caption         =   "Add"
      Height          =   375
      Left            =   4680
      TabIndex        =   32
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7920
      TabIndex        =   31
      Top             =   1560
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "prodDET.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   4471
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
   Begin VB.CommandButton Command9 
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5280
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2280
      TabIndex        =   15
      Top             =   3000
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2280
      TabIndex        =   14
      Top             =   3720
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2280
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   10920
      TabIndex        =   12
      Top             =   720
      Width           =   3735
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print collective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Print selective"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   840
         TabIndex        =   17
         Text            =   "Select ID"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9960
      Top             =   6960
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
      Connect         =   "Provider=MSDAORA.1;Password=PRJ2531B;User ID=PRJ2531B;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=PRJ2531B;User ID=PRJ2531B;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM PRODUCT"
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
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   1320
      TabIndex        =   28
      Top             =   6120
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16761024
      Appearance      =   1
      StartOfWeek     =   146341889
      CurrentDate     =   46032
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      Left            =   4920
      TabIndex        =   30
      Top             =   1560
      Width           =   660
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT DETAIL"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      Width           =   3570
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand name"
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
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim activebox As String

Private Sub Combo1_CLICK()
Combo3.Clear
Combo2.Clear
If Combo1.Text = "Grocery" Then
Combo3.AddItem "Edible Oil"
Combo3.AddItem "Rice"
Combo3.AddItem "Flour"
Combo3.AddItem "Pulses"
ElseIf Combo1.Text = "Personal Care" Then
Combo3.AddItem "Soap"
Combo3.AddItem "Shampoo"
Combo3.AddItem "Toothpaste"
Combo3.AddItem "Hair Oil"
ElseIf Combo1.Text = "Household" Then
Combo3.AddItem "Detergent"
Combo3.AddItem "Phenyl"
Combo3.AddItem "Dishwash"
ElseIf Combo1.Text = "Beverages" Then
Combo3.AddItem "Soft Drink"
End If
End Sub


Private Sub Combo3_Click()
Combo2.Clear
If Combo3.Text = "Soap" Then
Combo2.AddItem "Dove"
Combo2.AddItem "Pears"
Combo2.AddItem "Lifeboy"
Combo2.AddItem "Lux"
ElseIf Combo3.Text = "Edible Oil" Then
Combo2.AddItem "Fortune"
Combo2.AddItem "Dhara"
Combo2.AddItem "Ghani"
Combo2.AddItem "Engine"
Combo2.AddItem "Saffola"
ElseIf Combo3.Text = "Detergent" Then
Combo2.AddItem "Surf Excel"
Combo2.AddItem "Ariel"
Combo2.AddItem "Tide"
Combo2.AddItem "Rin"
ElseIf Combo3.Text = "Soft Drink" Then
Combo2.AddItem "Coca Cola"
Combo2.AddItem "Sting"
Combo2.AddItem "thums up"
Combo2.AddItem "Sprite"
Combo2.AddItem "Maja"
Combo2.AddItem "Amul Lassi"
ElseIf Combo3.Text = "Shampoo" Then
Combo2.AddItem "Dove"
Combo2.AddItem "Head & shoulders"
Combo2.AddItem "Pantene"
Combo2.AddItem "Clinic Plus"
ElseIf Combo3.Text = "Toothpaste" Then
Combo2.AddItem "Closeup"
Combo2.AddItem "Patanjali Dant Kanti"
Combo2.AddItem "Sensodyne"
Combo2.AddItem "Colgate"
ElseIf Combo3.Text = "Hair Oil" Then
Combo2.AddItem "Bajaj Almond Drops"
Combo2.AddItem "Navratna"
Combo2.AddItem "Himalaya"
Combo2.AddItem "Dabur Amla"
ElseIf Combo3.Text = "Phenyl" Then
Combo2.AddItem "Lizol"
Combo2.AddItem "Harpic"
Combo2.AddItem "Dettol"
Combo2.AddItem "Home ninza"
ElseIf Combo3.Text = "Dishwash" Then
Combo2.AddItem "Pril"
Combo2.AddItem "Vim"
Combo2.AddItem "Exo"
ElseIf Combo3.Text = "Flour" Then
Combo2.AddItem "Ashirvaad"
Combo2.AddItem "Patanjali Atta"
Combo2.AddItem "Rajdhani Atta"
ElseIf Combo3.Text = "Rice" Then
Combo2.AddItem "India Gate"
Combo2.AddItem "Daawat"
Combo2.AddItem "Patanjali Rice"
Combo2.AddItem "Royal"
ElseIf Combo3.Text = "Skin Care" Then
Combo2.AddItem "Mamaearth"
Combo2.AddItem "Garnier"
Combo2.AddItem "Himalaya"
Combo2.AddItem "Lakme"
Combo2.AddItem "Pond's"
ElseIf Combo3.Text = "Pulses" Then
Combo2.AddItem "Tata Sampann"
Combo2.AddItem "Ashirvaad"
Combo2.AddItem "Patanjali"
Combo2.AddItem "Rajdhani"
Combo2.AddItem "Fortune"
End If
Combo5.Clear
If Combo3.Text = "Rice" Or Combo3.Text = "Flour" Or Combo3.Text = "Pulses" Then
    Combo5.AddItem "Kg"
    Combo5.AddItem "Gram"
    Combo5.AddItem "Bag"

ElseIf Combo3.Text = "Edible Oil" Or Combo3.Text = "Hair Oil" Then
    Combo5.AddItem "Liter"
    Combo5.AddItem "Milliliter"
    Combo5.AddItem "Bottle"
ElseIf Combo3.Text = "Soap" Or Combo3.Text = "Shampoo" Or Combo3.Text = "Toothpaste" Then
    Combo5.AddItem "Piece"
    Combo5.AddItem "Packet"
    Combo5.AddItem "Box"

ElseIf Combo3.Text = "Detergent" Or Combo3.Text = Dishwash Then
    Combo5.AddItem "Kg"
    Combo5.AddItem "Gram"
    ElseIf Combo3.Text = "Phenyl" Then
    Combo5.AddItem "Litre"
    Combo5.AddItem "Bottle"

ElseIf Combo3.Text = "Soft Drink" Then
    Combo5.AddItem "Liter"
    Combo5.AddItem "Bottle"
    Combo5.AddItem "Carton"

End If
End Sub

Private Sub Combo4_GotFocus()
Combo4.Clear
SQL = "SELECT PR_ID FROM PRODUCT ORDER BY PR_ID"
Set R = C.Execute(SQL)
Do While Not R.EOF
Combo4.AddItem R.Fields("PR_ID")
R.MoveNext
Loop
End Sub
Private Sub Command1_Click()
SQL = "SELECT * FROM PRODUCT WHERE PR_ID='" + Combo4.Text + "'"
Set R = C.Execute(SQL)
Label13.Caption = R.Fields(0)
Text1.Text = R.Fields(1)
Combo1.Text = R.Fields(2)
Combo2.Text = R.Fields(3)
Text4.Text = R.Fields(4)
Text5.Text = R.Fields(5)
Text6.Text = R.Fields(6)
Text7.Text = R.Fields(7)
Combo5.Text = R.Fields(8)
Combo3.Text = R.Fields(9)
Command5.Enabled = True
Command6.Enabled = True
End Sub

Private Sub Command2_Click()
DataEnvironment1.Command1 Combo4.Text
DataReport1.Show
DataReport1.Refresh
Set DataEnvironment1 = Nothing
End Sub

Private Sub Command3_Click()
DataReport3.Show
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or _
   Text4.Text = "" Or Text5.Text = "" Or Combo5.Text = "" Or _
   Combo3.Text = "" Then

    MsgBox "please fill all required fields", vbExclamation
    Exit Sub
End If
C.Execute "insert into product (PR_ID, PR_NM, CTG, BR_NM, PRICE, WT, UNIT, TYPE) values ('" & _
Label13.Caption & "','" & Text1.Text & "','" & Combo1.Text & "','" & Combo2.Text & "'," & _
Val(Text4.Text) & "," & Val(Text5.Text) & ",'" & Combo5.Text & "','" & Combo3.Text & "')"
' STOCK TABLE INSERT (QTY = 0)
'C.Execute "insert into stock (PR_ID, PR_NM, CTG, BR_NM, PRICE, WT, UNIT, TYPE) values('" & Label13.Caption & "','" & Text1.Text & "','" & Combo1.Text & "','" & Combo2.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & "','" & Combo5.Text & "','" & Combo3.Text & "', 0)"

MsgBox "Product Saved"



Label13.Caption = ""
Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo3.Text = ""
Combo5.Text = ""
Text1.SetFocus
Adodc1.Refresh
End Sub


Private Sub Command5_Click()

If Label13.Caption = "" Then
    MsgBox "Please select PRODUCT ID to update", vbExclamation
    Exit Sub
End If

CONN

SQL = "UPDATE PRODUCT SET " & _
      "PR_NM='" & Text1.Text & "', " & _
      "CTG='" & Combo1.Text & "', " & _
      "BR_NM='" & Combo2.Text & "', " & _
      "PRICE=" & Text4.Text & ", " & _
      "WT='" & Text5.Text & "', " & _
      "EXP_DT=TO_DATE('" & Format(MonthView1.Value, "dd-mm-yyyy") & "','DD-MM-YYYY'), " & _
      "MFG_DT=TO_DATE('" & Format(MonthView1.Value, "dd-mm-yyyy") & "','DD-MM-YYYY'), " & _
      "TYPE='" & Combo3.Text & "' " & _
      "WHERE PR_ID='" & Label13.Caption & "'"

C.Execute SQL
MsgBox "Record updated successfully", vbInformation

Label13.Caption = ""
Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo5.Text = ""

Combo3.Text = ""

Text1.SetFocus

End Sub

Private Sub Command6_Click()
If Label13.Caption = "" Then
MsgBox "No Record Selected to Delete", vbExclamation
Exit Sub
End If
If MsgBox("Are you sure you want to delete this record ", vbYesNo + vbQuestion) = vbNo Then Exit Sub
CONN
C.Execute "delete from product where pr_id='" & Label13.Caption & "'"
MsgBox "Record deleted successfully", vbInformation
Label13.Caption = ""
Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo5.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo3.Text = ""
Adodc1.Refresh
End Sub

Private Sub Command7_Click()
Label13.Caption = ""
Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo3.Text = ""
Combo5.Text = ""

End Sub

Private Sub Command8_Click()
Unload Me
End Sub



Private Sub Command9_Click()
If Trim(Text1.Text) = "" Then
MsgBox "Enter Product Name"
Text1.SetFocus
Exit Sub
End If
If Trim(Combo2.Text) = "" Then
MsgBox "Select Brand Name"
Combo2.SetFocus
Exit Sub
End If
If Trim(Text5.Text) = "" Then
MsgBox "Enter Weight"
Text5.SetFocus
Exit Sub
End If
Label13.Caption = UCase(Left(Text1.Text, 3)) & UCase(Left(Combo2.Text, 3)) & Text5.Text
MsgBox "Product ID generated"
End Sub

Private Sub DataGrid1_Click()
SQL = "SELECT * FROM PRODUCT WHERE PR_ID='" + DataGrid1.Text + "'"
Set R = C.Execute(SQL)
Label13.Caption = R.Fields(0)
Text1.Text = R.Fields(1)
Combo1.Text = R.Fields(2)
Combo2.Text = R.Fields(3)
Text4.Text = R.Fields(4)
Text5.Text = R.Fields(5)
Text6.Text = R.Fields(6)
Text7.Text = R.Fields(7)
Combo3.Text = R.Fields(10)
Command5.Enabled = True
Command6.Enabled = True
End Sub

Private Sub Form_Load()
CONN
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo1.AddItem "Grocery"
Combo1.AddItem "Personal Care"
Combo1.AddItem "Household"
Combo1.AddItem "Beverages"
Combo5.AddItem "Kg"
Combo5.AddItem "Gram"
Combo5.AddItem "Liter"
Combo5.AddItem "Milliliter"
Combo5.AddItem "Piece"
Combo5.AddItem "Dozen"
Combo5.AddItem "Packet"
Combo5.AddItem "Box"
Combo5.AddItem "Carton"
Combo5.AddItem "Bag"
Combo5.AddItem "Bottle"
Combo5.AddItem "Tin"
Adodc1.RecordSource = "select * from PRODUCT ORDER BY PR_ID ASC"
Adodc1.Refresh
Command5.Enabled = False
Command6.Enabled = False
MonthView1.Visible = False


End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
If activebox = "EXP" Then
Text6.Text = Format(MonthView1.Value, "dd/mmm/yyyy")
ElseIf activebox = "MFG" Then
Text7.Text = Format(MonthView1.Value, "dd/mmm/yyyy")
End If

MonthView1.Visible = False
End Sub



Private Sub Text6_GotFocus()
MonthView1.Visible = True
MonthView1.Left = Text6.Left + Text6.Width + 100
MonthView1.Top = Text6.Top
activebox = "EXP"
MonthView1.ZOrder 0
End Sub
Private Sub Text7_GotFocus()
MonthView1.Visible = True
MonthView1.Left = Text7.Left + Text7.Width + 100
MonthView1.Top = Text7.Top
activebox = "MFG"
End Sub

