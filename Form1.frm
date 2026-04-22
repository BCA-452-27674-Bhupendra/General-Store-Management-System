VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   14940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   495
      Left            =   6240
      TabIndex        =   53
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add New"
      Height          =   495
      Left            =   4440
      TabIndex        =   52
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2040
      TabIndex        =   51
      Text            =   "Text5"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   50
      Text            =   "Text4"
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   49
      Text            =   "Text3"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12840
      TabIndex        =   48
      Text            =   "Text2"
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   11040
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   7080
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   12000
      TabIndex        =   37
      Text            =   "Combo3"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   35
      Text            =   "Combo2"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   14175
      Begin VB.CommandButton Command1 
         Caption         =   "Add Product"
         Height          =   375
         Left            =   6480
         TabIndex        =   41
         Top             =   3720
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   0
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox List9 
         Height          =   2400
         Left            =   12600
         TabIndex        =   25
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List8 
         Height          =   2400
         Left            =   11040
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List7 
         Height          =   2400
         Left            =   9480
         TabIndex        =   23
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List6 
         Height          =   2400
         Left            =   7920
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List5 
         Height          =   2400
         Left            =   6360
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List4 
         Height          =   2400
         Left            =   4800
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List3 
         Height          =   2400
         Left            =   3240
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   1680
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label24"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   12840
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label23"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11160
         TabIndex        =   33
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label22"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   9600
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label21"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8040
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label20"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6480
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label19"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4800
         TabIndex        =   29
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label6"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total Price"
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
         Left            =   12600
         TabIndex        =   16
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "GST%"
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
         Left            =   11400
         TabIndex        =   15
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
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
         Left            =   9840
         TabIndex        =   14
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Price"
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
         Left            =   8280
         TabIndex        =   13
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         Left            =   6720
         TabIndex        =   12
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Brand"
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
         Left            =   5040
         TabIndex        =   11
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Category"
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
         Left            =   3360
         TabIndex        =   10
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Left            =   2040
         TabIndex        =   9
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Product ID"
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
         TabIndex        =   8
         Top             =   120
         Width           =   1470
      End
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Total Amount"
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
      Left            =   12720
      TabIndex        =   46
      Top             =   6480
      Width           =   1905
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Total GST"
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
      Left            =   10920
      TabIndex        =   45
      Top             =   6480
      Width           =   1395
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "Net Amount"
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
      TabIndex        =   44
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
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
      Left            =   360
      TabIndex        =   43
      Top             =   7440
      Width           =   690
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "Advance"
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
      TabIndex        =   42
      Top             =   6720
      Width           =   1170
   End
   Begin VB.Label Label29 
      Caption         =   "Label29"
      Height          =   495
      Left            =   9240
      TabIndex        =   40
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label28 
      Caption         =   "Label28"
      Height          =   375
      Left            =   5160
      TabIndex        =   39
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label27 
      Caption         =   "Label27"
      Height          =   495
      Left            =   6240
      TabIndex        =   38
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label25 
      Caption         =   "Label25"
      Height          =   495
      Left            =   2040
      TabIndex        =   36
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Left            =   7920
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Mode Of Payment"
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
      Left            =   11640
      TabIndex        =   5
      Top             =   600
      Width           =   2475
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Phone No"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Left            =   4200
      TabIndex        =   3
      Top             =   720
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Customer ID"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sale ID"
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
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sale Detail"
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
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
