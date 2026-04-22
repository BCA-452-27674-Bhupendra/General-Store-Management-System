VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PURCHASE 
   BackColor       =   &H00FFC0C0&
   Caption         =   "GENERAL STORE MANAGEMENT SYSTEM"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15990
   LinkTopic       =   "Form10"
   ScaleHeight     =   8445
   ScaleWidth      =   15990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Height          =   4095
      Left            =   13920
      TabIndex        =   59
      Top             =   2040
      Width           =   1815
      Begin VB.CommandButton Command10 
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
         Height          =   660
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
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
         Height          =   660
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2520
         Width           =   1455
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
         Left            =   120
         TabIndex        =   62
         Text            =   "Combo5"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Value"
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
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add Product"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text14 
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
      Left            =   12840
      TabIndex        =   53
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox Text13 
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
      Left            =   10800
      TabIndex        =   52
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
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
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
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
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Remove Product"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add New"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6840
      Width           =   1215
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
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text10 
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
      TabIndex        =   42
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text9 
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
      TabIndex        =   41
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox Text7 
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
      TabIndex        =   37
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Height          =   4095
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   13935
      Begin VB.ListBox List8 
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
         Height          =   2010
         ItemData        =   "Form10.frx":0000
         Left            =   10920
         List            =   "Form10.frx":0002
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text12 
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
         Height          =   375
         Left            =   10800
         TabIndex        =   43
         Top             =   840
         Width           =   1455
      End
      Begin VB.ListBox List9 
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
         Height          =   2010
         ItemData        =   "Form10.frx":0004
         Left            =   12480
         List            =   "Form10.frx":0006
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ListBox List7 
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
         Height          =   2010
         ItemData        =   "Form10.frx":0008
         Left            =   9360
         List            =   "Form10.frx":000A
         TabIndex        =   34
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ListBox List6 
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
         Height          =   2010
         ItemData        =   "Form10.frx":000C
         Left            =   7800
         List            =   "Form10.frx":000E
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ListBox List5 
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
         Height          =   2010
         ItemData        =   "Form10.frx":0010
         Left            =   6240
         List            =   "Form10.frx":0012
         TabIndex        =   32
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ListBox List4 
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
         Height          =   2010
         ItemData        =   "Form10.frx":0014
         Left            =   4680
         List            =   "Form10.frx":0016
         TabIndex        =   31
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ListBox List3 
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
         Height          =   2010
         ItemData        =   "Form10.frx":0018
         Left            =   3120
         List            =   "Form10.frx":001A
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ListBox List2 
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
         Height          =   2010
         ItemData        =   "Form10.frx":001C
         Left            =   1560
         List            =   "Form10.frx":001E
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ListBox List1 
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
         Height          =   2010
         ItemData        =   "Form10.frx":0020
         Left            =   120
         List            =   "Form10.frx":0022
         TabIndex        =   28
         Top             =   1320
         Width           =   1335
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
         Height          =   405
         Left            =   12360
         TabIndex        =   27
         Top             =   840
         Width           =   1455
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
         Height          =   375
         Left            =   9240
         TabIndex        =   26
         Top             =   840
         Width           =   1455
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
         Height          =   375
         Left            =   7680
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   120
         TabIndex        =   24
         Text            =   "Combo3"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6120
         TabIndex        =   57
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4560
         TabIndex        =   56
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   55
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   54
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GST %"
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
         TabIndex        =   40
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   12360
         TabIndex        =   23
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   9720
         TabIndex        =   22
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   7920
         TabIndex        =   21
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   6480
         TabIndex        =   20
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3120
         TabIndex        =   18
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label15 
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
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   16
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   14160
      TabIndex        =   12
      Text            =   "Combo2"
      Top             =   840
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9480
      TabIndex        =   11
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   40828929
      CurrentDate     =   46111
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   40828929
      CurrentDate     =   46111
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1920
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T0TAL AMOUNT"
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
      Left            =   12840
      TabIndex        =   51
      Top             =   6120
      Width           =   2430
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL GST"
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
      Left            =   10800
      TabIndex        =   50
      Top             =   6120
      Width           =   1725
   End
   Begin VB.Label Label26 
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
      Left            =   240
      TabIndex        =   39
      Top             =   6960
      Width           =   690
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   38
      Top             =   6360
      Width           =   1170
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net AMOUNT"
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
      TabIndex        =   36
      Top             =   7560
      Width           =   1980
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT ORDER DETAIL"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   4155
   End
   Begin VB.Label Label8 
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
      Left            =   7440
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   3840
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   11520
      TabIndex        =   3
      Top             =   840
      Width           =   2475
   End
   Begin VB.Label Label3 
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
      Left            =   7440
      TabIndex        =   2
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label2 
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
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order ID"
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
      TabIndex        =   0
      Top             =   840
      Width           =   1230
   End
End
Attribute VB_Name = "PURCHASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_CLICK()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    ' ?? Supplier details
    rs.Open "SELECT SUP_PH, ADDR FROM SUP_DET WHERE SUP_ID='" & Combo1.Text & "'", C, adOpenKeyset, adLockReadOnly

    If Not rs.EOF Then
        Label13.Caption = rs!SUP_PH
        Label12.Caption = rs!ADDR
    End If

    rs.Close

    ' ?? Product combo fill (quotation based)
    Combo3.Clear

    rs.Open "SELECT DISTINCT Q.PR_ID FROM Q_PROD_DET Q, QUOTATION QT WHERE Q.Q_ID = QT.Q_ID AND QT.SUP_ID = '" & Combo1.Text & "'", C, adOpenKeyset, adLockReadOnly

    While Not rs.EOF
        Combo3.AddItem rs!PR_ID
        rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing

End Sub


Private Sub Combo3_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    rs.Open "SELECT PR_NM, CTG, BR_NM, TYPE, PRICE FROM PRODUCT WHERE PR_ID='" & Combo3.Text & "'", C, adOpenKeyset, adLockReadOnly

    If Not rs.EOF Then
        Label27.Caption = rs!PR_NM
        Label29.Caption = rs!CTG
        Label30.Caption = rs!BR_NM
        Label31.Caption = rs!Type
        Text5.Text = rs!price
    End If

    rs.Close
    Set rs = Nothing

End Sub
    

Private Sub Command1_Click()

    Dim i As Integer
    Dim SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    ' ?? Validation
    If Label10.Caption = "" Then
        MsgBox "Click Add New first"
        Exit Sub
    End If

    If Combo1.Text = "" Then
        MsgBox "Select Supplier"
        Exit Sub
    End If

    If List1.ListCount = 0 Then
        MsgBox "Add at least one product"
        Exit Sub
    End If

    Call CONN

    ' =========================
    ' ?? CHECK MASTER EXISTS
    ' =========================
    rs.Open "SELECT * FROM ORDER_MASTER WHERE ORD_ID='" & Label10.Caption & "'", C, adOpenKeyset, adLockReadOnly

    If rs.EOF Then
        ' ?? Insert only first time
        SQL = "INSERT INTO ORDER_MASTER VALUES('" & Label10.Caption & "'," & _
              "TO_DATE('" & DTPicker1.Value & "','DD-MM-YYYY')," & _
              "TO_DATE('" & DTPicker2.Value & "','DD-MM-YYYY')," & _
              "'" & Combo1.Text & "'," & _
              Val(Text14.Text) & "," & _
              Val(Text7.Text) & "," & _
              Val(Text10.Text) & "," & _
              "'" & Combo2.Text & "','ACTIVE')"

        C.Execute SQL
    End If

    rs.Close

    ' =========================
    ' ?? INSERT DETAILS (MULTIPLE PRODUCTS)
    ' =========================

    For i = 0 To List1.ListCount - 1

        ' ?? Duplicate check (same product already saved or not)
        rs.Open "SELECT * FROM ORDER_DETAILS WHERE ORD_ID='" & Label10.Caption & "' AND PR_ID='" & List1.List(i) & "'", C, adOpenKeyset, adLockReadOnly

        If rs.EOF Then
            ' ?? Insert only if not already present
            SQL = "INSERT INTO ORDER_DETAILS VALUES('" & Label10.Caption & "','" & _
                  List1.List(i) & "'," & _
                  Val(List6.List(i)) & "," & _
                  Val(List7.List(i)) & "," & _
                  Val(List8.List(i)) & "," & _
                  Val(List9.List(i)) & ")"

            C.Execute SQL
' =========================
' ?? STOCK UPDATE (FINAL FIX)
' =========================

Dim rsCheck As New ADODB.Recordset

rsCheck.Open "SELECT * FROM STOCK WHERE PR_ID='" & List1.List(i) & "'", C, adOpenKeyset, adLockReadOnly

If rsCheck.EOF Then
    ' ?? Insert
    SQL = "INSERT INTO STOCK(PR_ID, CURRENT_QTY, LAST_UPDATE) VALUES('" & _
          List1.List(i) & "'," & Val(List7.List(i)) & ",SYSDATE)"
    C.Execute SQL
Else
    ' ?? Update
    SQL = "UPDATE STOCK SET CURRENT_QTY = CURRENT_QTY + " & Val(List7.List(i)) & _
          ", LAST_UPDATE = SYSDATE WHERE PR_ID='" & List1.List(i) & "'"
    C.Execute SQL
End If

rsCheck.Close
Set rsCheck = Nothing
        End If

        rs.Close

    Next i

    MsgBox "Order Saved Successfully!", vbInformation

End Sub

Private Sub Command2_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim newID As String
    Dim lastID As String
    Dim num As Long

    ' ?? Connection call (VERY IMPORTANT)
    Call CONN

    ' ?? Recordset open (simple query)
    rs.Open "SELECT ORD_ID FROM ORDER_MASTER ORDER BY ORD_ID DESC", C, adOpenKeyset, adLockReadOnly

    If rs.EOF Then
        newID = "ORD001"
    Else
        lastID = rs.Fields("ORD_ID").Value
        num = Val(Mid(lastID, 4)) + 1
        newID = "ORD" & Format(num, "000")
    End If

    Label10.Caption = newID

    rs.Close
    Set rs = Nothing

    ' ?? Clear controls
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Label13.Caption = ""
    Label12.Caption = ""

    Label27.Caption = ""
    Label29.Caption = ""
    Label30.Caption = ""
    Label31.Caption = ""

    Text5.Text = ""
    Text6.Text = ""
    Text12.Text = ""
    Text8.Text = ""

    Text7.Text = ""
    Text10.Text = ""
    Text9.Text = ""

    Text13.Text = ""
    Text14.Text = ""

    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    List7.Clear
    List8.Clear
    List9.Clear

    DTPicker1.Value = Date
    DTPicker2.Value = Date

    MsgBox "New Order Ready!", vbInformation
End Sub

Private Sub Command3_Click()

    Dim i As Integer
    Dim prID As String
    Dim SQL As String

    If List1.ListIndex = -1 Then
        MsgBox "Select a product to remove"
        Exit Sub
    End If

    i = List1.ListIndex
    prID = List1.List(i)

    Call CONN

    ' ?? DELETE FROM DATABASE
    SQL = "DELETE FROM ORDER_DETAILS WHERE ORD_ID='" & Label10.Caption & "' AND PR_ID='" & prID & "'"
    C.Execute SQL

    ' ?? REMOVE FROM LIST
    List1.RemoveItem i
    List2.RemoveItem i
    List3.RemoveItem i
    List4.RemoveItem i
    List5.RemoveItem i
    List6.RemoveItem i
    List7.RemoveItem i
    List8.RemoveItem i
    List9.RemoveItem i

    ' ?? RECALCULATE
    Call CalculateGrandTotal

    MsgBox "Product Removed Permanently!", vbInformation

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()

    ' ?? Combo clear
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1

    ' ?? Supplier labels clear
    Label11.Caption = ""
    Label13.Caption = ""
    Label12.Caption = ""

    ' ?? Product labels clear
    Label27.Caption = ""
    Label29.Caption = ""
    Label30.Caption = ""
    Label31.Caption = ""

    ' ?? Product fields clear
    Text5.Text = ""   ' Price
    Text6.Text = ""   ' Qty
    Text12.Text = ""  ' GST
    Text8.Text = ""   ' Total

    ' ?? Payment fields clear
    Text7.Text = ""   ' Advance
    Text10.Text = ""  ' Dues
    Text9.Text = ""   ' Net Amount

    Text13.Text = ""  ' Total GST
    Text14.Text = ""  ' Total Amount

    ' ?? List clear (all products)
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    List7.Clear
    List8.Clear
    List9.Clear

    ' ?? Date reset
    DTPicker1.Value = Date
    DTPicker2.Value = Date

    MsgBox "Form Cleared!", vbInformation

End Sub

Private Sub Command7_Click()
    ' ?? Validation
    If Combo3.Text = "" Then
        MsgBox "Select Product"
        Exit Sub
    End If

    If Text5.Text = "" Or Text6.Text = "" Or Text12.Text = "" Or Text8.Text = "" Then
        MsgBox "Enter complete details"
        Exit Sub
    End If

    ' ?? Add to list
    List1.AddItem Combo3.Text          ' Product ID
    List2.AddItem Label27.Caption      ' Name
    List3.AddItem Label29.Caption      ' Category
    List4.AddItem Label30.Caption      ' Brand
    List5.AddItem Label31.Caption      ' Type
    List6.AddItem Text5.Text           ' Price
    List7.AddItem Text6.Text           ' Qty
    List8.AddItem Text12.Text          ' GST %
    List9.AddItem Text8.Text           ' Total

    ' ?? Calculate totals
    Call CalculateGrandTotal

    ' ?? Clear product fields
    Combo3.ListIndex = -1
    Label27.Caption = ""
    Label29.Caption = ""
    Label30.Caption = ""
    Label31.Caption = ""

    Text5.Text = ""
    Text6.Text = ""
    Text12.Text = ""
    Text8.Text = ""

End Sub


Private Sub Command8_Click()

    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim SQL As String

    Set rs = New ADODB.Recordset

    Call CONN

    ' ?? Check selection
    If Combo5.Text = "" Then
        MsgBox "Select Order ID"
        Exit Sub
    End If

    ' =========================
    ' ?? MASTER LOAD
    ' =========================

    SQL = "SELECT * FROM ORDER_MASTER WHERE ORD_ID='" & Combo5.Text & "'"

    rs.Open SQL, C, adOpenKeyset, adLockReadOnly

    If rs.EOF Then
        MsgBox "No Record Found"
        Exit Sub
    End If

    ' ?? Fill Master
    Label10.Caption = rs!ORD_ID
    DTPicker1.Value = rs!ORD_DATE
    DTPicker2.Value = rs!DEL_DATE
    Combo1.Text = rs!SUP_ID
    Combo2.Text = rs!PAY_MODE

    Text14.Text = rs!NET_AMOUNT
    Text7.Text = rs!CURR_ADV
    Text10.Text = rs!dues
    Text9.Text = rs!NET_AMOUNT

    rs.Close

    ' =========================
    ' ?? CLEAR LIST
    ' =========================

    List1.Clear: List2.Clear: List3.Clear
    List4.Clear: List5.Clear: List6.Clear
    List7.Clear: List8.Clear: List9.Clear

    ' =========================
    ' ?? DETAILS LOAD
    ' =========================

    SQL = "SELECT * FROM ORDER_DETAILS WHERE ORD_ID='" & Combo5.Text & "'"

    rs.Open SQL, C, adOpenKeyset, adLockReadOnly

    While Not rs.EOF

        List1.AddItem rs!PR_ID

        ' ?? Product details
        Set rs2 = New ADODB.Recordset
        rs2.Open "SELECT PR_NM, CTG, BR_NM, TYPE FROM PRODUCT WHERE PR_ID='" & rs!PR_ID & "'", C

        If Not rs2.EOF Then
            List2.AddItem rs2!PR_NM
            List3.AddItem rs2!CTG
            List4.AddItem rs2!BR_NM
            List5.AddItem rs2!Type
        End If

        rs2.Close

        List6.AddItem rs!price
        List7.AddItem rs!qty
        List8.AddItem rs!GST
        List9.AddItem rs!total

        rs.MoveNext
    Wend

    rs.Close
    Call CalculateGrandTotal


    MsgBox "Order Loaded Successfully", vbInformation

End Sub

Private Sub Form_Load()

    Call CONN
    Combo2.AddItem "CASH"
    Combo2.AddItem "CARD"

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Combo1.Clear

    rs.Open "SELECT DISTINCT SUP_ID FROM QUOTATION", C, adOpenKeyset, adLockReadOnly

    While Not rs.EOF
        Combo1.AddItem rs.Fields("SUP_ID").Value
        rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing


    Set rs = New ADODB.Recordset

    Combo5.Clear

    rs.Open "SELECT ORD_ID FROM ORDER_MASTER", C, adOpenKeyset, adLockReadOnly

    While Not rs.EOF
        Combo5.AddItem rs!ORD_ID
        rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing

End Sub
Private Sub CalculateGrandTotal()

    Dim i As Integer
    Dim totalAmt As Double
    Dim totalGST As Double

    Dim price As Double
    Dim qty As Double
    Dim gstPer As Double
    Dim gstAmt As Double

    totalAmt = 0
    totalGST = 0

    For i = 0 To List1.ListCount - 1

        price = Val(List6.List(i))
        qty = Val(List7.List(i))
        gstPer = Val(List8.List(i))

        gstAmt = (price * qty * gstPer) / 100

        totalGST = totalGST + gstAmt
        totalAmt = totalAmt + Val(List9.List(i))

    Next i

    Text13.Text = Format(totalGST, "0.00")   ' Total GST
    Text14.Text = Format(totalAmt, "0.00")   ' Total Amount

End Sub
Private Sub CalculateTotal()

    Dim price As Double
    Dim qty As Double
    Dim gstPer As Double
    Dim gstAmt As Double
    Dim total As Double

    ' ?? Check empty
    If Text5.Text = "" Or Text6.Text = "" Or Text12.Text = "" Then Exit Sub

    price = Val(Text5.Text)
    qty = Val(Text6.Text)
    gstPer = Val(Text12.Text)

    gstAmt = (price * qty * gstPer) / 100
    total = (price * qty) + gstAmt

    Text8.Text = Format(total, "0.00")

End Sub

Private Sub Text12_Change()
    Call CalculateTotal
End Sub

Private Sub Text6_Change()
    Call CalculateTotal
End Sub

Private Sub Text7_Change()
    Call CalculatePayment
End Sub


Private Sub CalculatePayment()

    Dim totalAmt As Double
    Dim adv As Double
    Dim dues As Double

    If Text14.Text = "" Then Exit Sub

    totalAmt = Val(Text14.Text)
    adv = Val(Text7.Text)

    dues = totalAmt - adv

    If dues < 0 Then
        MsgBox "Advance cannot be greater than Total Amount"
        Text7.Text = ""
        Exit Sub
    End If

    Text10.Text = Format(dues, "0.00")   ' Dues
    Text9.Text = Format(totalAmt, "0.00") ' Net Amount

End Sub
