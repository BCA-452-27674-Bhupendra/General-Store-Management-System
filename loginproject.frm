VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "General store management system"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   Picture         =   "loginproject.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   4080
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Text            =   "Type Your Password"
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   435
         Left            =   2280
         TabIndex        =   4
         Text            =   "Type Your Username"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Image Image2 
         Height          =   945
         Left            =   2040
         Picture         =   "loginproject.frx":F005
         Top             =   5160
         Width           =   2445
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   2040
         Picture         =   "loginproject.frx":FB6B
         Top             =   4320
         Width           =   2445
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         X1              =   2280
         X2              =   4800
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   2280
         X2              =   4800
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3360
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "logIn"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   2010
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text2.PasswordChar = "*"
End Sub

Private Sub Image1_Click()
    Dim user As String
    Dim pass As String

    user = LCase(Trim(Text1.Text))
    pass = UCase(Trim(Text2.Text))

    If user <> "prj2531b" Then
        MsgBox "Wrong Username", vbCritical
        Text1.SetFocus
        Exit Sub
    End If

    If pass <> "PRJ2531B" Then
        MsgBox "Wrong Password", vbCritical
        Text2.SetFocus
        Exit Sub
    End If
MsgBox "Login Successful"
    ' ?? Splash Form open karo
    frmSplash.Show
    Unload Me   ' login form band
End Sub


Private Sub Image2_Click()

    Dim result As Integer
    
    result = MsgBox("Do you want to exit?", vbYesNo + vbQuestion, "Confirm")

    If result = vbYes Then
        
        ' Exit form
        Unload Me
        
    Else
        
        ' Clear fields
        Text1.Text = ""
        Text2.Text = ""
        
        ' Cursor wapas username pe
        Text1.SetFocus
        
    End If

End Sub

Private Sub Text1_Click()
Text1.Text = ""
Text1.ForeColor = vbBlack
End Sub
Private Sub Text2_Click()
Text2.Text = ""
Text2.ForeColor = vbBlack
End Sub
