VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8445
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   13500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   24111.35
   ScaleMode       =   0  'User
   ScaleWidth      =   11595.24
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   6435
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   11145
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   840
         TabIndex        =   2
         Top             =   3840
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Timer Timer1 
         Left            =   2400
         Top             =   3840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lavanya Kumar"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   4800
         Width           =   1320
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Developed By: Priyanshu Kumar Singh (Team Leader)"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   4560
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..... "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   3480
         Width           =   1110
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   360
         Picture         =   "frmSplash.frx":C311
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GENERAL STORE MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2640
         TabIndex        =   1
         Top             =   600
         Width           =   6570
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
Timer1.Interval = 70
Timer1.Enabled = True
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub



Private Sub Timer1_Timer()
X = X + 1
ProgressBar1.Value = X
If X >= 100 Then
Timer1.Enabled = False
MDIForm1.Show
Unload Me
End If
End Sub
