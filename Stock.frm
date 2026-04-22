VERSION 5.00
Begin VB.Form Stock 
   Caption         =   "STOCK"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10740
   LinkTopic       =   "Form10"
   ScaleHeight     =   6885
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "pri nt selective"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "current stock"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "unit"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Brand"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Category"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "product name"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Product Id"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim rs As New ADODB.Recordset
    
    If Combo1.Text = "" Then
        MsgBox "Select Product ID"
        Exit Sub
    End If
    
    Call CONN

    ' ?? Product details
    rs.Open "SELECT PR_NM, CTG, BR_NM, UNIT FROM PRODUCT WHERE PR_ID='" & Combo1.Text & "'", C, adOpenKeyset, adLockReadOnly
    
    If Not rs.EOF Then
        Label7.Caption = rs!PR_NM
        Label8.Caption = rs!CTG
        Label9.Caption = rs!BR_NM
        Label10.Caption = rs!UNIT
    End If
    
    rs.Close

    ' ?? Stock details
    rs.Open "SELECT CURRENT_QTY FROM STOCK WHERE PR_ID='" & Combo1.Text & "'", C, adOpenKeyset, adLockReadOnly
    
    If Not rs.EOF Then
        Label11.Caption = rs!CURRENT_QTY
    Else
        Label11.Caption = "0"
    End If
    
    rs.Close

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim rs As New ADODB.Recordset
    
    Call CONN
    
    Combo1.Clear
    
    rs.Open "SELECT DISTINCT PR_ID FROM ORDER_DETAILS", C, adOpenKeyset, adLockReadOnly
    
    While Not rs.EOF
        Combo1.AddItem rs!PR_ID
        rs.MoveNext
    Wend
    
    rs.Close

End Sub
