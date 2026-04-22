VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00FFC0C0&
   Caption         =   "GENERAL STORE MANAGEMENT SYSTEM"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12660
   LinkTopic       =   "Form9"
   ScaleHeight     =   8775
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   9480
      TabIndex        =   27
      Top             =   2880
      Width           =   2655
      Begin VB.CommandButton Command8 
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2040
         Width           =   1935
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
         Left            =   240
         TabIndex        =   30
         Text            =   "Combo3"
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By ID"
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
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
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
      Height          =   285
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6120
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   9840
      TabIndex        =   23
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      Format          =   105119745
      CurrentDate     =   46111
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2400
      TabIndex        =   21
      Text            =   "Combo2"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4680
      TabIndex        =   18
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   6960
      TabIndex        =   17
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   4680
      TabIndex        =   15
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   2400
      TabIndex        =   14
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
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
      Height          =   285
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Height          =   285
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   285
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      Format          =   104529921
      CurrentDate     =   46109
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
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   5160
      TabIndex        =   25
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   495
      Left            =   2400
      TabIndex        =   22
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID:"
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
      TabIndex        =   20
      Top             =   1680
      Width           =   1665
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6960
      TabIndex        =   19
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pack of "
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
      Left            =   5160
      TabIndex        =   16
      Top             =   2640
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valid Upto:"
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
      Left            =   8040
      TabIndex        =   12
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   7560
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID:"
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
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quotation ID:"
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
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT QUOTATION"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   3795
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


   Private Sub Combo1_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    rs.Open "SELECT PR_NM, PRICE FROM PRODUCT WHERE PR_ID='" & Trim(Combo1.Text) & "'", C

    If Not rs.EOF Then
        
        Label5.Caption = rs!PR_NM          ' Product Name
        
        Label13.Caption = rs!price         ' ?? Price store
        
        Label10.Caption = ""               ' Rate blank
        Text2.Text = ""                    ' Pack clear

    End If

    rs.Close
    Set rs = Nothing
End Sub

   



Private Sub Command1_Click()  ' Add New Button

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    rs.Open "SELECT NVL(MAX(TO_NUMBER(SUBSTR(Q_ID,3))),0)+1 FROM QUOTATION", C

    Label12.Caption = "QU" & Format(rs.Fields(0).Value, "000")

    rs.Close
    Set rs = Nothing

    ' Clear ???? fields
    Combo2.Enabled = True
    Combo2.Text = ""
    Combo1.Text = ""
    Label5.Caption = ""
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    Text2.Text = ""
    Label10.Caption = ""

    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear

End Sub
 

Private Sub Command2_Click()
    Dim ans As Integer

    If Label12.Caption = "" Then
        MsgBox "No record selected"
        Exit Sub
    End If

    ' Confirmation
    ans = MsgBox("Are you sure you want to delete?", vbYesNo + vbQuestion)

    If ans = vbYes Then

        ' First delete child table
        C.Execute "DELETE FROM Q_PROD_DET WHERE Q_ID='" & Label12.Caption & "'"

        ' Then delete main table
        C.Execute "DELETE FROM QUOTATION WHERE Q_ID='" & Label12.Caption & "'"

        MsgBox "Record Deleted Successfully"

        ' Clear form
        Label12.Caption = ""
        Combo2.Text = ""
        Combo2.Enabled = True

        Combo1.Text = ""
        Label5.Caption = ""
        Text2.Text = ""
        Label10.Caption = ""

        List1.Clear
        List2.Clear
        List3.Clear
        List4.Clear

    End If

End Sub

Private Sub Command3_Click()  ' Save Button

    Dim rs As ADODB.Recordset
    Dim SQL As String
    Dim i As Integer

    Set rs = New ADODB.Recordset

    Call CONN

    ' =========================
    ' ?? CHECK MASTER EXISTS
    ' =========================
    rs.Open "SELECT * FROM QUOTATION WHERE Q_ID='" & Label12.Caption & "'", C, adOpenKeyset, adLockReadOnly

    If rs.EOF Then
        ' ?? Insert only once
        SQL = "INSERT INTO QUOTATION VALUES('" & Label12.Caption & "'," & _
              "TO_DATE('" & DTPicker1.Value & "','DD-MM-YYYY')," & _
              "TO_DATE('" & DTPicker2.Value & "','DD-MM-YYYY')," & _
              "'" & Combo2.Text & "')"

        C.Execute SQL
    End If

    rs.Close

    ' =========================
    ' ?? INSERT DETAILS
    ' =========================

    For i = 0 To List1.ListCount - 1

        ' ?? Duplicate check
        rs.Open "SELECT * FROM Q_PROD_DET WHERE Q_ID='" & Label12.Caption & "' AND PR_ID='" & List1.List(i) & "'", C, adOpenKeyset, adLockReadOnly

        If rs.EOF Then
            SQL = "INSERT INTO Q_PROD_DET VALUES('" & Label12.Caption & "','" & _
                  List1.List(i) & "'," & _
                  Val(List3.List(i)) & "," & _
                  Val(List4.List(i)) & ")"

            C.Execute SQL
        End If

        rs.Close

    Next i

    MsgBox "Quotation Saved Successfully!", vbInformation

End Sub


Private Sub Command4_Click()
    ' Validation
    If Combo1.Text = "" Or Text2.Text = "" Or Label10.Caption = "" Then
        MsgBox "Fill all product details"
        Exit Sub
    End If

    ' Add to list
    List1.AddItem Combo1.Text
    List2.AddItem Label5.Caption
    List3.AddItem Text2.Text
    List4.AddItem Label10.Caption

    ' Supplier lock
    Combo2.Enabled = False

    ' Clear product fields
    Combo1.Text = ""
    Label5.Caption = ""
    Text2.Text = ""
    Label10.Caption = ""

End Sub

Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Command6_Click()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim qid As String
    qid = Combo3.Text

    If qid = "" Then
        MsgBox "Select Quotation ID"
        Exit Sub
    End If

    If C.State = 0 Then C.Open

    '==============================
    '?? 1. TOP DATA (quotation table)
    '==============================
    
    rs.Open "SELECT * FROM quotation WHERE Q_ID='" & qid & "'", C

    If Not rs.EOF Then
        
         Label12.Caption = rs!Q_ID      '?? textbox name apna check karo
        Combo2.Text = rs!SUP_ID               '?? supplier
        
        DTPicker1.Value = rs!Q_DATE           '?? date
        DTPicker2.Value = rs!VALID_UPTO       '?? valid upto
        
    End If

    rs.Close

    '==============================
    '?? 2. LIST DATA (products)
    '==============================

    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear

    rs.Open "SELECT PRODUCT.PR_ID, PRODUCT.PR_NM, q_prod_det.PACK, q_prod_det.RATE FROM PRODUCT, q_prod_det WHERE PRODUCT.PR_ID = q_prod_det.PR_ID AND q_prod_det.Q_ID='" & qid & "'", C

    While Not rs.EOF
        
        List1.AddItem rs!PR_ID
        List2.AddItem rs!PR_NM
        List3.AddItem rs!PACK
        List4.AddItem rs!Rate
        
        rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing

End Sub


                                                                                                                                                                                                                                                            

Private Sub Form_Load()
Call CONN
Dim rs As ADODB.Recordset
Set rs = C.Execute("SELECT PR_ID FROM PRODUCT")

While Not rs.EOF
    Combo1.AddItem rs("PR_ID")
    rs.MoveNext
Wend

rs.Close
Set rs = C.Execute("SELECT SUP_ID FROM SUP_DET")

While Not rs.EOF
    Combo2.AddItem rs("SUP_ID")
    rs.MoveNext
Wend

rs.Close
Set rs = C.Execute("SELECT Q_ID FROM QUOTATION")

Combo3.Clear

While Not rs.EOF
    Combo3.AddItem rs("Q_ID")
    rs.MoveNext
Wend

rs.Close

    Call LoadAllData

End Sub


Private Sub LoadAllData()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    '?? SAME connection use ???
    If C.State = 0 Then C.Open

    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear

    rs.CursorLocation = adUseClient

    rs.Open "SELECT PRODUCT.PR_ID, PRODUCT.PR_NM, q_prod_det.PACK, q_prod_det.RATE FROM PRODUCT, q_prod_det WHERE PRODUCT.PR_ID = q_prod_det.PR_ID", C

    While Not rs.EOF
        
        List1.AddItem rs!PR_ID
        List2.AddItem rs!PR_NM
        List3.AddItem rs!PACK
        List4.AddItem rs!Rate
        
        rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing

End Sub










   Private Sub Text2_Change()
    If IsNumeric(Text2.Text) And IsNumeric(Label13.Caption) Then
        
        Label10.Caption = Val(Text2.Text) * Val(Label13.Caption)
        
    Else
        Label10.Caption = ""
        
    End If
End Sub

