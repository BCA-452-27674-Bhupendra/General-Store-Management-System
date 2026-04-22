VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "GENERAL STORE MANAGEMENT SYSTEM"
   ClientHeight    =   9180
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   14085
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMASTER 
      Caption         =   "MASTER"
      Begin VB.Menu mnuEmployee 
         Caption         =   "Employee Detail"
      End
      Begin VB.Menu mnuSupplier 
         Caption         =   "Supplier Detail"
      End
      Begin VB.Menu mnuProduct 
         Caption         =   "Product DETAIL"
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer Detail"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "TRANSACTION"
      Begin VB.Menu mnuQuotation 
         Caption         =   "Product Quotation"
      End
      Begin VB.Menu mnuPurchase 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu mnuSalesDeatail 
         Caption         =   "Sales Details"
      End
      Begin VB.Menu mnuStock 
         Caption         =   "Stock"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCustomer_Click()
Form5.Show
End Sub

Private Sub mnuEmployee_Click()
Form2.Show
End Sub

Private Sub mnuProduct_Click()
Form4.Show
End Sub

Private Sub mnuPurchase_Click()
PURCHASE.Show
End Sub

Private Sub mnuQuotation_Click()
Form9.Show
End Sub

Private Sub mnuSupplier_Click()
Form3.Show
End Sub
