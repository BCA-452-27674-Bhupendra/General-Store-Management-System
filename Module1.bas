Attribute VB_Name = "Module1"
Public C As ADODB.Connection
Public R As ADODB.Recordset
Public SQL As String
Public Function CONN()
Set C = New ADODB.Connection
C.Open "Provider=MSDAORA.1;User ID=PRJ2531B/PRJ2531B;Persist Security Info=False"
Set R = New ADODB.Recordset
End Function
