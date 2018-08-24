Attribute VB_Name = "Conexão"
Public codcliente As Integer
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public path As String
Public Sub connectBD()
path = App.path & "\BaseDeDados.mdb"
db.Open "provider=microsoft.jet.oledb.4.0;data source=" & path
End Sub
Public Sub fechaBD()
rs.Close: Set rs = Nothing
db.Close: Set db = Nothing
End Sub

