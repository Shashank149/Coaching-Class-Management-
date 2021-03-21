Attribute VB_Name = "Module1"
Public CN As New ADODB.Connection
Public res As New ADODB.Recordset


Public Sub connectDB()
CN.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\project cc\CC\Third\Final\db1.mdb;Persist Security Info=False")
End Sub
