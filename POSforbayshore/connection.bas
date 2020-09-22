Attribute VB_Name = "connection"
Global con As New ADODB.connection
Global rst As New ADODB.Recordset
Global rst1 As New ADODB.Recordset

Public Sub connect()
Set con = Nothing
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\POSforbayshore\database.mdb;"
con.Open
End Sub
