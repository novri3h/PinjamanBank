Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RS As ADODB.Recordset
Public RSNasabah As ADODB.Recordset
Public RSKasir As ADODB.Recordset
Public RSPinjam As ADODB.Recordset
Public RSDetailPjm As ADODB.Recordset
Public RSBayar As ADODB.Recordset

Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RS = New ADODB.Recordset
Set RSNasabah = New ADODB.Recordset
Set RSPinjam = New ADODB.Recordset
Set RSKasir = New ADODB.Recordset
Set RSDetailPjm = New ADODB.Recordset
Set RSBayar = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBKeuangan.mdb"
End Sub


