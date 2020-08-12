Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RSBarang As ADODB.Recordset
Public RSKasir As ADODB.Recordset
Public RSPembelian As ADODB.Recordset
Public RSDetailBeli As ADODB.Recordset
Public RSTransaksi As ADODB.Recordset
Public RSPemasok As ADODB.Recordset
Public obj As Form


Public Sub BukaDB()
Dim STR As String
Set Conn = New ADODB.Connection
Set RSBarang = New ADODB.Recordset
Set RSKasir = New ADODB.Recordset
Set RSPembelian = New ADODB.Recordset
Set RSDetailBeli = New ADODB.Recordset
Set RSTransaksi = New ADODB.Recordset
Set RSPemasok = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ADOBeli.mdb"
End Sub

