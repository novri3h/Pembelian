VERSION 5.00
Begin VB.Form Tampilkan 
   BackColor       =   &H80000009&
   Caption         =   "ESC = Tutup ** Enter = Cetak"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Tampilkan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Unload Me
ElseIf Keyascii = 13 Then
    Pesan = MsgBox("Printer sudah siap", vbYesNo)
    If Pesan = vbYes Then
        Call Cetak
    Else
        Unload Me
    End If
End If
End Sub

Function Cetak()
Call BukaDB
'cari faktur terakhir
RSPembelian.Open "select * from Pembelian Where Faktur In(Select Max(Faktur)From Pembelian)Order By Faktur Desc", Conn
Dim JmlHarga, JmlBeli, JmlHasil As Double
Dim MGrs As String
Printer.Font = "Courier New"
Printer.Print
Printer.Print
'buka tabel kasir dan pemasok
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSPembelian!kodeksr & "'", Conn
RSPemasok.Open "select * From pemasok where Kodepms= '" & RSPembelian!kodepms & "'", Conn

'cetak data ke printer
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print Tab(5); "Faktur     :   "; RSPembelian!Faktur
Printer.Print Tab(5); "Tanggal    :   "; Format(RSPembelian!Tanggal, "DD-MMMM-YYYY")
Printer.Print Tab(5); "Jam        :   "; Format(RSPembelian!Jam, "HH:MM:SS")
Printer.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
Printer.Print Tab(5); "Pemasok    :   "; RSPemasok!namapms
Printer.Print Tab(5); "Telepon    :   "; RSPemasok!teleponpms

MGrs = String$(33, "-")
Printer.Print Tab(5); MGrs

'cari data di tabel detailbeli yang fakturnya =di tbl pembelian
RSDetailBeli.Open "select * from DetailBeli Where Faktur='" & RSPembelian!Faktur & "'", Conn
RSDetailBeli.MoveFirst

No = 0
Do While Not RSDetailBeli.EOF
    No = No + 1
    
    Set RSBarang = New ADODB.Recordset
    'cari barang yang kodenya disimpan di tabel detailbeli
    RSBarang.Open "select * From Barang where Kodebrg= '" & RSDetailBeli!KodeBrg & "'", Conn
    RSBarang.Requery
    Harga = RSBarang!HargaBeli
    jumlah = RSDetailBeli!JmlBeli
    Hasil = Harga * jumlah
    'printer berulang-ulang kode,nama,harga,jumlah dan total
    Printer.Print Tab(5); No; Space(2); RSBarang!NamaBrg
    Printer.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    Printer.Print Tab(15); Format(Harga, "###,###,###");
    Printer.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailBeli.MoveNext
Loop

'printer total harga
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total      :";
Printer.Print Tab(25); RKanan(RSPembelian!Jmltotal, "###,###,###");
Printer.Print Tab(5); "Dibayar    :";
'printer dibayar
Printer.Print Tab(25); RKanan(RSPembelian!Dibayar, "###,###,###");
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Kembali    :";
'printer kembalian
If RSPembelian!Dibayar = RSPembelian!Jmltotal Then
    Printer.Print Tab(34); RSPembelian!Dibayar - RSPembelian!Jmltotal
Else
    Printer.Print Tab(25); RKanan(RSPembelian!Dibayar - RSPembelian!Jmltotal, "###,###,###");
End If
Printer.Print Tab(5); MGrs
Printer.Print
Printer.Print
Printer.Print
Printer.EndDoc
Conn.Close

End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

