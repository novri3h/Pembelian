VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pembelian 
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
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
   ScaleHeight     =   5985
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Bantuan"
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   350
      Left            =   3240
      TabIndex        =   28
      Top             =   2160
      Width           =   4000
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   3240
      TabIndex        =   27
      Top             =   1800
      Width           =   4000
   End
   Begin VB.TextBox TxtDibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   8160
      TabIndex        =   4
      Top             =   5040
      Width           =   1250
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   960
      TabIndex        =   8
      Top             =   4680
      Width           =   850
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   1800
      TabIndex        =   7
      Top             =   4680
      Width           =   850
   End
   Begin VB.ListBox List1 
      Height          =   1410
      Left            =   7440
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4200
      Top             =   4680
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   5280
      TabIndex        =   1
      Top             =   720
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   4000
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   4000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1845
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3254
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Kode"
         Caption         =   "Kode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Nama"
         Caption         =   "Nama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Harga"
         Caption         =   "Harga"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   405
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Transaksi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Person"
      Height          =   345
      Left            =   2400
      TabIndex        =   30
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   2400
      TabIndex        =   29
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaksi Pembelian Barang"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   9735
   End
   Begin VB.Label LblJam 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   960
      TabIndex        =   25
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   960
      TabIndex        =   24
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   345
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label LblFaktur 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   960
      TabIndex        =   21
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faktur"
      Height          =   345
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   7320
      TabIndex        =   19
      Top             =   4680
      Width           =   795
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8160
      TabIndex        =   18
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   7320
      TabIndex        =   17
      Top             =   5040
      Width           =   795
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   7320
      TabIndex        =   16
      Top             =   5400
      Width           =   795
   End
   Begin VB.Label LblKembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   8160
      TabIndex        =   15
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   2880
      TabIndex        =   14
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label LblItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3480
      TabIndex        =   13
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " KodePms"
      Height          =   345
      Left            =   2400
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   2400
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   2400
      TabIndex        =   10
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "Pembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
MsgBox "Cara transaksi" & vbNewLine & _
"Jika konsumen baru, maka lengkapi data terlebih dahulu" & vbNewLine & _
"Jika konsumen sudah terdaftar, silakan pilih kode konsumen dalam combo" & vbNewLine & _
"Kode barang dapat diketik di kolom kode" & vbNewLine & _
"atau pilih nama barang dalam list, lalu tekan enter" & vbNewLine & _
"Selanjutnya silakan isi jumlah barang di kolom jumlah"
End Sub

Private Sub form_activate()
'hubungkan objek ke database
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBeli.mdb"
'hubungkan objek ke tabel transaksi
DT.RecordSource = "Transaksi"
'hubungkan grid ke objek
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
'buka database
Call BukaDB
'buka tabel barang
RSBarang.Open "Barang", Conn
'tampilkan nama dan kode barang di list
List1.Clear
Do Until RSBarang.EOF
    List1.AddItem RSBarang!NamaBrg & Space(50) & RSBarang!KodeBrg
    RSBarang.MoveNext
Loop
'buka tabel pemasok
RSPemasok.Open "pemasok", Conn

Combo1.Clear
'tampilkan kod epemasok di combo
Do Until RSPemasok.EOF
    Combo1.AddItem RSPemasok!kodepms & Space(5) & RSPemasok!namapms
    RSPemasok.MoveNext
Loop
'panggil nomor faktur otomatis
Call Auto
'kosongkan tabel transaksi
Call Tabel_Kosong
'simpan kursor di baris pertama tabel transaksi
DT.Recordset.MoveFirst
LblTanggal = Format(Date, "dd-mm-yyyy")
'matikan dulu command simpan
CmdSimpan.Enabled = False
Call Gelap
Call Kosongpms
End Sub

Private Sub Form_Load()
    'kode kasir diambil dari login
    'nama kasir diambil dari login
    LblKodeKsr = Login.TxtKodeKsr
    LblNamaKsr = Login.TxtNamaKsr
    'klik dulu command batal untuk mengosongkan tabel transaksi
    CmdBatal_Click
    Text1.MaxLength = 5
    Text2.MaxLength = 30
    Text3.MaxLength = 30
    Text4.MaxLength = 15
    Text4.MaxLength = 25
End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "pilih kode pemasok...!"
        Combo1.SetFocus
    Else
        DataGrid1.SetFocus
    End If
End If
If Keyascii = 27 Then
    Kosongpms
    Text1.SetFocus
End If
End Sub

Private Sub Combo1_Click()
    Call BukaDB
    'cari data pemasok
    RSPemasok.Open "Select * from Pemasok where Kodepms='" & Left(Combo1, 6) & "'", Conn
    'jika ditemukan tampilkan datanya
    If Not RSPemasok.EOF Then
        Text1 = RSPemasok!kodepms
        Text2 = RSPemasok!namapms
        Text3 = RSPemasok!alamatpms
        Text4 = RSPemasok!teleponpms
        Text5 = RSPemasok!personpms
    End If
    Conn.Close
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Text1 = "" Then
        MsgBox "Isi dulu kode pemasok, contoh PMS01"
        Text1.SetFocus
        Exit Sub
    Else
        Call BukaDB
        RSPemasok.Open "Select * from Pemasok where Kodepms='" & Text1 & "'", Conn
        If RSPemasok.EOF Then
            MsgBox ("Ini Pemasok baru, isi datanya lalu simpan")
            Call Terang
            Text2.SetFocus
        Else
            Text2 = RSPemasok!namapms
            Text3 = RSPemasok!alamatpms
            Text4 = RSPemasok!teleponpms
            Text5 = RSPemasok!personpms
            DataGrid1.SetFocus
            DataGrid1.Col = 1
        End If
        Conn.Close
    End If
End If
End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then Text5.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Text5_keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then DataGrid1.SetFocus
End Sub

Private Sub Timer1_Timer()
    LblJam = Time$
End Sub

Sub Gelap()
Text1.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
End Sub

Sub Terang()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Sub

Sub Kosongpms()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Combo1 = ""
End Sub

'mencari nomor otomatis
Private Sub Auto()
Call BukaDB
'baca tabelpembelian yang fakturnya paling akhir
RSPembelian.Open "select * from Pembelian Where Faktur In(Select Max(Faktur)From Pembelian)Order By Faktur Desc", Conn
RSPembelian.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPembelian
        'jika tidak ditemukan maka...
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "0001"
            'no fakturnya adalah YYMMDD0001
            LblFaktur = Urutan
        Else
            'jika ganti hari maka... nomor fakturnya
            If Left(!Faktur, 6) <> Format(Date, "yymmdd") Then
                'YYMMDD0001
                Urutan = Format(Date, "yymmdd") + "0001"
            Else
                'jika harinya sama maka... YYMMDD0001+1
                Hitung = (!Faktur) + 1
                Urutan = Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        LblFaktur = Urutan
    End With
End Sub

'hapus dulu isi tbl transaksi
Function Tabel_Kosong()
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        DT.Recordset.Delete
        DT.Recordset.MoveNext
    Loop
    'lalu tambahkan satu baris kosong
    For i = 1 To 1
        DT.Recordset.AddNew
        DT.Recordset!Nomor = i
        DT.Recordset.Update
    Next i
    DataGrid1.Col = 1
End Function

'jumlah data + 1
Function Tambah_Baris()
    For i = DT.Recordset.RecordCount To DT.Recordset.RecordCount
        DT.Recordset.AddNew
        DT.Recordset!Nomor = i + 1
        DT.Recordset.Update
    Next i
End Function

Private Sub DataGrid1_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If DataGrid1.Col = 3 Then
    'kolom 3 dan 4 hanya dapat diisi angka
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
ElseIf DataGrid1.Col = 4 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
End If
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If DataGrid1.Col = 1 Then
        'kode barang harus 6 digit
        If Len(DT.Recordset!Kode) < 6 Then
            MsgBox "Kode Harus 6 digit"
            DataGrid1.Col = 1
            Exit Sub
        End If
    
        Call BukaDB
        'cari barang yg kodenya diketik di grid
        RSBarang.Open "Select * from Barang where Kodebrg='" & DT.Recordset!Kode & "'", Conn
        'jika tidak ada munculkan pesan
        If RSBarang.EOF Then
            MsgBox ("Ini Kode Barang baru, isi data dengan lengkap")
            DT.Recordset!Kode = DT.Recordset!Kode
            'isi nama barang (karena ini barang baru)
            DataGrid1.Col = 2
            DataGrid1.Refresh
            Exit Sub
        Else
            'jika ditemukan tampilkan nama,harga dst...
            DT.Recordset!Kode = RSBarang!KodeBrg
            DT.Recordset!nama = RSBarang!NamaBrg
            DT.Recordset!Harga = RSBarang!HargaBeli
            DataGrid1.Col = 4
            DataGrid1.Refresh
            Exit Sub
        End If
    End If
    
    'isi nama barang jika barang baru
    If DataGrid1.Col = 2 Then
        DT.Recordset!nama = DT.Recordset!nama
        DT.Recordset.Update
        DataGrid1.Col = 3
        DataGrid1.Refresh
        Exit Sub
    End If
    
    'isi harga barang jika barang baru
    If DataGrid1.Col = 3 Then
        DT.Recordset!Harga = DT.Recordset!Harga
        DT.Recordset.Update
        DataGrid1.Col = 4
        DataGrid1.Refresh
        Exit Sub
    End If
    
    'isi jumlah barang jika barang baru
    If DataGrid1.Col = 4 Then
        DT.Recordset!jumlah = DT.Recordset!jumlah
        'total dihasilkan dari harga x jumlah
        DT.Recordset!Total = DT.Recordset!Harga * DT.Recordset!jumlah
        DT.Recordset.Update
        Call Tambah_Baris
        DT.Recordset.MoveNext
        DataGrid1.Col = 1
        DT.Recordset.MoveLast
        'tampilkan total harga dan total item
        LblTotal = Format(TotalHarga, "#,###,###")
        LblItem = Format(TotalItem, "#,###,###")
    End If
End Sub

Private Sub Bersihkan()
    LblItem = ""
    LblTotal = ""
    TxtDibayar = ""
    LblKembali = ""
    Call Kosongpms
End Sub

Private Sub TxtDibayar_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        'pembayaran tidak boleh kosong atau lebih kecil
        If TxtDibayar = "" Or Val(TxtDibayar) < (LblTotal) Then
            MsgBox "Jumlah Pembayaran Kurang"
            TxtDibayar.SetFocus
        Else
            TxtDibayar = Format(TxtDibayar, "###,###,###")
            If TxtDibayar = LblTotal Then
                LblKembali = TxtDibayar - LblTotal
            Else
                LblKembali = Format(TxtDibayar - LblTotal, "###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub CmdSimpan_Keypress(Keyascii As Integer)
    If Keyascii = 27 Then
        CmdSimpan.Enabled = False
        TxtDibayar = ""
        TxtDibayar.SetFocus
    End If
End Sub

Private Sub CmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or TxtDibayar = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
Else
    If LblItem = "" Then
        MsgBox "tidak transaksi pembelian"
        Exit Sub
    End If
End If

    Call BukaDB
    'simpan data pemasok jika ini data baru
    RSPemasok.Open "select * from pemasok where kodePMS='" & Text1 & "'", Conn
    If RSPemasok.EOF Then
        Dim TambahPemasok As String
        TambahPemasok = "Insert Into Pemasok(Kodepms,Namapms,AlamatPms,TeleponPms,PersonPms)" & _
        "values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')"
        Conn.Execute (TambahPemasok)
    End If
    
    'simpan transaksi ke tbl pembelian
    Dim SQLTambahJual As String
    SQLTambahJual = "Insert Into Pembelian(Faktur,Tanggal,Jam,JmlItem,JmlTotal,Dibayar,Kembali,KodeKsr,KodePms)" & _
    "values('" & LblFaktur & "','" & LblTanggal & "','" & LblJam & "','" & LblItem & "','" & LblTotal & "','" & TxtDibayar & "','" & LblKembali & "','" & Menu.STBar.Panels(1).Text & "','" & Text1 & "')"
    Conn.Execute (SQLTambahJual)
    
    'simpan data transaksi ke tabel detailbeli
    'jika ada kode yang sama maka jumlahnya akan disatukan
    'RSTransaksi.Open "select kode as KodeBrg,sum(Jumlah) as JumlahBrg from Transaksi group by kode", Conn
    'RSTransaksi.MoveFirst
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        If DT.Recordset!Kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into DetailBeli(Faktur,Kodebrg,JmlBeli,subtotal) " & _
            "values ('" & LblFaktur & "','" & DT.Recordset!Kode & "','" & DT.Recordset!jumlah & "','" & DT.Recordset!Total & "')"
            Conn.Execute (SQLTambahDetail)
        End If
    DT.Recordset.MoveNext
    Loop
        
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        If DT.Recordset!Kode <> vbNullString Then
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kodebrg='" & DT.Recordset!Kode & "'", Conn
            If Not RSBarang.EOF Then
                'tambah barang jika kodenya ditemukan
                Dim TambahBarang1 As String
                TambahBarang1 = "update barang set jumlahbrg='" & RSBarang!JumlahBrg + DT.Recordset!jumlah & "' where kodebrg='" & DT.Recordset!Kode & "'"
                Conn.Execute (TambahBarang1)
            Else
                'input data barang jika kodenya baru
                Dim TambahBarang2 As String
                TambahBarang2 = "Insert Into Barang(Kodebrg,NamaBrg,HargaBeli,HargaJual,JumlahBrg)" & _
                "values('" & DT.Recordset!Kode & "','" & DT.Recordset!nama & "','" & DT.Recordset!Harga & "','" & DT.Recordset!Harga * 1.5 & "','" & DT.Recordset!jumlah & "')"
                Conn.Execute (TambahBarang2)
            End If
        End If
    DT.Recordset.MoveNext
    Loop
    
    Bersihkan
    form_activate
    Text1.SetFocus
    'panggil prosedur pencetakan
    Call Cetak
End Sub

Private Sub CmdBatal_Click()
    TxtDibayar = ""
    LblTotal = ""
    LblItem = ""
    LblKembali = ""
    form_activate
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

'mencari total harga
Function TotalHarga()
    Set TTlHarga = New ADODB.Recordset
    TTlHarga.Open "select sum(Total) as JumTotal from Transaksi", Conn
    TotalHarga = TTlHarga!JumTotal
End Function

'mencari total item
Function TotalItem()
    Set TTlItem = New ADODB.Recordset
    TTlItem.Open "select sum(Jumlah) as JumItem from Transaksi", Conn
    TotalItem = TTlItem!Jumitem
End Function


Function Cetak()
Call BukaDB
'cari faktur terakhir
RSPembelian.Open "select * from Pembelian Where Faktur In(Select Max(Faktur)From Pembelian)Order By Faktur Desc", Conn
Tampilkan.Show

Dim JmlHarga, JmlBeli, JmlHasil As Double
Dim MGrs As String
Tampilkan.Font = "Courier New"
Tampilkan.Print
Tampilkan.Print
'buka tabel kasir dan pemasok
RSKasir.Open "select * From Kasir where KodeKsr= '" & RSPembelian!kodeksr & "'", Conn
RSPemasok.Open "select * From pemasok where Kodepms= '" & RSPembelian!kodepms & "'", Conn

'cetak data ke layar
Tampilkan.Print Tab(5); "Faktur     :   "; RSPembelian!Faktur
Tampilkan.Print Tab(5); "Tanggal    :   "; Format(RSPembelian!Tanggal, "DD-MMMM-YYYY")
Tampilkan.Print Tab(5); "Jam        :   "; Format(RSPembelian!Jam, "HH:MM:SS")
Tampilkan.Print Tab(5); "Kasir      :   "; RSKasir!NamaKsr
Tampilkan.Print Tab(5); "Pemasok    :   "; RSPemasok!namapms
Tampilkan.Print Tab(5); "Telepon    :   "; RSPemasok!teleponpms

MGrs = String$(33, "-")
Tampilkan.Print Tab(5); MGrs

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
    'tampilkan berulang-ulang kode,nama,harga,jumlah dan total
    Tampilkan.Print Tab(5); No; Space(2); RSBarang!NamaBrg
    Tampilkan.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    Tampilkan.Print Tab(15); Format(Harga, "###,###,###");
    Tampilkan.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailBeli.MoveNext
Loop

'tampilkan total harga
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Total      :";
Tampilkan.Print Tab(25); RKanan(RSPembelian!Jmltotal, "###,###,###");
Tampilkan.Print Tab(5); "Dibayar    :";
'tampilkan dibayar
Tampilkan.Print Tab(25); RKanan(RSPembelian!Dibayar, "###,###,###");
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print Tab(5); "Kembali    :";
'tampilkan kembalian
If RSPembelian!Dibayar = RSPembelian!Jmltotal Then
    Tampilkan.Print Tab(34); RSPembelian!Dibayar - RSPembelian!Jmltotal
Else
    Tampilkan.Print Tab(25); RKanan(RSPembelian!Dibayar - RSPembelian!Jmltotal, "###,###,###");
End If
Tampilkan.Print Tab(5); MGrs
Tampilkan.Print
Tampilkan.Print
Tampilkan.Print
Conn.Close
End Function

'ratakan angka di kanan
Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

'kode barang dapat diambil dari list
Private Sub List1_keyPress(Keyascii As Integer)
    'jika menekan enter setelah memilih data barang, maka...
    If Keyascii = 13 Then
        'jika isi data di kolom kode <> kode barang...
        If DataGrid1.SelText <> Right(List1, 6) Then
            'maka ganti (tiban) dengan kode barang dari list
            DataGrid1.SelText = Right(List1, 6)
            DT.Recordset.Update
            Call BukaDB
            'cari data barang yang kdoenya dipilih di list
            RSBarang.Open "Select * from Barang where KodeBrg='" & Right(List1, 6) & "'", Conn
            RSBarang.Requery
            'jika ditemukan
            If Not RSBarang.EOF Then
                'tampilkan data datanya
                DT.Recordset!Kode = RSBarang!KodeBrg
                DT.Recordset!nama = RSBarang!NamaBrg
                DT.Recordset!Harga = RSBarang!HargaBeli
                DT.Recordset.Update
                DataGrid1.SetFocus
                'kursor pindah ke kolom 4
                DataGrid1.Col = 4
            End If
        End If
    End If
End Sub

