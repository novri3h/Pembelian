VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Rincian 
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   5475
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6720
      TabIndex        =   19
      Top             =   3600
      Width           =   1000
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   350
      Left            =   6240
      TabIndex        =   18
      Top             =   3600
      Width           =   500
   End
   Begin VB.TextBox Text8 
      Height          =   350
      Left            =   4920
      TabIndex        =   16
      Top             =   4320
      Width           =   1000
   End
   Begin VB.TextBox Text7 
      Height          =   350
      Left            =   4920
      TabIndex        =   13
      Top             =   3960
      Width           =   2800
   End
   Begin VB.TextBox Text6 
      Height          =   350
      Left            =   4920
      TabIndex        =   12
      Top             =   3600
      Width           =   1000
   End
   Begin VB.TextBox Text5 
      Height          =   350
      Left            =   4920
      TabIndex        =   6
      Top             =   4680
      Width           =   2800
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   960
      TabIndex        =   5
      Top             =   4680
      Width           =   3000
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   4320
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   960
      TabIndex        =   2
      Top             =   3600
      Width           =   1000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   3600
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Rincian.frx":0000
      Height          =   2655
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nama Barang"
         Caption         =   "Nama Barang"
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
         DataField       =   "Harga Beli"
         Caption         =   "Harga Beli"
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
      BeginProperty Column03 
         DataField       =   "Total"
         Caption         =   "Total"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rincian Transaksi Pembelian"
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
      TabIndex        =   20
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   4080
      TabIndex        =   17
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   4080
      TabIndex        =   15
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   4080
      TabIndex        =   14
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Person"
      Height          =   345
      Left            =   4080
      TabIndex        =   11
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " KodePms"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "Rincian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
'buka database
Call BukaDB
'bersihkan dulu list
List1.Clear
'cari nomor faktur di tabel pembelian
RSDetailBeli.Open "Select Distinct Faktur from detailbeli", Conn
'tampilkan di list
Do Until RSDetailBeli.EOF
    List1.AddItem RSDetailBeli!Faktur
    RSDetailBeli.MoveNext
Loop
Conn.Close
Call Gelap
End Sub

'ketika salah satu faktur dipilih, maka...
Private Sub list1_click()
'buka database
Call BukaDB
Conn.CursorLocation = adUseClient
'cari data pembelian yang fakturnya dipilih
RSPembelian.Open "select * from Pembelian where Faktur='" & List1.Text & "'", Conn
RSPembelian.Requery
'jika ditemukan tampilkan tanggalnya
If Not RSPembelian.EOF Then Text8 = RSPembelian!Tanggal
'buka tabel pemasok yang ada di tabel pembelian sesuai noor faktur
RSPemasok.Open "select * from pemasok where KodePms='" & RSPembelian!kodepms & "'", Conn
'jika ditemukan tampilkan data-datanya
If Not RSPemasok.EOF Then
    Text1 = RSPemasok!kodepms
    Text2 = RSPemasok!namapms
    Text3 = RSPemasok!alamatpms
    Text4 = RSPemasok!teleponpms
    Text5 = RSPemasok!personpms
End If
'buka tabel kasir yang kodenya disimpan di tabel pembelian berdasarkan nomor faktur
RSKasir.Open "select * from Kasir where KodeKsr='" & RSPembelian!kodeksr & "'", Conn
'jika ditemukan tampilkan kode dan nama kasir
If Not RSKasir.EOF Then
    Text6 = RSKasir!kodeksr
    Text7 = RSKasir!NamaKsr
End If

Conn.Close
'hubungkan objek adodc ke database
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBeli.mdb"
'tampilkan nama barang, harga beli, jumlah beli dan total di tabel pembelian,detail beli yang fakturnya dipilih dalam list
'Adodc1.RecordSource = "select NamaBrg as [Nama Barang], HargaBeli,JmlBeli as Jumlah, HargaBeli*JmlBeli as Total from Barang,detailBeli where DetailBeli.kodeBrg=Barang.kodeBrg and Faktur='" & List1.Text & "'"
Adodc1.RecordSource = "select namabrg as [Nama Barang], HargaBeli as [Harga Beli],JmlBeli as Jumlah, subtotal as Total from barang,detailBeli,pembelian  where pembelian.Faktur='" & List1.Text & "' and barang.kodebrg=detailbeli.kodebrg and pembelian.faktur=detailbeli.faktur"
Adodc1.Refresh
'hubungkan datagrid1 dengan objek adodc
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'tampilkan total dan item
Call Total
Call Item
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'mencari jumlah total item barang
Function Item()
Adodc1.Recordset.MoveFirst
jumlah = 0
Do While Not Adodc1.Recordset.EOF
    jumlah = jumlah + Adodc1.Recordset!jumlah
    Adodc1.Recordset.MoveNext
Loop
Text9 = jumlah
End Function

'mencari jumlah total harga beli
Function Total()
Adodc1.Recordset.MoveFirst
jumlah = 0
Do While Not Adodc1.Recordset.EOF
    jumlah = jumlah + Adodc1.Recordset!Total
    Adodc1.Recordset.MoveNext
Loop
Text10 = jumlah
End Function

Sub Gelap()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
End Sub
