VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PINJAMAN 
   Caption         =   "PINJAMAN"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   2000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1200
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3240
      Width           =   1200
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   2000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8281
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
         DataField       =   "No"
         Caption         =   "No"
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
         DataField       =   "CICILAN"
         Caption         =   "CICILAN"
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
         DataField       =   "BUNGA"
         Caption         =   "BUNGA"
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
         DataField       =   "TOTAL"
         Caption         =   "TOTAL"
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
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Pinjaman"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1605
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bunga/Tahun (%)"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   1605
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Lama Cicilan (BLN)"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1605
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Angsuran"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1605
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   14
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Pinjam"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tgl Pinjam"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1605
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Nasabah"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1605
   End
   Begin VB.Label NomorPjm 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   10
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label TanggalPjm 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   9
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label NamaNsb 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   3645
   End
End
Attribute VB_Name = "PINJAMAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
'hub adodc ke database
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBkeuangan.mdb"
'hub adodc ke tabel TBLTrans
Adodc1.RecordSource = "TBLTRANS"
Adodc1.Refresh
'hub datagrid ke adodc
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'panggil nomor pinjam otomatis
Call NomorPinjam
'tanggal diambil dari sistem komputer
TanggalPjm = Date
End Sub

'nomor pinjam otomatis
Private Sub NomorPinjam()
Call BukaDB
RSPinjam.Open "select * from Pinjam Where nomor_pjm In(Select Max(nomor_pjm)From Pinjam)Order By Nomor_pjm Desc", Conn
RSPinjam.Requery
    Dim Urutan As String * 12
    Dim Hitung As Long
    With RSPinjam
        If .EOF Then
            Urutan = "PJM" + Format(Date, "YYMMDD") + "001"
        Else
            If Mid(!nomor_PJM, 4, 6) <> Format(Date, "YYMMDD") Then
                Urutan = "PJM" + Format(Date, "YYMMDD") + "001"
            Else
                Hitung = Right(!nomor_PJM, 3) + 1
                Urutan = "PJM" + Format(Date, "YYMMDD") + Right("000" & Hitung, 3)
            End If
        End If
        NomorPjm = Urutan
    End With
End Sub

'pada saat form dipanggil...
Private Sub Form_Load()
'buka database
Call BukaDB
'buka tabel nasabah
RSNasabah.Open "select * from nasabah", Conn
'bersihkan dulu combo1
Combo1.Clear
'lakukan perulangan untuk ......
Do While Not RSNasabah.EOF
    'menampilkan kode nasabah dalam combo
    Combo1.AddItem RSNasabah!Kode_Nsb
    RSNasabah.MoveNext
Loop
'tutup koneksi ke database
Conn.Close
End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then Text1.SetFocus
End Sub

'jika combo1 dipilih kode nasabahnya maka
Private Sub Combo1_Click()
'buka database
Call BukaDB
'definisikan sebuah recordset bari dengan nama RS1
Dim RS1 As New ADODB.Recordset
'buka recordset dengan mencari data kode dan nama nasabah di tabel nasabah dan di tabel pinjam
'yang kode nasabahnya dipilih dalam combo
RS1.Open "SELECT DISTINCT NASABAH.KODE_NSB,NAMA_NSB FROM NASABAH,PINJAM WHERE PINJAM.KODE_NSB=NASABAH.KODE_NSB AND PINJAM.KODE_NSB='" & Combo1 & "'", Conn
'jika ditemukan, tampilkan nama sabahnya
'dan beri keterangan sebelah kannya bahwa nasabah tsb pernah pinjam
If Not RS1.EOF Then
    NamaNsb = RS1!Nama_Nsb & Space(2) & "Pernah Pinjam"
Else
    'jika tidak ditemukan buka database
    Call BukaDB
    'buka tabel nasabah yang kodenya dipilih di combo1
    RSNasabah.Open "select * from nasabah where kode_nsb='" & Combo1 & "'", Conn
    'tampilkan nama nasabahnya tanpa keterangan pernah pinjam
    NamaNsb = RSNasabah!Nama_Nsb
End If
End Sub

'memformat jumlah pinjamnan
Private Sub Text1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    Text2.SetFocus
    Text1 = Format(Text1, "###,###,###")
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

'bunga pertahun
Private Sub Text2_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then Text3.SetFocus
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

'lama pinjaman dalam bulan
Private Sub Text3_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    'tampilkan berapa angsuran tiap bulannya
    Label5 = Round(Pmt(Text2 / 100 / 12, Text3, Text1), 0) * -1
    'hapus isi tabel temporer jika ada bekas data
    'transaksi sebelumnya (lihat prosedur hapustabel di bawahnya)
    Call HapusTabel
    'buatlah looping
    For No = 1 To Text3
        'tambahkan ke tabel temporer
        Adodc1.Recordset.AddNew
        'tambahkan no
        Adodc1.Recordset!No = No
        Adodc1.Recordset.MoveFirst
        Do While Not Adodc1.Recordset.EOF
            'tampilkan pokok pembayaran
            Adodc1.Recordset!cicilan = Round(PPmt(Text2 / 100 / 12, Adodc1.Recordset!No, Text3, Text1), 2) * -1
            'tampilkan bunga
            Adodc1.Recordset!Bunga = Round(IPmt(Text2 / 100 / 12, Adodc1.Recordset!No, Text3, Text1), 2) * -1
            'tampilkan total (pokok + bunga)
            Adodc1.Recordset!total = Adodc1.Recordset!cicilan + Adodc1.Recordset!Bunga
            Adodc1.Recordset.MoveNext
        Loop
    Next No
    Label5 = Format(Label5, "###,###,###")
    CmdSimpan.SetFocus
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Sub HapusTabel()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub CmdSimpan_Click()
'proses simpan tidak dapat dilakukan jika
'kode nasabah, jumlah pinjam, bunga dan lama cicilan masih kosong
If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
End If

'simpan ke tabel pinjam
Dim simpan As String
simpan = "insert into Pinjam (Nomor_pjm,Tanggal_pjm,Bunga,Waktu,Jumlah_Pjm,Angsuran,Kode_Nsb) values " & _
"('" & NomorPjm & "','" & TanggalPjm & "','" & Text2 & "','" & Text3 & "','" & Text1 & "','" & Label5 & "','" & Combo1 & "')"
Conn.Execute simpan

'simpan ke tabel detailpjm
Adodc1.Recordset.MoveFirst
Dim tmp As Date
tmp = Date
Do While Not Adodc1.Recordset.EOF
    'jatuh tempo dihitung mulai tanggal sekarang + 31 hari
    'agar bulan terus bertambah
    tmp = tmp + 31
    Dim SimpanDetail As String
    SimpanDetail = "Insert Into DetailPjm(Nomor_pjm,Nomor,CICILAN,BUNGA,Total,TEmpo,dibayar,DENDA,ket) values " & _
    "('" & NomorPjm & "','" & Adodc1.Recordset!No & "','" & Adodc1.Recordset!cicilan & "','" & Adodc1.Recordset!Bunga & "','" & Adodc1.Recordset!total & "','" & tmp & "','0','0','-')"
    Conn.Execute (SimpanDetail)
    Adodc1.Recordset.MoveNext
Loop

Form_Activate
Call Kosongkan
Combo1.SetFocus
End Sub

Sub Kosongkan()
Combo1 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Label5 = ""
NamaNsb = ""
Call HapusTabel
Combo1.SetFocus
End Sub

Private Sub CmdBatal_Click()
Form_Activate
Call Kosongkan
Call HapusTabel
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

