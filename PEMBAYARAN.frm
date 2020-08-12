VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PEMBAYARAN 
   Caption         =   "PEMBAYARAN CICILAN"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
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
   ScaleHeight     =   6570
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus Dulu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   32
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   8775
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No Bayar"
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tgl Bayar"
         Height          =   345
         Left            =   4680
         TabIndex        =   30
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label NomorByr 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label TanggalByr 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   6720
         TabIndex        =   28
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.CommandButton CmdMaju 
      Caption         =   "Majur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdMundur 
      Caption         =   "Mundur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "PEMBAYARAN.frx":0000
      Height          =   1905
      Left            =   240
      TabIndex        =   23
      Top             =   4560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3360
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "NOMOR_PJM"
         Caption         =   "NOMOR_PJM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NOMOR"
         Caption         =   "NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "CICILAN"
         Caption         =   "CICILAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """Rp""#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "BUNGA"
         Caption         =   "BUNGA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TOTAL"
         Caption         =   "TOTAL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TEMPO"
         Caption         =   "TEMPO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "DENDA"
         Caption         =   "DENDA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "DIBAYAR"
         Caption         =   "DIBAYAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "KEMBALI"
         Caption         =   "KEMBALI"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "KET"
         Caption         =   "KET"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   500
      Left            =   4800
      TabIndex        =   4
      Top             =   3960
      Width           =   4035
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   500
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   3435
   End
   Begin VB.TextBox TxtDibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6840
      TabIndex        =   3
      Top             =   3120
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   2000
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telah Dibayar"
      Height          =   345
      Left            =   4800
      TabIndex        =   40
      Top             =   1320
      Width           =   1995
   End
   Begin VB.Label TlhBayar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6840
      TabIndex        =   39
      Top             =   1320
      Width           =   1995
   End
   Begin VB.Label SisaPjm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6840
      TabIndex        =   38
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa Pinjaman"
      Height          =   345
      Left            =   4800
      TabIndex        =   37
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Kembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Left            =   6840
      TabIndex        =   36
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   4800
      TabIndex        =   35
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   4800
      TabIndex        =   34
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Left            =   6840
      TabIndex        =   33
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Label LblDenda 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6840
      TabIndex        =   26
      Top             =   2400
      Width           =   1995
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Angsuran"
      Height          =   345
      Left            =   4800
      TabIndex        =   22
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Label Angsuran 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6840
      TabIndex        =   21
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Waktu"
      Height          =   345
      Left            =   240
      TabIndex        =   20
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Pinjaman"
      Height          =   345
      Left            =   4800
      TabIndex        =   19
      Top             =   960
      Width           =   1995
   End
   Begin VB.Label Waktu 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   18
      Top             =   2400
      Width           =   1995
   End
   Begin VB.Label JumlahPjm 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6840
      TabIndex        =   17
      Top             =   960
      Width           =   1995
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tgl Pinjam"
      Height          =   345
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bunga"
      Height          =   345
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label TanggalPjm 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   14
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Bunga 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   13
      Top             =   2040
      Width           =   1995
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jatuh Tempo"
      Height          =   345
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Label LblTempo 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   11
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Bayar"
      Height          =   345
      Left            =   4800
      TabIndex        =   10
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cicilan Ke"
      Height          =   345
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Denda"
      Height          =   345
      Left            =   4800
      TabIndex        =   8
      Top             =   2400
      Width           =   1995
   End
   Begin VB.Label NamaNsb 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   1995
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Peminjam"
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Pinjaman"
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1395
   End
End
Attribute VB_Name = "PEMBAYARAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
'hub adodc ke database
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBkeuangan.mdb"
'sumber data adodc adalah tabel detailpjm yang kodenya xxx
'hal ini akan menyebabkan datagrid jadi kosong karena nomorpinjamnya tidak diketahui
Adodc1.RecordSource = "select * from Detailpjm where nomor_pjm='xxx'"
Adodc1.Refresh
'hub datagrid ke adodc
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'panggil nomor pembayaran otomatis
Call Nomorbayar
'tanggal bayar diambil dari sistem komputer
TanggalByr = Date
End Sub

'nomor bayar otomatis

Private Sub Nomorbayar()
Call BukaDB
RSBayar.Open "select * from Bayar Where nomor_BYR In(Select Max(nomor_BYR)From Bayar)Order By Nomor_BYR Desc", Conn
RSBayar.Requery
    Dim Urutan As String * 12
    Dim Hitung As Long
    With RSBayar
        If .EOF Then
            Urutan = "BYR" + Format(Date, "YYMMDD") + "001"
        Else
            If Mid(!Nomor_BYR, 4, 6) <> Format(Date, "YYMMDD") Then
                Urutan = "BYR" + Format(Date, "YYMMDD") + "001"
            Else
                Hitung = Right(!Nomor_BYR, 3) + 1
                Urutan = "BYR" + Format(Date, "YYMMDD") + Right("000" & Hitung, 3)
            End If
        End If
        NomorByr = Urutan
    End With
End Sub


'pada saat form dipanggil
Private Sub Form_Load()
'kunci semua objek kecuali setelah cmdinput diklik
Call Terkunci
End Sub

Sub Terkunci()
Combo1.Enabled = False
Combo2.Enabled = False
LblDenda.Enabled = False
TxtDibayar.Enabled = False
End Sub

'pada saat cmdinput diklik...
Private Sub CmdInput_Click()
If CmdInput.Caption = "&Input" Then
    CmdInput.Caption = "&Simpan"
    CmdTutup.Caption = "&Batal"
    'buka database
    Call BukaDB
    'buka tabel pinjam
    RSPinjam.Open "select * from Pinjam", Conn
    Combo1.Clear
    'tampilkan nomor pinjam dalam combo
    Do While Not RSPinjam.EOF
        Combo1.AddItem RSPinjam!nomor_PJM
        RSPinjam.MoveNext
    Loop
    Conn.Close
    'buka objek-objek
    Call Terbuka
    Combo1.SetFocus
Else
    'jika nomot pinjam atau cicilan atau pembayaran masih kosong
    If Combo1 = "" Or Combo2 = "" Or TxtDibayar = "" Then
        'tampilkan pesan...
        MsgBox "Data belum lengkap...!"
        Exit Sub
    Else
        'jika data sudah lengkap buka database
        Call BukaDB
        'buka tabel detailpjm yang nomor pinjamnya = combo1 dan nomornya = combo2 (cicilan)
        RSDetailPjm.Open "select * from detailpjm where nomor_pjm='" & Combo1 & "' and val(nomor)='" & Combo2 & "'", Conn
        If Not RSDetailPjm.EOF Then
            'jika ditemukan update datanya
            Dim bayar As String
            bayar = "update detailpjm set dibayar='" & Val(TxtDibayar) & "',DENDA='" & Val(LblDenda) & "',KEMBALI='" & Val(Kembali) & "',Ket='LUNAS' where nomor_pjm='" & Combo1 & "' and val(nomor)='" & Combo2 & "'"
            Conn.Execute bayar
        End If
        
        'simpan juga data ke tabel Bayar dan beri ket LUNAS
        'denda dihitung berdasarkan tanggal bayar dan tgl jatuh tempo
        Dim simpanbayar As String
        simpanbayar = "insert into bayar (nomor_byr,tanggal_byr,nomor_pjm,denda,jumlah_byr,ket) values " & _
        "('" & NomorByr & "','" & CDate(TanggalByr) & "','" & Combo1 & "','" & Val(LblDenda) & "','" & Val(TxtDibayar) & "','LUNAS')"
        Conn.Execute simpanbayar
        
        Call Awal
        Form_Activate
    End If
End If
End Sub

'jika nomor pinjam dipilih dalam combo maka...
Private Sub Combo1_Click()
'buka Database
Call BukaDB
'buka tabel pinjam yang nomor pinjamnya dipilih di combo
RSPinjam.Open "select * from Pinjam where nomor_pjm='" & Combo1 & "'", Conn
'tampilkan tanggal pinjam, bunga, waktu dan sebagainya
TanggalPjm = RSPinjam!Tanggal_Pjm
Bunga = RSPinjam!Bunga & Space(2) & "%"
Waktu = RSPinjam!Waktu & Space(2) & "Bulan"
'JumlahPjm = Format(RSPinjam!Jumlah_Pjm, "###,###,###")
JumlahPjm = RSPinjam!Jumlah_Pjm
Angsuran = RSPinjam!Angsuran 'Format(RSPinjam!Angsuran, "###,###,###")

'buka tabel detail pinjam dan berapa yang telah dibayar (digabungkan)
RSDetailPjm.Open "select sum(dibayar) as telahbayar from detailpjm where nomor_pjm='" & Combo1 & "'", Conn
If Not RSDetailPjm.EOF Then
    TlhBayar = RSDetailPjm!telahbayar
    SisaPjm = JumlahPjm - RSDetailPjm!telahbayar
End If

'cari juga data nasabah yang nomor pinjamnya tersimpan di tabel pinjam
'jika ditemukan
If Not RSPinjam.EOF Then
    RSNasabah.Open "select * from nasabah where kode_Nsb='" & RSPinjam!Kode_Nsb & "'", Conn
    'tampilkan nama nasabahnnya
    NamaNsb = RSNasabah!Nama_Nsb
End If
Conn.Close
'tampilkan data pinjamnnya diambil dari tabel detailpjm
'yang nomor pinjamnya dipilih dalam combo
Adodc1.RecordSource = "select * from detailpjm where nomor_pjm='" & Combo1 & "' ORDER BY NOMOR"
Adodc1.Refresh
'koding dibawah ini hanya digunakan untuk
'menampilkan jumlah cicilannya saja
Call BukaDB
RSDetailPjm.Open "select * from detailpjm where nomor_pjm='" & Combo1 & "'", Conn
Combo2.Clear
Do While Not RSDetailPjm.EOF
    'inilah yang diperlukan yaitu nomor (cicilan ke...)
    Combo2.AddItem RSDetailPjm!NOMOR
    RSDetailPjm.MoveNext
Loop
End Sub

'sebaiknya tekan dulu tombol tanda panah ke atas
'kursor di datagrid akan menuju ke baris nomor cicilan
'sesuai dengan nomor yang dipilih di combo
Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    'jika menekan tanda panah ke atas
    Case vbKeyUp
        'panggil function cmdmundur
'        CmdMundur_Click
    'jika menekan tanda panah ke bawah
    Case vbKeyDown
        'panggil function cmdmaju
'        CmdMaju_Click
  End Select
End Sub

'koding di bawah ini digunakan untuk mencari
'tanggal jatuh tempo dan denda. denda dihitung 5000 per hari keterlambatan
Private Sub Combo2_Click()
Call BukaDB
RSDetailPjm.Open "select * from detailpjm where nomor_pjm='" & Combo1 & "' and val(nomor)='" & Combo2 & "' ", Conn
If Not RSDetailPjm.EOF Then
    If RSDetailPjm!KET = "LUNAS" Then
        TxtDibayar.Enabled = False
        MsgBox "CICILAN KE '" & Combo2 & "' SUDAH LUNAS"
        Exit Sub
    Else
        TxtDibayar.Enabled = True
        LblTempo = RSDetailPjm!tempo
        If CDate(TanggalByr) > CDate(LblTempo) Then
            LblDenda = (CDate(TanggalByr) - CDate(LblTempo)) * 5000
        Else
            LblDenda = 0
        End If
        LblTotal = Val(LblDenda) + Val(Angsuran)
    End If
End If
End Sub

Private Sub CmdBatal_Click()
Form_Activate
'Call Kosongkan
End Sub

Private Sub CmdTutup_Click()
On Error Resume Next
Select Case CmdTutup.Caption
    Case "&Tutup"
        Unload Me
    Case "&Batal"
        Call Awal
        Form_Activate
End Select
End Sub

Sub Terbuka()
Combo1.Enabled = True
Combo2.Enabled = True
LblDenda.Enabled = True
TxtDibayar.Enabled = True
End Sub

'jumlah pembayaran cicilan tidak boleh kosong
'tidak boleh kurang. pembayaran akan langsung
'disimpan ke datagrid tepat di cicilan ke.....
'sesuai nilai dalam combo dan ket diisi LUNAS
Private Sub TxtDibayar_KeyPress(Keyascii As Integer)
'On Error Resume Next
If Keyascii = 13 Then
    'TxtDibayar = Format(TxtDibayar, "###,###,###")
    If TxtDibayar = "" Then
        MsgBox "Jumlah Bayar masih kosong"
        TxtDibayar.SetFocus
        Exit Sub
    ElseIf Val(TxtDibayar) < Val(LblTotal) Then
        MsgBox "pembayaran kurang"
        TxtDibayar.SetFocus
        Exit Sub
    Else
        Kembali = Val(TxtDibayar) - Val(LblTotal)
        DataGrid1.Row = Val(Combo2) - 1
        Adodc1.Recordset!DENDA = Val(LblDenda)
        Adodc1.Recordset!DIBAYAR = Val(TxtDibayar)
        Adodc1.Recordset!Kembali = Val(Kembali)
        Adodc1.Recordset!KET = "LUNAS"
        CmdInput.SetFocus
    End If
End If
End Sub

Sub Awal()
    Combo1.Clear
    CmdTutup.Caption = "&Tutup"
    CmdInput.Caption = "&Input"
    Terkunci
    TanggalPjm = ""
    Bunga = ""
    Waktu = ""
    JumlahPjm = ""
    Angsuran = ""
    NamaNsb = ""
    Combo2.Clear
    LblTempo = ""
    LblDenda = ""
    TxtDibayar = ""
    LblTotal = ""
    Kembali = ""
    TlhBayar = ""
    SisaPjm = ""
End Sub

Private Sub command1_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    Dim hapus As String
    hapus = "Update detailpjm set DENDA='0',dibayar='0',KEMBALI='0',ket='_' where nomor_pjm='" & Combo1 & "'"
    Conn.Execute hapus
    Adodc1.Recordset.MoveNext
Loop
Adodc1.Refresh
DataGrid1.Refresh
End Sub


