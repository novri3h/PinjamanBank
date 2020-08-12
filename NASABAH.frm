VERSION 5.00
Begin VB.Form NASABAH 
   Caption         =   "DATA NASABAH"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
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
   ScaleHeight     =   2160
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   120
      Width           =   3000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   2760
      TabIndex        =   9
      Top             =   1680
      Width           =   1250
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   350
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   1250
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1250
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "NASABAH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
'buka database yang telah didefinisikan dalam modul
Call BukaDB
End Sub

Sub Form_Load()
'panggil prosedur kondisiawal
KondisiAwal
End Sub

Private Sub KondisiAwal()
    Form_Activate
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
End Sub

'buat prosedur pencarian data nasabah
Function CariData()
    Call BukaDB
    RSNasabah.Open "Select * From Nasabah where Kode_Nsb='" & Combo1 & "'", Conn
End Function

'nomor nasabah otomatis
Private Sub AutoKode_Nsb()
Call BukaDB
RSNasabah.Open "select * from Nasabah Where Kode_Nsb In(Select Max(Kode_Nsb)From Nasabah)Order By Kode_Nsb Desc", Conn
RSNasabah.Requery
    Dim Urutan As String * 12
    Dim Hitung As Long
    With RSNasabah
        If .EOF Then
            Urutan = "NSB" + Format(Date, "YY") + Format(Date, "MM") + Format(Date, "DD") + "001"
        Else
            If Mid(!Kode_Nsb, 4, 6) <> Format(Date, "YY") + Format(Date, "MM") + Format(Date, "DD") Then
                Urutan = "NSB" + Format(Date, "YY") + Format(Date, "MM") + Format(Date, "DD") + "001"
            Else
                Hitung = Right(!Kode_Nsb, 3) + 1
                Urutan = "NSB" + Format(Date, "YY") + Format(Date, "MM") + Format(Date, "DD") + Right("000" & Hitung, 3)
            End If
        End If
        Combo1 = Urutan
    End With
End Sub

'pada saat combo1 diklik
Private Sub Combo1_Click()
'cari data
Call CariData
'jika ditemukan tampilkan datanya
'dengan memanggil prosedur tampilkandata
Call TampilkanData
End Sub

Private Sub TampilkanData()
Text1 = RSNasabah!Nama_Nsb
Text2 = RSNasabah!Alamat_Nsb
Text3 = RSNasabah!Telepon_Nsb
End Sub

Private Sub KosongkanText()
    Combo1 = ""
    Text1 = ""
    Text2 = ""
    Text3 = ""
End Sub

Private Sub SiapIsi()
    Combo1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Combo1.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
End Sub

'pada saat cmdinput diklik, maka
Private Sub CmdInput_Click()
    'jika cmdinput captionya "Input" maka
    If CmdInput.Caption = "&Input" Then
        'atur caption masing-masing command
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        'semua textbox dan combo daat dimasuki kursor
        SiapIsi
        KosongkanText
        Combo1.SetFocus
        'panggil nomor nasabah otomatis
        Call AutoKode_Nsb
        'matikan combo1 agar nomor nsabah tidak dapat diubah
        Combo1.Enabled = False
        'kursor ke text1 (nama nasabah)
        Text1.SetFocus
    Else
        'jika masih ada data yg kosong maka..
        If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Then
            'tampilkan pesan
            MsgBox "Data Belum Lengkap...!"
        Else
            'jika semua data telah diisi, maka simpan data
            Dim SQLTambah As String
            SQLTambah = "Insert Into Nasabah (Kode_Nsb,Nama_Nsb,Alamat_Nsb,Telepon_Nsb) values ('" & Combo1 & "','" & Text1 & "','" & Text2 & "','" & Text3 & "')"
            Conn.Execute SQLTambah
            Form_Activate
            'kembali ke kondisi awal
            Call KondisiAwal
        End If
    End If
End Sub

'pola program di command edit hampir sama dengan
'program di command input. bedanya hanya mengedit (update) saja
Private Sub CmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdTutup.Caption = "&Batal"
        SiapIsi
        'buka database
        Call BukaDB
        'tampilkan kode nasabah di combo
        RSNasabah.Open "select * from nasabah", Conn
        Combo1.Clear
        Do While Not RSNasabah.EOF
            Combo1.AddItem RSNasabah!Kode_Nsb
            RSNasabah.MoveNext
        Loop
        Combo1.SetFocus
    Else
        'jika masih ada data yg kosong...
        If Text1 = "" Or Text2 = "" Or Text3 = "" Then
            'tampilkan pesan
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            'jika semua data telah diisi, maka update data
            Dim SQLEdit As String
            SQLEdit = "Update Nasabah Set Nama_Nsb= '" & Text1 & "', Alamat_Nsb='" & Text2 & "',Telepon_Nsb='" & Text3 & "' where Kode_Nsb='" & Combo1 & "'"
            Conn.Execute SQLEdit
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

'command tutup bekerja berdasarkan kondisi captionya
Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

'jika menekan enter setelah memilih data di combo, maka
Private Sub Combo1_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    'jika saat itu cmdinput captionya simpan maka...
    If CmdInput.Caption = "&Simpan" Then
        'cari data nasabah
        Call CariData
        'jika ditemukan
        If Not RSNasabah.EOF Then
            'tampilkan datanya
            TampilkanData
            'munculkan pesan
            MsgBox "Kode_Nsb Nasabah Sudah Ada"
            KosongkanText
            Combo1.SetFocus
        Else
            'jika tidak ditemukan, lanjutkan mengisi nama nasabah
            Text1.SetFocus
        End If
    End If
    
    'jika saat itu cmdedit captionnya simpan, maka
    If CmdEdit.Caption = "&Simpan" Then
        'cari data
        Call CariData
        'jika ditemukan
        If Not RSNasabah.EOF Then
            'tampilkan datanya
            TampilkanData
            'matikan combo
            Combo1.Enabled = False
            'ganti nama nasabah
            Text1.SetFocus
        Else
            'jika tidak ditemukan, munculkan pesan
            MsgBox "Kode Nasabah Tidak Ada"
            Combo1 = ""
            Combo1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
    'ubah huruf jadi besar semua
    Keyascii = Asc(UCase(Chr(Keyascii)))
    'jika menekan enter kursor pindah ke text2
    If Keyascii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    'no telepon hanya dapat diisi angka
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

