VERSION 5.00
Begin VB.Form Kasir 
   Caption         =   "Data Kasir"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4140
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
   ScaleHeight     =   1800
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   900
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   900
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "X"
      TabIndex        =   6
      Top             =   840
      Width           =   2475
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Kasir"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Password"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1245
   End
End
Attribute VB_Name = "Kasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pola program kasir sama persis dengan program barang
'perbedaannya hanya pada tujuan tabel (nama-nama field yg beda)

Private Sub Form_Activate()
Call BukaDB
End Sub

Sub Form_Load()
Text1.MaxLength = 30
Text2.MaxLength = 10
KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSKasir.Open "Select * From Kasir where kodeksr='" & Combo1 & "'", Conn
End Function

Private Sub KosongkanText()
    Combo1 = ""
    Text1 = ""
    Text2 = ""
End Sub

Private Sub SiapIsi()
    Combo1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Combo1.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
End Sub

Private Sub KondisiAwal()
    Form_Activate
    KosongkanText
    TidakSiapIsi
    Command1.Caption = "&Input"
    Command2.Caption = "&Edit"
    Command3.Caption = "&Hapus"
    Command4.Caption = "&Tutup"
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
End Sub

Private Sub TampilkanData()
With RSKasir
    Text1 = RSKasir!NamaKsr
    Text2 = RSKasir!PasswordKsr
End With
End Sub

Private Sub command1_Click()
    If Command1.Caption = "&Input" Then
        Command1.Caption = "&Simpan"
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Caption = "&Batal"
        Combo1.Clear
        SiapIsi
        KosongkanText
        Combo1.SetFocus
    Else
        If Combo1 = "" Or Text1 = "" Or Text2 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Kasir (kodeksr,NamaKsr,PasswordKsr) values ('" & Combo1 & "','" & Text1 & "','" & Text2 & "')"
            Conn.Execute SQLTambah
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub command2_Click()
    If Command2.Caption = "&Edit" Then
        Command1.Enabled = False
        Command2.Caption = "&Simpan"
        Command3.Enabled = False
        Command4.Caption = "&Batal"
        SiapIsi
        Combo1.SetFocus
        Call BukaDB
        RSKasir.Open "Kasir", Conn
        Combo1.Clear
        Do Until RSKasir.EOF
            Combo1.AddItem RSKasir!KodeKsr
            RSKasir.MoveNext
        Loop
    Else
        If Combo1 = "" Or Text1 = "" Or Text2 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Kasir Set NamaKsr= '" & Text1 & "',PasswordKsr='" & Text2 & "' where kodeksr='" & Combo1 & "'"
            Conn.Execute SQLEdit
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub command3_Click()
    If Command3.Caption = "&Hapus" Then
        Command1.Enabled = False
        Command2.Enabled = False
        Command4.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Combo1.SetFocus
        Call BukaDB
        RSKasir.Open "Kasir", Conn
        Combo1.Clear
        Do Until RSKasir.EOF
            Combo1.AddItem RSKasir!KodeKsr
            RSKasir.MoveNext
        Loop
    End If
End Sub

Private Sub command4_Click()
    Select Case Command4.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Combo1_Click()
Call CariData
Call TampilkanData
End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Combo1) <> 5 Then
        MsgBox "Kode Harus 5 Digit, Contoh 'KSR01'"
        Combo1.SetFocus
        Exit Sub
    Else
        Text1.SetFocus
    End If

    If Command1.Caption = "&Simpan" Then
        Call CariData
        If Not RSKasir.EOF Then
            TampilkanData
            MsgBox "Kode Kasir Sudah Ada"
            KosongkanText
            Combo1.SetFocus
        Else
            Text1.SetFocus
        End If
    End If
    
    If Command2.Caption = "&Simpan" Then
        Call CariData
        If Not RSKasir.EOF Then
            TampilkanData
            Combo1.Enabled = False
            Text1.SetFocus
        Else
            MsgBox "Kode Kasir Tidak Ada"
            Combo1 = ""
            Combo1.SetFocus
        End If
    End If
    
    If Command3.Enabled = True Then
        Call CariData
        If Not RSKasir.EOF Then
            TampilkanData
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                Dim SQLHapus As String
                SQLHapus = "Delete From Kasir where kodeksr= '" & Combo1 & "'"
                Conn.Execute SQLHapus
                Form_Activate
                Call KondisiAwal
            Else
                Form_Activate
                Call KondisiAwal
                Command3.SetFocus
            End If
        Else
            MsgBox "Data Tidak ditemukan"
            Combo1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then
        If Command1.Enabled = True Then
            Command1.SetFocus
        ElseIf Command2.Enabled = True Then
            Command2.SetFocus
        End If
    End If
End Sub


