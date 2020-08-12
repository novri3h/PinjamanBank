VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3810
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   ScaleHeight     =   1575
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKodeKsr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   2000
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.TextBox TxtNamaKsr 
         Height          =   350
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   2000
      End
      Begin VB.TextBox TxtPasswordKsr 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama User"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'pada saat form dipanggil maka
Private Sub Form_Load()
    'batasi panjang karakter yang diketik di TxtNamaKsr maksimal 30 digit
    TxtNamaKsr.MaxLength = 25
    TxtPasswordKsr.MaxLength = 10
    'bentuk huruf yg diketik di TxtPasswordKsr jadi huruf x
    TxtPasswordKsr.PasswordChar = "X"
    'txtpasswordksr tdk dpt dimasuki kursor
    TxtPasswordKsr.Enabled = False
    TxtKodeKsr.Enabled = False
End Sub

Private Sub TxtNamakSR_KeyPress(Keyascii As Integer)
'ubah huruf jd besar semua
Keyascii = Asc(UCase(Chr(Keyascii)))
'27 adalah tombol ESC (Escape)
'jika menekan tombol ESC maka tutup form. jadi hati-hati dengan menekan ESC
If Keyascii = 27 Then Unload Me
'jika menekan enter setelah mengetik nama user
If Keyascii = 13 Then
    'buka database
    Call BukaDB
    'buka tabel kasir dan cari nama kasir yg diketik di txtkodeksr
    RSKasir.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
    'jika nama tidak ditemukan maka
    If RSKasir.EOF Then
        'tampilkan pesan
        MsgBox "Nama Kasir tidak terdaftar"
        TxtNamaKsr = ""
        'kursor kembali ke nama kasir
        TxtNamaKsr.SetFocus
        Exit Sub
    Else
        'jika nama kasir ditemukan, txtnamaksr mati
        TxtNamaKsr.Enabled = False
        'txtpasswordksr diaktifkan
        TxtPasswordKsr.Enabled = True
        'password jadi fokus kursor
        TxtPasswordKsr.SetFocus
        Conn.Close
    End If
End If
End Sub

Private Sub txtpasswordkSR_KeyPress(Keyascii As Integer)
'ubah hrf jd besar semua
Keyascii = Asc(UCase(Chr(Keyascii)))
'jika menekan esc tutup form
If Keyascii = 27 Then Unload Me
'definisikan loginkasir,kodekasir dan namakasir sebagai string
Dim KodeKasir As String
Dim NamaKasir As String
'jika menekan enter setelah mengetik password
If Keyascii = 13 Then
    'buka database
    Call BukaDB
    'buka dan cari di tabel kasir data namakasir dan passwordnya yang cocok (yang diketik
    'di txtnamaksr dan txtpassowrdksr
    RSKasir.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PasswordKsr='" & TxtPasswordKsr & "'", Conn
    'jika data tidak ditemukan (nama dan passwordnya tdk cocok)
    If RSKasir.EOF Then
        'tampilkan pesan
        MsgBox "Password salah"
        TxtPasswordKsr = ""
        'kursor kembali ke password untuk diketik lagi
        TxtPasswordKsr.SetFocus
        Exit Sub
    Else
        'jika data ditemukan (nama dan passwordnya cocok)maka tutup form login
        Unload Me
        'panggil form menu
        MENU.Show
        'string kodekasir yang telah didefinisikan diisi dengan kodeksr yg diambil dari tabel
        KodeKasir = RSKasir!KodeKsr
        'string namakasir yang telah didefinisikan diisi dengan nama kasir yg diambil dari tabel
        NamaKasir = RSKasir!NamaKsr
        TxtKodeKsr = KodeKasir
        TxtNamaKsr = NamaKasir
        MENU.StatusBar1.Panels(1) = NamaKasir
       
    End If
End If
End Sub

