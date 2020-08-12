VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MENU 
   Caption         =   "MENU UTAMA"
   ClientHeight    =   3285
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   4725
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
   Picture         =   "MENU.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2790
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "NAMA KASIR : "
            TextSave        =   "NAMA KASIR : "
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "14:17"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/08/20"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MNFILE 
      Caption         =   "FILE"
      Begin VB.Menu MNKASIR 
         Caption         =   "KASIR"
      End
      Begin VB.Menu MNNASABAH 
         Caption         =   "NASABAH"
      End
   End
   Begin VB.Menu MNTRANSAKSI 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu MNPINJAMAN 
         Caption         =   "PINJAMAN"
      End
      Begin VB.Menu MNBAYARCICILAN 
         Caption         =   "BAYAR CICILAN"
      End
   End
   Begin VB.Menu MNLAP 
      Caption         =   "LAPORAN"
   End
   Begin VB.Menu MNKELUAR 
      Caption         =   "KELUAR"
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
If Keyascii = 13 Then PEMBAYARAN.Show vbModal
End Sub

Private Sub MNBAYARCICILAN_Click()
PEMBAYARAN.Show vbModal
End Sub

Private Sub MNKASIR_Click()
Kasir.Show vbModal
End Sub

Private Sub MNKELUAR_Click()
End
End Sub

Private Sub MNLAP_Click()
LAPORAN.Show
End Sub

Private Sub MNNASABAH_Click()
NASABAH.Show vbModal
End Sub

Private Sub MNPINJAMAN_Click()
PINJAMAN.Show vbModal
End Sub

Private Sub SQL_Click()
UjiSQL.Show vbModal
End Sub
