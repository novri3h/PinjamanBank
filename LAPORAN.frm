VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LAPORAN 
   Caption         =   "LAPORAN"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2730
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
   ScaleHeight     =   3585
   ScaleWidth      =   2730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Data Nasabah"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1920
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Laporan Pembayaran"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2500
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2000
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laporan Pinjaman"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2500
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2000
      End
   End
End
Attribute VB_Name = "LAPORAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
CR.ReportFileName = App.Path & "\Lap nasabah.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub Form_Load()
Call BukaDB
RSPinjam.Open "select * from Pinjam", Conn
Combo1.Clear
Do While Not RSPinjam.EOF
    Combo1.AddItem RSPinjam!nomor_PJM
    Combo2.AddItem RSPinjam!nomor_PJM
    RSPinjam.MoveNext
Loop

RSBayar.Open "select DISTINCT NOMOR_PJM from bayar", Conn
Combo3.Clear
Do While Not RSBayar.EOF
    Combo3.AddItem RSBayar!nomor_PJM
    RSBayar.MoveNext
Loop

End Sub

Private Sub Combo1_Click()
CR.SelectionFormula = "{Pinjam.nomor_Pjm}='" & Combo1 & "'"
CR.ReportFileName = App.Path & "\Lap pinjam.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub Combo2_Click()
CR.SelectionFormula = "{Pinjam.nomor_Pjm}='" & Combo2 & "'"
CR.ReportFileName = App.Path & "\Lap Bayar1.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub

Private Sub Combo3_Click()
CR.SelectionFormula = "{Pinjam.nomor_Pjm}='" & Combo3 & "'"
CR.ReportFileName = App.Path & "\Lap Bayar2.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1
End Sub


