VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBuatSEP 
   Caption         =   "Form2"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10470
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "BUAT SEP"
      Height          =   10335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      Begin VB.CommandButton Command1 
         Caption         =   "---"
         Height          =   375
         Left            =   4680
         TabIndex        =   40
         Top             =   1320
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   0
         TabIndex        =   37
         Top             =   9360
         Width           =   10815
         Begin VB.CommandButton cmdCreateSEP 
            Caption         =   "BUAT SEP"
            Height          =   495
            Left            =   4680
            TabIndex        =   39
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton Command5 
            Caption         =   "TUTUP"
            Height          =   495
            Left            =   6360
            TabIndex        =   38
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   36
         Top             =   8880
         Width           =   5415
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   35
         Top             =   8400
         Width           =   5415
      End
      Begin VB.TextBox txtCOB 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   34
         Top             =   6480
         Width           =   1935
      End
      Begin VB.TextBox txtKDPoli 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   33
         Top             =   6000
         Width           =   1455
      End
      Begin VB.TextBox txtKdDiag 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   32
         Top             =   5520
         Width           =   1935
      End
      Begin VB.TextBox txtCatatan 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   31
         Top             =   5040
         Width           =   5055
      End
      Begin VB.ComboBox cbKelasRawat 
         Height          =   360
         Left            =   2520
         TabIndex        =   30
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox cbJenisPelayanan 
         Height          =   360
         Left            =   2520
         TabIndex        =   29
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtNamaPPK 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5280
         TabIndex        =   28
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox ttxNoPPK 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtNoKartu 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   360
         Width           =   5055
      End
      Begin VB.Frame fraJaminan 
         Caption         =   "JAMINAN"
         Height          =   1455
         Left            =   360
         TabIndex        =   16
         Top             =   6840
         Width           =   7935
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   44
            Top             =   960
            Width           =   5415
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ya"
            Height          =   240
            Left            =   1800
            TabIndex        =   19
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Tidak"
            Height          =   240
            Left            =   2640
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Top             =   600
            Width           =   5415
         End
         Begin VB.Label Label6 
            Caption         =   "Lakalantas"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Penjamin"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Lokasi Laka"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "RUJUKAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   9975
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   43
            Top             =   1800
            Width           =   5055
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   42
            Top             =   1320
            Width           =   5055
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   41
            Top             =   360
            Width           =   5055
         End
         Begin MSComCtl2.DTPicker dtpTglRujukan 
            Height          =   375
            Left            =   1800
            TabIndex        =   45
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            Format          =   107216897
            CurrentDate     =   43144
         End
         Begin VB.Label Label8 
            Caption         =   "Tgl Rujukan"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "No Rujukan"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "PPK RUJUKAN"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Asal Rujukan"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   107216897
         CurrentDate     =   43144
      End
      Begin VB.Label Label18 
         Caption         =   "User"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   8880
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "No Telpon"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   8400
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "COB"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "POLI"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "DIAG AWAL"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "CATATAN"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Kelas Rawat"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Jenis Pelayanan"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "PPK PELAYANAN"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Tgl SEP"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "No Kartu"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   13680
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmBuatSEP.frx":0000
      Top             =   13320
      Width           =   7335
   End
End
Attribute VB_Name = "frmBuatSEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim
Dim setting As cSetting
Dim NoKartu As String
Dim KodePPK As String
Dim TglSEP As String
Dim JnsPelayanan As String
Dim KelasPerawatan As String
Dim AsalRujukan As String
Dim TglRujukan As String
Dim NoRujukan As String
Dim PPKRujukan As String
Dim Catatan As String
Dim DiagAwal As String
Dim PoliTujuan As String
Dim PoliEksekutif As String
Dim cob As String
Dim jaminanLakalantas As String
Dim jaminanPenjamin As String
Dim jaminanLokasiLaka As String
Dim noTelp As String
Dim User As String
Dim noMR As String

Private Sub cmdCariNoKartunik_Click()

End Sub

Private Sub cmdCariPoli_Click()
    frmCariPoli.Show
End Sub

Private Sub cmdCreateSEP_Click()
    Set vclaim = New vclaim
    Set setting = New cSetting
    Set vclaim = New vclaim
    setting.GetData
    vclaim.COnsID = setting.COnsID
    vclaim.SecretKey = setting.SecretKey
    vclaim.AlamatWebService = setting.urlWebService
   ' Text6.Text = ""
    
    NoKartu = "0000976032652"
    TglSEP = Format(Now, "yyyy/mm/dd")
    KodePPK = "1101R024"
    JnsPelayanan = "2"
    kelaspelayanan = "3"
    AsalRujukan = "2"
    TglRujukan = Format(Now, "yyyy/mm/dd")
    NoRujukan = "0"
    PPKRujukan = "1101R024"
    Catatan = "aaa"
    jaminanLakalantas = "1"
    jaminanPenjamin = "1"
    jaminanLokasiLaka = "jakarta"
    noTelp = "08552152"
    User = "cobaWS"
    cob = "0"
    DiagAwal = "I10"
    PoliTujuan = "INT"
    PoliEksekutif = "0"
    
    
    
    
    noMR = "000000"
    
    
    Call vclaim.CreateSEP(NoKartu, TglSEP, KodePPK, JnsPelayanan, kelaspelayanan, noMR, AsalRujukan, TglRujukan, NoRujukan, PPKRujukan, Catatan, DiagAwal, PoliTujuan, PoliEksekutif, cob, jaminanLakalantas, jaminanPenjamin, jaminanLokasiLaka, noTelp, User)
    Text6.Text = vclaim.jsonKirim
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Command6_Click()

End Sub
