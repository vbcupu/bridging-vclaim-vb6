VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCariDataPeserta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARI DATA PESERTA BPJS"
   ClientHeight    =   5955
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   6720
      TabIndex        =   12
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Frame fraCariPesertaBPJS 
      Caption         =   "Cari Peserta"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      Begin TabDlg.SSTab SSTab1 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2566
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Cari By NIK"
         TabPicture(0)   =   "frmBridging.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtNoKartuNIK"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdCariByNIK"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Cari by No Kartu"
         TabPicture(1)   =   "frmBridging.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(1)=   "txtNoKartuBPJS"
         Tab(1).Control(2)=   "cmdCariNoKartu"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Cari By No Rujukan"
         TabPicture(2)   =   "frmBridging.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3"
         Tab(2).Control(1)=   "txtNoRujukan"
         Tab(2).Control(2)=   "cmdCari"
         Tab(2).Control(3)=   "OptPCare"
         Tab(2).Control(4)=   "optRS"
         Tab(2).ControlCount=   5
         Begin VB.OptionButton optRS 
            Caption         =   "RUMAH SAKIT"
            Height          =   375
            Left            =   -72360
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptPCare 
            Caption         =   "PCARE"
            Height          =   375
            Left            =   -73440
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdCari 
            Caption         =   "CARI"
            Height          =   375
            Left            =   -71880
            TabIndex        =   11
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtNoRujukan 
            Height          =   375
            Left            =   -74760
            TabIndex        =   9
            Text            =   "110103030918Y000002"
            Top             =   840
            Width           =   2655
         End
         Begin VB.CommandButton cmdCariNoKartu 
            Caption         =   "CARI"
            Height          =   375
            Left            =   -71880
            TabIndex        =   8
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtNoKartuBPJS 
            Height          =   375
            Left            =   -74760
            TabIndex        =   6
            Text            =   "0000976032652"
            Top             =   840
            Width           =   2655
         End
         Begin VB.CommandButton cmdCariByNIK 
            Caption         =   "CARI"
            Height          =   375
            Left            =   4080
            TabIndex        =   4
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtNoKartuNIK 
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Text            =   "3374132708800008"
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "No Rujukan"
            Height          =   255
            Left            =   -74760
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "No Kartu BPJS"
            Height          =   255
            Left            =   -74760
            TabIndex        =   7
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "No NIK"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   2655
         End
      End
   End
   Begin VB.TextBox txtHasil 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2040
      Width           =   8535
   End
End
Attribute VB_Name = "frmCariDataPeserta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim
Dim setting As csetting
Dim TglUnix As String
Dim tglSekarang As Date

Private Sub cmdCari_Click()
    Set vclaim = New vclaim
    vclaim.CariByRujukanPcare (txtNoRujukan.Text)
    If vclaim.ServerCode = "200" Then
    With vclaim.Hasil
       txtHasil.Text = "kodediagnosa : " & .Item("kodeDiagnosa") & vbNewLine & "namadiagnosa : " & .Item("namaDiagnosa") & vbNewLine & "keluhan : " & .Item("keluhan") & vbNewLine & "nokunjungan : " & .Item("noKunjungan") & vbNewLine & "kodepelayanan : " & .Item("kodePelayanan") & vbNewLine & _
                       "namapelayanan : " & .Item("namaPelayanan") & vbNewLine & "cobNmAsuransi : " & .Item("cobNmAsuransi") & vbNewLine & "cobnoasuransi : " & .Item("cobNoAsuransi") & vbNewLine & "cobTglTAT : " & .Item("cobTglTAT") & vbNewLine & "cobtgltmt : " & .Item("cobTglTMT") & vbNewLine & _
                       "kodehakkelas : " & .Item("kodeHakKelas") & vbNewLine & "namaHakKelas : " & .Item("namaHakKelas") & vbNewLine & "informasidinsos : " & .Item("informasiDinsos") & vbNewLine & "informasiNoSKTM : " & .Item("informasiNoSKTM") & vbNewLine & "informasiprolanisPRB : " & .Item("informasiProlanisPRB") & vbNewLine & "kodejenispeserta : " & .Item("kodeJenisPeserta") & vbNewLine & _
                       "keteranganjenispeserta : " & .Item("namaJenisPeserta") & vbNewLine & "nomr : " & .Item("nomr") & vbNewLine & "nomrtelepon : " & .Item("nomrtelepon") & vbNewLine & "nama : " & .Item("nama") & vbNewLine & "nik : " & .Item("nik") & vbNewLine & "nokartu : " & .Item("noKartu") & vbNewLine & "pisa : " & .Item("pisa") & vbNewLine & "kodeprovumum : " & .Item("kodeProvUmum") & vbNewLine & "namaprovumum : " & .Item("namaProvumum") & vbNewLine & _
                       "JK : " & .Item("jk") & vbNewLine & "statuspeserta : " & .Item("namaStatusPeserta") & vbNewLine & "statuspesertakode : " & .Item("kodeStatusPeserta") & vbNewLine & "tglcetakkartu : " & .Item("tglCetakKartu") & vbNewLine & "tglLahir : " & .Item("tglLahir") & vbNewLine & "tgltat : " & .Item("tglTAT") & vbNewLine & "tglTMT : " & .Item("tglTMT") & vbNewLine & "umursaatpelayanan : " & .Item("umurSaatPelayanan") & vbNewLine & "umursekarang : " & .Item("umurSekarang") & vbNewLine & _
                       "kodepolirujukan : " & .Item("kodePoliRujukan") & vbNewLine & "namaPoliRujukan : " & .Item("namaPoliRujukan") & vbNewLine & "kodeprovperujuk : " & .Item("kodeProvPerujuk") & vbNewLine & "namaprovperujuk : " & .Item("namaProvPerujuk") & vbNewLine & "tglDirujuk : " & .Item("tgldirujuk") & vbNewLine
    End With
    End If
End Sub

Private Sub cmdCariByNIK_Click()
     Call vclaim.CariPesertaByNIK(txtNoKartuNIK.Text)
     If vclaim.ServerCode = "200" Then
     With vclaim.Hasil
       
        txtHasil.Text = "COBAsuransi: " & .Item("cobNmAsuransi") & vbNewLine & "COBNoAsuransi: " & .Item("cobNoAsuransi") & vbNewLine & "COBPesertaTglTAT: " & .Item("cobTglTAT") & vbNewLine & _
                        "COBPesertaTglTMT: " & .Item("COBTglTMT") & vbNewLine & _
                        "hakkelas: " & .Item("namaHakKelas") & vbNewLine & _
                        "hakkelasKode: " & .Item("kodeHakKelas") & vbNewLine & _
                        "informasiDinsos: " & .Item("informasiDinsos") & vbNewLine & _
                        "informasiNoSKTM: " & .Item("informasiNoSKTM") & vbNewLine & _
                        "informasiProlanisPRB: " & .Item("informasiProlanisPRB") & vbNewLine & _
                        "jenisPeserta: " & .Item("namaJenisPeserta") & vbNewLine & _
                        "jenisPesertaKode: " & .Item("kodeJenisPeserta") & vbNewLine & _
                        "noMR: " & .Item("nomr") & vbNewLine & _
                        "noMRnoTelepon: " & .Item("nomrnoTelepon") & vbNewLine & _
                        "nama: " & .Item("nama") & vbNewLine & _
                        "nik: " & .Item("nik") & vbNewLine & _
                        "NoKartu: " & .Item("noKartu") & vbNewLine & _
                        "pisa: " & .Item("pisa") & vbNewLine & _
                        "provUmum: " & .Item("kodeProvUmum") & vbNewLine & _
                        "provuUmum: " & .Item("namaProvUmum") & vbNewLine & _
                        "jk: " & .Item("jk") & vbNewLine & _
                        "statusPeserta: " & .Item("namaStatusPeserta") & vbNewLine & _
                        "statusPesertaKode: " & .Item("kodeStatusPeserta") & vbNewLine & _
                        "tglCetakKartu: " & .Item("tglCetakKartu") & vbNewLine & _
                        "tglLahir: " & .Item("tglLahir") & vbNewLine & _
                        "tglTAT: " & .Item("tglTAT") & vbNewLine & _
                        "tglTMT: " & .Item("tglTMT") & vbNewLine & _
                        "umurSaatPelayanan: " & .Item("umurSaatPelayanan") & vbNewLine & "umurSekarang: " & .Item("umurSekarang")
    End With
    End If
End Sub

Private Sub cmdCariNoKartu_Click()
    'Call vclaim.CariDataByNoKartu(txtNoKartuBPJS.Text)
    Call vclaim.CariPesertaByNoKartu("0000077082671")
    With vclaim.Hasil
        txtHasil.Text = "COBAsuransi: " & .Item("cobNmAsuransi") & vbNewLine & "COBNoAsuransi: " & .Item("cobNoAsuransi") & vbNewLine & "COBPesertaTglTAT: " & .Item("cobTglTAT") & vbNewLine & _
                        "COBPesertaTglTMT: " & .Item("COBTglTMT") & vbNewLine & _
                        "hakkelas: " & .Item("namaHakKelas") & vbNewLine & _
                        "hakkelasKode: " & .Item("kodeHakKelas") & vbNewLine & _
                        "informasiDinsos: " & .Item("informasiDinsos") & vbNewLine & _
                        "informasiNoSKTM: " & .Item("informasiNoSKTM") & vbNewLine & _
                        "informasiProlanisPRB: " & .Item("informasiProlanisPRB") & vbNewLine & _
                        "jenisPeserta: " & .Item("namaJenisPeserta") & vbNewLine & _
                        "jenisPesertaKode: " & .Item("kodeJenisPeserta") & vbNewLine & _
                        "noMR: " & .Item("nomr") & vbNewLine & _
                        "noMRnoTelepon: " & .Item("nomrnoTelepon") & vbNewLine & _
                        "nama: " & .Item("nama") & vbNewLine & _
                        "nik: " & .Item("nik") & vbNewLine & _
                        "NoKartu: " & .Item("noKartu") & vbNewLine & _
                        "pisa: " & .Item("pisa") & vbNewLine & _
                        "provUmum: " & .Item("kodeProvUmum") & vbNewLine & _
                        "provuUmum: " & .Item("namaProvUmum") & vbNewLine & _
                        "jk: " & .Item("jk") & vbNewLine & _
                        "statusPeserta: " & .Item("namaStatusPeserta") & vbNewLine & _
                        "statusPesertaKode: " & .Item("kodeStatusPeserta") & vbNewLine & _
                        "tglCetakKartu: " & .Item("tglCetakKartu") & vbNewLine & _
                        "tglLahir: " & .Item("tglLahir") & vbNewLine & _
                        "tglTAT: " & .Item("tglTAT") & vbNewLine & _
                        "tglTMT: " & .Item("tglTMT") & vbNewLine & _
                        "umurSaatPelayanan: " & .Item("umurSaatPelayanan") & vbNewLine & "umurSekarang: " & .Item("umurSekarang")
    End With
End Sub
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, _
    (Screen.Height - Me.Height) / 2
    Set setting = New csetting
    Set vclaim = New vclaim
    setting.GetData
    'vclaim.ConsID = setting.ConsID
    'vclaim.SecretKey = setting.SecretKey
    'vclaim.AlamatWebService = setting.urlWebService
    Call centerForm(Me, MDIForm1)
    SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCariDataPeserta = Nothing
End Sub
