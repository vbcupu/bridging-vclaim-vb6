VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCreateSEP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUAT SEP"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   18750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   375
      Left            =   16800
      TabIndex        =   76
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateSEP 
      Caption         =   "BUAT SEP"
      Height          =   375
      Left            =   15000
      TabIndex        =   58
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CARI DATA PESERTA"
      Height          =   1215
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   18495
      Begin VB.TextBox txtCariNoRujukan 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Text            =   "110103030918Y000002"
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "CARI"
         Height          =   375
         Left            =   3120
         TabIndex        =   68
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton OptPCare 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PCARE"
         Height          =   375
         Left            =   1560
         TabIndex        =   67
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optRS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RUMAH SAKIT"
         Height          =   375
         Left            =   2640
         TabIndex        =   66
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NO RUJUKAN"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   18495
      Begin MSDataListLib.DataCombo dcKelasRawat 
         Height          =   315
         Left            =   1800
         TabIndex        =   85
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcJenisPelayanan 
         Height          =   315
         Left            =   1800
         TabIndex        =   84
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Poli Eksekutif"
         Height          =   195
         Left            =   1080
         TabIndex        =   83
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "KATARAK"
         Height          =   195
         Left            =   3600
         TabIndex        =   82
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "COB"
         Height          =   195
         Left            =   2760
         TabIndex        =   81
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox txtkdPoli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   79
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdCariPoli 
         Caption         =   "---"
         Height          =   255
         Left            =   5520
         TabIndex        =   78
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox txtnmPoli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   77
         Top             =   5400
         Width           =   3255
      End
      Begin VB.CommandButton cmdCariFaskes 
         Caption         =   "---"
         Height          =   255
         Left            =   9000
         TabIndex        =   75
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtnmPPKPelayanan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   74
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtNoMR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   72
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtnmDiagnosaAwal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   65
         Top             =   4680
         Width           =   3975
      End
      Begin VB.CommandButton cmdCariDiagnosa 
         Caption         =   "---"
         Height          =   255
         Left            =   6240
         TabIndex        =   59
         Top             =   4680
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpTglSEP 
         Height          =   255
         Left            =   1800
         TabIndex        =   50
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   116195329
         CurrentDate     =   43341
      End
      Begin VB.TextBox txtCatatan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   45
         Top             =   4320
         Width           =   7575
      End
      Begin VB.TextBox txtKdDiagAwal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   44
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11640
         TabIndex        =   38
         Top             =   6480
         Width           =   2895
      End
      Begin VB.TextBox txtNoTelpon 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11640
         TabIndex        =   37
         Top             =   6120
         Width           =   2895
      End
      Begin VB.TextBox txtPPkPelayananan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtNoKartu 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   360
         Width           =   4335
      End
      Begin VB.Frame Frame8 
         Caption         =   "SKDP"
         Height          =   1215
         Left            =   10320
         TabIndex        =   27
         Top             =   4800
         Width           =   5295
         Begin VB.TextBox txtNamaDokter 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   60
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton cmdCariDokter 
            Caption         =   "---"
            Height          =   195
            Left            =   4200
            TabIndex        =   39
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox txtKdDokter 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   36
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtNoSurat 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   35
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label27 
            Caption         =   "No Surat"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "KdDPJP"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Jaminan"
         Height          =   4575
         Left            =   10320
         TabIndex        =   14
         Top             =   120
         Width           =   7815
         Begin VB.OptionButton OptLakaTidak 
            Caption         =   "TIDAK"
            Height          =   255
            Left            =   2160
            TabIndex        =   54
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptLakaYa 
            Caption         =   "YA"
            Height          =   255
            Left            =   1200
            TabIndex        =   53
            Top             =   360
            Width           =   735
         End
         Begin VB.Frame Frame5 
            Caption         =   "Penjamin"
            Height          =   3735
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   7575
            Begin VB.TextBox txtKeterangan 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1680
               TabIndex        =   47
               Top             =   1080
               Width           =   5775
            End
            Begin VB.TextBox txtPenjamin 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1680
               TabIndex        =   46
               Top             =   360
               Width           =   4095
            End
            Begin VB.Frame Frame6 
               Caption         =   "Suplesi"
               Height          =   2295
               Left            =   240
               TabIndex        =   20
               Top             =   1320
               Width           =   7095
               Begin VB.CommandButton cmdCariKemungkinanSuplesi 
                  Caption         =   "Cari Suplesi"
                  Height          =   255
                  Left            =   3360
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.OptionButton OptSuplesiTdk 
                  Caption         =   "TIDAK"
                  Height          =   255
                  Left            =   2400
                  TabIndex        =   56
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton OptSuplesiYa 
                  Caption         =   "YA"
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   55
                  Top             =   240
                  Width           =   735
               End
               Begin VB.TextBox txtNoSuplesi 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   43
                  Top             =   600
                  Width           =   5175
               End
               Begin VB.Frame Frame7 
                  Caption         =   "Lokasi Laka"
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   23
                  Top             =   840
                  Width           =   6735
                  Begin VB.TextBox txtNamaKecamatan1 
                     Appearance      =   0  'Flat
                     Height          =   285
                     Left            =   2520
                     TabIndex        =   64
                     Top             =   960
                     Width           =   3015
                  End
                  Begin VB.TextBox txtKdKecamatan1 
                     Appearance      =   0  'Flat
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   63
                     Top             =   960
                     Width           =   1095
                  End
                  Begin VB.TextBox txtNamaKabupaten 
                     Appearance      =   0  'Flat
                     Height          =   285
                     Left            =   2520
                     TabIndex        =   62
                     Top             =   600
                     Width           =   3015
                  End
                  Begin VB.TextBox txtNamaPropinsi 
                     Appearance      =   0  'Flat
                     Height          =   285
                     Left            =   2520
                     TabIndex        =   61
                     Top             =   240
                     Width           =   3015
                  End
                  Begin VB.CommandButton cmdCariPropinsi 
                     Caption         =   "Cari"
                     Height          =   975
                     Left            =   5640
                     TabIndex        =   42
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.TextBox txtKdKabupaten 
                     Appearance      =   0  'Flat
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   41
                     Top             =   600
                     Width           =   1095
                  End
                  Begin VB.TextBox txtKdPropinsi 
                     Appearance      =   0  'Flat
                     Height          =   285
                     Left            =   1320
                     TabIndex        =   40
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.Label Label25 
                     Caption         =   "KdKecamatan"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   26
                     Top             =   960
                     Width           =   1095
                  End
                  Begin VB.Label Label24 
                     Caption         =   "KdKabupaten"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   25
                     Top             =   600
                     Width           =   1215
                  End
                  Begin VB.Label Label23 
                     Caption         =   "kdPropinsi"
                     Height          =   255
                     Left            =   240
                     TabIndex        =   24
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.Label Label22 
                  Caption         =   "No Suplesi"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   22
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label21 
                  Caption         =   "Suplesi"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   21
                  Top             =   240
                  Width           =   855
               End
            End
            Begin MSComCtl2.DTPicker dtpTglKejadian 
               Height          =   255
               Left            =   1680
               TabIndex        =   52
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   116195329
               CurrentDate     =   43341
            End
            Begin VB.Label Label20 
               Caption         =   "Keterangan"
               Height          =   255
               Left            =   240
               TabIndex        =   19
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label19 
               Caption         =   "TglKejadian"
               Height          =   255
               Left            =   240
               TabIndex        =   18
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "Penjamin"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Label Label18 
            Caption         =   "Lakalantas"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rujukan"
         Height          =   1815
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   8775
         Begin VB.TextBox txtNmPPKPerujuk 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   73
            Top             =   1440
            Width           =   5775
         End
         Begin VB.ComboBox cbAsalRujukan 
            Height          =   315
            Left            =   1560
            TabIndex        =   57
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtPPKRujukan 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   49
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtNoRujukan 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   48
            Top             =   1080
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker dtpTglRujukan 
            Height          =   255
            Left            =   1560
            TabIndex        =   51
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   116195329
            CurrentDate     =   43341
         End
         Begin VB.Label Label10 
            Caption         =   "PPK Rujukan"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "No Rujukan"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Tgl Rujukan"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Asal Rujukan"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label Label13 
         Caption         =   "POLI"
         Height          =   255
         Left            =   360
         TabIndex        =   80
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "User"
         Height          =   255
         Left            =   10440
         TabIndex        =   31
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "No telpon"
         Height          =   255
         Left            =   10440
         TabIndex        =   30
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Diag Awal"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Catatatan"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "No. MR"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Kls Rawat"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Jns Pelayanan"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PPK Pelayanan"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tgl SEP"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "No Kartu"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCreateSEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim
Dim setting As csetting
Private Sub cbFaskes_Change()

End Sub

Private Sub cbFaskes_GotFocus()
    
End Sub

Private Sub cbAsalRujukan_GotFocus()
    cbAsalRujukan.Clear
    cbAsalRujukan.AddItem "FASKES 1"
    cbAsalRujukan.AddItem "FASKES 2"
End Sub

Private Sub cbJenisPelayanan_GotFocus()
    cbJenisPelayanan.Clear
    cbJenisPelayanan.AddItem "Rawat Inap"
    cbJenisPelayanan.AddItem "Rawat Jalan"
End Sub

Private Sub cbKelasRawat_GotFocus()
    cbKelasRawat.Clear
    cbKelasRawat.AddItem "Kelas I"
    cbKelasRawat.AddItem "Kelas II"
    cbKelasRawat.AddItem "Kelas III"
End Sub

Private Sub Text15_Change()

End Sub

Private Sub cmdCari_Click()
    Call CariNoRujukan
End Sub

Private Sub cmdCariDiagnosa_Click()
    frmReferensiDiagnosa.FormPengirim = Me.Name
    frmReferensiDiagnosa.Show
End Sub

Private Sub cmdCariDokter_Click()
    frmRefDPJP.FormPengirim = Me.Name
    frmRefDPJP.Show
End Sub

Private Sub cmdCariFaskes_Click()
    frmCariFaskes.FormPengirim = Me.Name
    frmCariFaskes.Show
End Sub

Private Sub cmdCariPoli_Click()
    frmCariPoli.FormPengirim = Me.Name
    frmCariPoli.Show
End Sub

Private Sub cmdCariPropinsi_Click()
    frmWilayah.FormPengirim = Me.Name
    frmWilayah.Show
End Sub

Private Sub cmdCreateSEP_Click()
    Set vclaim = New vclaim
    Call vclaim.BuatSEP(txtNoKartu.Text, Format(dtpTglSEP.Value, "yyyy-MM-dd"), txtPPkPelayananan.Text, IIf(cbJenisPelayanan.Text = "Rawat Inap", "1", "2"), IIf(cbKelasRawat.Text = "Kelas 1", "1", IIf(cbKelasRawat.Text = "Kelas 2", "2", "3")), txtNoMR.Text, IIf(cbAsalRujukan.Text = "FASKES 1", "1", "2"), Format(dtpTglRujukan, "yyyy-MM-dd"), txtNoRujukan.Text, txtPPKRujukan.Text, txtcatatan.Text, txtKdDiagAwal.Text, txtKdPoli.Text, IIf(OptPoliEksekutifYa.Value = True, "1", "0"), IIf(OptCOBYA.Value = True, "1", "0"), IIf(OptKatarakYa.Value = True, "1", "0"), IIf(OptLakaYa.Value = True, "1", "0"), txtPenjamin.Text, IIf(OptLakaYa.Value = True, Format(dtpTglKejadian.Value, "yyyy-MM-dd"), ""), txtKeterangan.Text, IIf(OptSuplesiYa.Value = True, "1", "0"), txtNoSuplesi.Text, txtKdPropinsi.Text, txtKdKabupaten.Text, txtKdKecamatan1.Text, txtNoSurat.Text, txtKdDokter.Text, txtNoTelpon.Text, txtUser.Text)
End Sub
Private Sub CariNoRujukan()
    Set vclaim = New vclaim
    Set setting = New csetting
    Call setting.GetData
    
    vclaim.CariByRujukanPcare (txtCariNoRujukan.Text)
    If vclaim.ServerCode = "200" Then
    With vclaim.Hasil
        txtNoKartu.Text = .Item("nokartu")
        cbJenisPelayanan.Text = "Rawat Jalan"
        If .Item("kodehakkelas") = "1" Then
            cbKelasRawat.Text = "Kelas I"
        ElseIf .Item("kodehakkelas") = "2" Then
            cbKelasRawat.Text = "Kelas 2"
        Else
            cbKelasRawat.Text = "Kelas 3"
        End If
        
        txtNoMR.Text = IIf(IsNull(.Item("nomr")) = True, "", .Item("nomr"))
        cbAsalRujukan.Text = "FASKES 1"
        dtpTglRujukan.Value = CDate(.Item("tgldirujuk"))
        txtNoRujukan.Text = txtCariNoRujukan.Text
        txtPPkPelayananan.Text = setting.NoPPK
        txtnmPPKPelayanan.Text = setting.NamaPPK
        txtPPKRujukan.Text = .Item("kodeprovperujuk")
        txtNmPPKPerujuk.Text = .Item("namaprovperujuk")
        txtKdDiagAwal.Text = .Item("kodediagnosa")
        txtnmDiagnosaAwal.Text = .Item("namadiagnosa")
        txtKdPoli.Text = .Item("kodepolirujukan")
        txtnmPoli.Text = .Item("namapolirujukan")
        If IsNull(.Item("cobnmasuransi")) = True Then
            OptCOBTdk.Value = True
        Else
            OptCOBYA.Value = True
        End If
        txtNoTelpon.Text = ""
    End With
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set setting = New csetting
    dtpTglSEP.Value = Now
    txtPPkPelayananan.Text = setting.NoPPK
    txtnmPPKPelayanan.Text = setting.NamaPPK
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateSEP = Nothing
End Sub
