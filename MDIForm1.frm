VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Bridging VCLAIM"
   ClientHeight    =   6810
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11505
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuBerkas 
      Caption         =   "Referensi"
      Begin VB.Menu mnuCariPeserta 
         Caption         =   "Cari Peserta"
      End
      Begin VB.Menu mnuCariFaskes 
         Caption         =   "Cari Faskes"
      End
      Begin VB.Menu mnuCariPotensiSuplesi 
         Caption         =   "Cari Potensi Suplesi"
      End
      Begin VB.Menu mnuDiagnosa 
         Caption         =   "Diagnosa"
      End
      Begin VB.Menu mnuCariDokterDPJP 
         Caption         =   "Dokter DPJP"
      End
      Begin VB.Menu mnuWilayah 
         Caption         =   "Wilayah"
      End
      Begin VB.Menu mnuPropinsi 
         Caption         =   "Propinsi"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuRujukan 
      Caption         =   "RUJUKAN"
      Begin VB.Menu mnurujukankeluar 
         Caption         =   "Rujukan keluar"
      End
      Begin VB.Menu mnuRujukankeRS 
         Caption         =   "Rujukan ke RS"
      End
   End
   Begin VB.Menu mnuSEP 
      Caption         =   "SEP"
      Begin VB.Menu mnuCreateSEP 
         Caption         =   "CREATE SEP"
      End
      Begin VB.Menu mnuDetailSEP 
         Caption         =   "Detail SEP"
      End
      Begin VB.Menu mnuRiwayatBridging 
         Caption         =   "Riwayat Bridging"
      End
   End
   Begin VB.Menu mnuMonitoring 
      Caption         =   "Monitoring"
      Begin VB.Menu mnuDataKunjungan 
         Caption         =   "Data Kunjungan"
      End
      Begin VB.Menu mnuDataKlaim 
         Caption         =   "Data Klaim"
      End
      Begin VB.Menu mnuHistoryPasien 
         Caption         =   "History Pasien"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Call getKoneksi
End Sub

Private Sub mnuCariDokterDPJP_Click()
    frmRefDPJP.Show
End Sub

Private Sub mnuCariFaskes_Click()
    frmCariFaskes.Show
End Sub

Private Sub mnuCariPeserta_Click()
    frmCariDataPeserta.Show
End Sub

Private Sub mnuCariPotensiSuplesi_Click()
    frmCariPotensiSuplesi.Show
End Sub

Private Sub mnuCreateSEP_Click()
    'frmCreateSEP.Show
    frmUbahJenisPasienBPJS.Show
End Sub

Private Sub mnuDataKlaim_Click()
    frmDataKlaim.Show
End Sub

Private Sub mnuDataKunjungan_Click()
    frmdatakunjungan.Show
End Sub

Private Sub mnuDetailSEP_Click()
    frmDetailSEP.Show
End Sub

Private Sub mnuDiagnosa_Click()
    frmReferensiDiagnosa.Show
End Sub

Private Sub mnuHistoryPasien_Click()
    frmHistoryPasien.Show
End Sub

Private Sub mnuPropinsi_Click()
    frmPropinsi.Show
End Sub

Private Sub mnuRiwayatBridging_Click()
    frmRiwayatBridging.Show
End Sub

Private Sub mnurujukankeluar_Click()
    frmBuatRujukanBPJS.Show
End Sub

Private Sub mnuRujukankeRS_Click()
    frmrujukanbpjs.Show
End Sub

Private Sub mnuWilayah_Click()
    frmWilayah.Show
End Sub
