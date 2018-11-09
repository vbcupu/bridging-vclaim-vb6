VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vclaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' Class:        vclaim
' FILENAME:     vclaim.cls
' AUTHOR:       Martin Eko Setiawan
' CREATED:      12/09/2018
' COPYRIGHT:    @2018 martin eko setiawan
'
' DESCRIPTION:
' Class Setting data untuk aplikasi bridging VClaim dengan visual basic 6
' message. Baca keseluruhan class ini untuk lebih detailnya
' Saya berbagi class ini bukan karena karena saya ingin berbagi kepada teman-teman programmer rumah sakit yang kesulitan dalam bridging vb6 dengan vclaim
' Pada dasarnya saya masih belajar dan terus belajar, jadi maaf bila source ini masih banyak kekurangan, acak-acakan dan banyak bug-nya

' Ini adalah class yang gratis tis, tanpa biaya, Dengan Catatan:
' 1. Kode ini boleh didistribusikan ulang ke source code dan aplikasi lain, namun jangan hapus author, silahkan tambahkan nama anda di author.
' 2. Class ini hanya rintisan, minimal dengan modal class ini, kamu bisa mengembangkan modul lainnya karena prinsip kerja baik request maupun pengiriman parameter hampir sama.
' 3. Menegaskan kembali. Tolong author nama saya jangan dihapus ya, kalau mau ditambah nggak apa-apa, kasih aja di samping/bawah/atas nama saya.
' 4. Karena Source Code ini saya berikan dengan gratis, maka saya berharap source code ini tidak dipakai oleh perusahaan developer berbayar. Karena perusahaan developer harusnya bisa membuat/meresearch API sendiri.
' 5. Saya tidak menjamin aplikasi ini bebas bug, tapi aplikasi ini saya pakai di kerjaan juga, jadi InsyaAllah bebas bug.
' 6. Sekali lagi saya berharap anda menghormati tentang hak cipta meski tidak sy daftarke ke panitia hak cipta, dan tidak ada konsekwensi apapun jika anda mengakui sebagai karya anda.


' Bila kamu mengalami kesulitan dalam implementasi class ini, silahkan berkonsultasi kapan saja, dengan catatan:
' 1. Jangan meminta jawaban secepatnya, karena mungkin saya sedang sibuk.
' 2. Jangan meminta saya bisa menyelesaikan masalah kamu, kalau mungkin masalahnya tidak bisa saya selesaikan. Bagaimanapun saya cuma programmer biasa yang masih kelas cupu
' 3. Class ini hanya rintisan, minimal dengan modal class ini, kamu bisa mengembangkan modul lainnya karena prinsip kerja baik request maupun pengiriman parameter hampir sama


' Web Site:  www.martinsetiawan.com
' E-mail:vbcupu@gmail.com/martinsetiawan@gmail.com
' Whatsapp: Silahkan lihat di web

'*******************************************************************************


Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Public Url As String
Public EndPoint As String
Public Method As String
Public status As String
Public ResponseText As String
Public ConsId As String
Public SecretKey As String
Public RequestJson As String

'-----------------------------------------------------------------



Public HasilJson As String
Public Hasil As Dictionary
Public HasilKode As Dictionary
Public HasilKet As Dictionary
Public potsuplesiNoRegister As Dictionary
Public potsuplesinoSep As Dictionary
Public potsuplesinoSepAwal As Dictionary
Public potsuplesinoSuratJaminan As Dictionary
Public potsuplesitglKejadian As Dictionary
Public potsuplesitglSep As Dictionary
Public ServerMessage As String
Public ServerCode As String

'Simpan Data Rujukan
Public RTglRujukan As Dictionary
Public RNoPeserta As Dictionary
Public RNoRujukan As Dictionary
Public RNoCM As Dictionary
Public RNama As Dictionary
Public RKdPoli As Dictionary
Public RNmPoli As Dictionary
Public RKodeDiagnosa As Dictionary
Public RNamaDiagnosa As Dictionary
Public RNoPPKPerujuk As Dictionary
Public RNmPPkPerujuk As Dictionary

'VARIABLE MONITORING
'Riwayat histori pasien
Public historysepDiagnosa As Dictionary
Public historysepJnspelayanan As Dictionary
Public historysepKelasrawat As Dictionary
Public historysepNamapeserta As Dictionary
Public historysepNokartu As Dictionary
Public historysepNosep As Dictionary
Public historysepNorujukan As Dictionary
Public historysepPoli As Dictionary
Public historysepPpkpelayanan As Dictionary
Public historysepTglpulangsep As Dictionary
Public historysepTglsep As Dictionary

'Data Klaim
Public dataklaimKodeinacbg As Dictionary
Public dataklaimNamainacbg As Dictionary
Public dataklaimBiayabypengajuan As Dictionary
Public dataklaimBiayabySetujui As Dictionary
Public dataklaimBiayabyTarifGruper As Dictionary
Public dataklaimBiayabyTarifRS As Dictionary
Public dataklaimBiayabyTopup As Dictionary
Public dataklaimKelasrawat As Dictionary
Public dataklaimNoFPK As Dictionary
Public dataklaimNosep As Dictionary
Public dataklaimPesertanama As Dictionary
Public dataklaimPesertanokartu As Dictionary
Public dataklaimPesertanomr As Dictionary
Public dataklaimPoli As Dictionary
Public dataklaimstatus As Dictionary
Public dataklaimTglpulang As Dictionary
Public dataklaimTglsep As Dictionary






Private req As callws
Private Result As String
Private p As Object
Private i As Integer




'############################ CARI PASIEN ############################################################################################
'############################ CARI PASIEN ############################################################################################
Public Function CariPesertaByNoKartu(nokartu As String)
   Call panggil(req, "Peserta/nokartu/" & nokartu & "/tglSEP/" & Format(Now, "yyyy-MM-dd"))
   Me.status = req.status
   If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        Set Hasil = New Dictionary
        With Hasil
            .Add "cobNmAsuransi", p.Item("response").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobKdAsuransi", p.Item("response").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobTglTAT", p.Item("response").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobTglTMT", p.Item("response").Item("peserta").Item("cob").Item("tglTMT")
            .Add "namaHakKelas", p.Item("response").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "kodeHakKelas", p.Item("response").Item("peserta").Item("hakKelas").Item("kode")
            .Add "informasiDinsos", p.Item("response").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "namaJenisPeserta", p.Item("response").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "kodeJenisPeserta", p.Item("response").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "nomr", p.Item("response").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrnoTelepon", p.Item("response").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("peserta").Item("pisa")
            .Add "kodeProvUmum", p.Item("response").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "namaProvUmum", p.Item("response").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("peserta").Item("sex")
            .Add "namaStatusPeserta", p.Item("response").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "kodeStatusPeserta", p.Item("response").Item("peserta").Item("statusPeserta").Item("kode")
            .Add "tglCetakKartu", p.Item("response").Item("peserta").Item("tglCetakKartu")
            .Add "tglLahir", p.Item("response").Item("peserta").Item("tglLahir")
            .Add "tglTAT", p.Item("response").Item("peserta").Item("tglTAT")
            .Add "tglTMT", p.Item("response").Item("peserta").Item("tglTMT")
            .Add "umurSaatPelayanan", p.Item("response").Item("peserta").Item("umur").Item("umurSaatPelayanan")
            .Add "umurSekarang", p.Item("response").Item("peserta").Item("umur").Item("umurSekarang")
        End With
   Else
       Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
   End If
End Function
Public Function CariPesertaByNIK(NIK As String)
   Call panggil(req, "Peserta/nik/" & NIK & "/tglSEP/" & Format(Now, "yyyy-MM-dd"))
   Me.status = req.status
   If req.status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "cobNmAsuransi", p.Item("response").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobNoAsuransi", p.Item("response").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobTglTAT", p.Item("response").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobTglTMT", p.Item("response").Item("peserta").Item("cob").Item("tglTMT")
            .Add "namaHakKelas", p.Item("response").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "kodeHakKelas", p.Item("response").Item("peserta").Item("hakKelas").Item("kode")
            .Add "informasiDinsos", p.Item("response").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "keteranganjenispeserta", p.Item("response").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "kodejenispeserta", p.Item("response").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "nomr", p.Item("response").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrnoTelepon", p.Item("response").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("peserta").Item("pisa")
            .Add "kodeProvUmum", p.Item("response").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "namaProvUmum", p.Item("response").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("peserta").Item("sex")
            .Add "namaStatusPeserta", p.Item("response").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "kodeStatusPeserta", p.Item("response").Item("peserta").Item("statusPeserta").Item("kode")
            .Add "tglCetakKartu", p.Item("response").Item("peserta").Item("tglCetakKartu")
            .Add "tglLahir", p.Item("response").Item("peserta").Item("tglLahir")
            .Add "tglTAT", p.Item("response").Item("peserta").Item("tglTAT")
            .Add "tglTMT", p.Item("response").Item("peserta").Item("tglTMT")
            .Add "umurSaatPelayanan", p.Item("response").Item("peserta").Item("umur").Item("umurSaatPelayanan")
            .Add "umurSekarang", p.Item("response").Item("peserta").Item("umur").Item("umurSekarang")
        End With
    Else
       Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function CariByRujukanPcare(NoRujukan As String)
   Call panggil(req, "Rujukan/" & NoRujukan)
   Me.status = req.status
   If req.status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "kodeDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("kode")
            .Add "namaDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("nama")
            .Add "keluhan", p.Item("response").Item("rujukan").Item("keluhan")
            .Add "noKunjungan", p.Item("response").Item("rujukan").Item("noKunjungan")
            .Add "kodePelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("kode")
            .Add "namaPelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("nama")
            .Add "cobNmAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobNoAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobTgltat", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobTgltmt", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTMT")
            .Add "kodeHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("kode")
            .Add "namaHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "informasiDinsos", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "kodeJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "namaJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "nomr", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrTelepon", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("rujukan").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("rujukan").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("rujukan").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("rujukan").Item("peserta").Item("pisa")
            .Add "kodeProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "namaProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("rujukan").Item("peserta").Item("sex")
            .Add "namaStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "kodeStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("kode")
            .Add "tglCetakKartu", p.Item("response").Item("rujukan").Item("peserta").Item("tglCetakKartu")
            .Add "tglLahir", p.Item("response").Item("rujukan").Item("peserta").Item("tglLahir")
            .Add "tglTAT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTAT")
            .Add "tglTMT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTMT")
            .Add "umurSaatPelayanan", p.Item("response").Item("rujukan").Item("peserta").Item("umurSaatPelayanan")
            .Add "umurSekarang", p.Item("response").Item("rujukan").Item("peserta").Item("umurSekarang")
            .Add "kodePoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("kode")
            .Add "namaPoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("nama")
            .Add "kodeProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("kode")
            .Add "namaProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("nama")
            .Add "tglDirujuk", p.Item("response").Item("rujukan").Item("tglKunjungan")
       End With
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function CariByRujukanRS(NoRujukan As String)
   Call panggil(req, "/Rujukan/RS/" & NoRujukan)
   Me.status = req.status
   If req.status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "kodeDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("kode")
            .Add "namaDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("nama")
            .Add "keluhan", p.Item("response").Item("rujukan").Item("keluhan")
            .Add "noKunjungan", p.Item("response").Item("rujukan").Item("noKunjungan")
            .Add "kodePelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("kode")
            .Add "namaPelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("nama")
            .Add "cobNmAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobNoAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobTglTAT", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobTglTMT", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTMT")
            .Add "kodeHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("kode")
            .Add "namaHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "informasiDinsos", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "kodeJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "namaJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "nomr", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrTelepon", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("rujukan").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("rujukan").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("rujukan").Item("peserta").Item("nokartu")
            .Add "pisa", p.Item("response").Item("rujukan").Item("peserta").Item("pisa")
            .Add "kodeProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "namaProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("rujukan").Item("peserta").Item("sex")
            .Add "namaStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "kodeStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("kode")
            .Add "tglCetakKartu", p.Item("response").Item("rujukan").Item("peserta").Item("tglCetakKartu")
            .Add "tglLahir", p.Item("response").Item("rujukan").Item("peserta").Item("tglLahir")
            .Add "tglTAT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTAT")
            .Add "tglTMT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTMT")
            .Add "umurSaatPelayanan", p.Item("response").Item("rujukan").Item("peserta").Item("umurSaatPelayanan")
            .Add "umurSekarang", p.Item("response").Item("rujukan").Item("peserta").Item("umurSekarang")
            .Add "kodePoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("kode")
            .Add "namaPoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("nama")
            .Add "kodeProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("kode")
            .Add "namaProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("nama")
            .Add "tglDirujuk", p.Item("response").Item("rujukan").Item("tglKunjungan")
       End With
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function CariByRujukanByNoKartu1recPcare(nokartu As String)
   Call panggil(req, "Rujukan/Peserta/" & nokartu)
   Me.status = req.status
   If req.status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "kodeDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("kode")
            .Add "namaDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("nama")
            .Add "keluhan", p.Item("response").Item("rujukan").Item("keluhan")
            .Add "noKunjungan", p.Item("response").Item("rujukan").Item("noKunjungan")
            .Add "kodePelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("kode")
            .Add "namaPelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("nama")
            .Add "cobNmAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobNoAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobTglTAT", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobTglTMT", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTMT")
            .Add "kodeHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("kode")
            .Add "namaHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "informasiDinsos", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "kodeJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "namaJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "nomr", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrTelepon", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("rujukan").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("rujukan").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("rujukan").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("rujukan").Item("peserta").Item("pisa")
            .Add "kodeProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "namaProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("rujukan").Item("peserta").Item("sex")
            .Add "namaStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "kodeStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("kode")
            .Add "tglCetakKartu", p.Item("response").Item("rujukan").Item("peserta").Item("tglCetakKartu")
            .Add "tglLahir", p.Item("response").Item("rujukan").Item("peserta").Item("tglLahir")
            .Add "tglTAT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTAT")
            .Add "tglTMT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTMT")
            .Add "umurSaatPelayanan", p.Item("response").Item("rujukan").Item("peserta").Item("umurSaatPelayanan")
            .Add "umurSekarang", p.Item("response").Item("rujukan").Item("peserta").Item("umurSekarang")
            .Add "kodePoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("kode")
            .Add "namaRoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("nama")
            .Add "kodeProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("kode")
            .Add "namaProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("nama")
            .Add "tglDirujuk", p.Item("response").Item("rujukan").Item("tglKunjungan")
       End With
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function CariByRujukanByNoKartu1recRS(nokartu As String)
   Call panggil(req, "Rujukan/RS/Peserta/" & nokartu)
   Me.status = req.status
   If req.status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "kodeDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("kode")
            .Add "namaDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("nama")
            .Add "keluhan", p.Item("response").Item("rujukan").Item("keluhan")
            .Add "noKunjungan", p.Item("response").Item("rujukan").Item("noKunjungan")
            .Add "kodePelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("kode")
            .Add "namaPelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("nama")
            .Add "cobNmAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobNoAsuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobTglTAT", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobTglTMT", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTMT")
            .Add "kodeHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("kode")
            .Add "namaHakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "informasiFinsos", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "kodeJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "namaJenisPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "nomr", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrtelepon", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("rujukan").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("rujukan").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("rujukan").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("rujukan").Item("peserta").Item("pisa")
            .Add "kodeProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "namaProvUmum", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("rujukan").Item("peserta").Item("sex")
            .Add "namaStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "kodeStatusPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("kode")
            .Add "tglCetakKartu", p.Item("response").Item("rujukan").Item("peserta").Item("tglCetakKartu")
            .Add "tglLahir", p.Item("response").Item("rujukan").Item("peserta").Item("tglLahir")
            .Add "tglTAT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTAT")
            .Add "tglTMT", p.Item("response").Item("rujukan").Item("peserta").Item("tglTMT")
            .Add "umurSaatPelayanan", p.Item("response").Item("rujukan").Item("peserta").Item("umurSaatPelayanan")
            .Add "umurSekarang", p.Item("response").Item("rujukan").Item("peserta").Item("umurSekarang")
            .Add "kodePoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("kode")
            .Add "namaPoliRujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("nama")
            .Add "kodeProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("kode")
            .Add "namaProvPerujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("nama")
            .Add "tglDirujuk", p.Item("response").Item("rujukan").Item("tglKunjungan")
       End With
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
'################################ END ##################################################################
'#######################################################################################################

'############################### REFERENSI #############################################################
'#######################################################################################################
Public Function refPoli(NamaPoli As String)
  Call panggil(req, "referensi/poli/" & NamaPoli)
  Me.status = req.status
  If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        i = 1
        Set Me.HasilKode = New Dictionary
        Set Me.HasilKet = New Dictionary
        For i = 1 To p("response")("poli").Count
           HasilKode.Add i, p("response")("poli")(i)("kode")
           HasilKet.Add i, p("response")("poli")(i)("nama")
         Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
   End If
End Function
Public Function CariFaskes(NamaFaskes, Faskes As String)
    Call panggil(req, "referensi/faskes/" & NamaFaskes & "/" & Faskes)
    Me.status = req.status
    If req.status = "200" Then
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       i = 1
       Set Me.HasilKode = New Dictionary
       Set Me.HasilKet = New Dictionary
       For i = 1 To p("response")("faskes").Count
           Me.HasilKode.Add i, p("response")("faskes")(i)("kode")
           Me.HasilKet.Add i, p("response")("faskes")(i)("nama")
        Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function CariDiagnosa(NamaDiagnosa As String)
    Call panggil(req, "/referensi/diagnosa/" & NamaDiagnosa)
    Me.status = req.status
    If req.status = "200" Then
        Set p = JSON.parse(req.ResponseText)
        Me.HasilJson = req.ResponseText
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        i = 1
        Set Me.HasilKode = New Dictionary
        Set Me.HasilKet = New Dictionary
        For i = 1 To p("response")("diagnosa").Count
            Me.HasilKode.Add i, p("response")("diagnosa")(i)("kode")
            Me.HasilKet.Add i, p("response")("diagnosa")(i)("nama")
        Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function refDokterDPJP(JenisPelayanan As String, TglPelayanan As String, Spesialis As String)
    Call panggil(req, "/referensi/dokter/pelayanan/" & JenisPelayanan & "/tglPelayanan/" & TglPelayanan & "/Spesialis/" & Spesialis)
    Me.status = req.status
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        i = 1
        Set Me.HasilKode = New Dictionary
        Set Me.HasilKet = New Dictionary
        For i = 1 To p("response")("list").Count
            HasilKode.Add i, p("response")("list")(i)("kode")
            HasilKet.Add i, p("response")("list")(i)("nama")
        Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function refSpesialis()
    Call panggil(req, "/referensi/spesialistik")
    Me.status = req.status
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        i = 1
        Set Me.HasilKode = New Dictionary
        Set Me.HasilKet = New Dictionary
        For i = 1 To p("response")("list").Count
            Me.HasilKode.Add i, p("response")("list")(i)("kode")
            Me.HasilKet.Add i, p("response")("list")(i)("nama")
        Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function refPropinsi()
    Call panggil(req, "/referensi/propinsi")
    Me.status = req.status
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        i = 1
        Set Me.HasilKode = New Dictionary
        Set Me.HasilKet = New Dictionary
        For i = 1 To p("response")("list").Count
            Me.HasilKode.Add i, p("response")("list")(i)("kode")
            Me.HasilKet.Add i, p("response")("list")(i)("nama")
        Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function refkota(kdPropinsi As String)
    Call panggil(req, "/referensi/kabupaten/propinsi/" & kdPropinsi)
    Me.status = req.status
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        i = 1
        Set Me.HasilKode = New Dictionary
        Set Me.HasilKet = New Dictionary
        For i = 1 To p("response")("list").Count
            Me.HasilKode.Add i, p("response")("list")(i)("kode")
            Me.HasilKet.Add i, p("response")("list")(i)("nama")
        Next i
    Else
         Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function refkecamatan(kdKota As String)
    Call panggil(req, "/referensi/kecamatan/kabupaten/" & kdKota)
    Me.status = req.status
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        i = 1
        Set Me.HasilKet = New Dictionary
        Set Me.HasilKode = New Dictionary
        For i = 1 To p("response")("list").Count
            Me.HasilKode.Add i, p("response")("list")(i)("kode")
            Me.HasilKet.Add i, p("response")("list")(i)("nama")
        Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
'################################ END ##################################################################
'#######################################################################################################

'################################ SEP ##################################################################
'#######################################################################################################
Public Function DetailSEP(NoSEP As String)
    Call panggil(req, "SEP/" & NoSEP)
    Me.status = req.status
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        Set Me.Hasil = New Dictionary
        With Me.Hasil
            .Add "catatan", p.Item("response").Item("catatan")
            .Add "diagnosa", p.Item("response").Item("diagnosa")
            .Add "jnsPelayanan", p.Item("response").Item("jnsPelayanan")
            .Add "kelasrawat", p.Item("response").Item("kelasRawat")
            .Add "nosep", p.Item("response").Item("noSep")
            .Add "penjamin", p.Item("response").Item("penjamin")
            .Add "pesertaasuransi", p.Item("response").Item("peserta").Item("asuransi")
            .Add "pesertahakkelas", p.Item("response").Item("peserta").Item("hakKelas")
            .Add "pesertajnspeserta", p.Item("response").Item("peserta").Item("jnsPeserta")
            .Add "pesertakelamin", p.Item("response").Item("peserta").Item("kelamin")
            .Add "pesertanama", p.Item("response").Item("peserta").Item("nama")
            .Add "pesertanokartu", p.Item("response").Item("peserta").Item("noKartu")
            .Add "pesertanomr", p.Item("response").Item("peserta").Item("noMr")
            .Add "pesertatgllahir", p.Item("response").Item("peserta").Item("tglLahir")
            .Add "poli", p.Item("response").Item("poli")
            .Add "polieksekutif", p.Item("response").Item("poliEksekutif")
            .Add "tglsep", p.Item("response").Item("tglSep")
        End With
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
    
End Function
Public Function BuatSEP(nokartu, tglSEP, PPKPelayanan, jnspelayanan, klsrawat, nomr, asalRujukan, tglrujukan, NoRujukan, ppkRujukan, catatan, diagawal, politujuan, polieksekutif, cob, katarak, jamlakaLantas, jampenjamin, jamtglKejadian, jamketerangan, jamsuplesi, jamnoSepSuplesi, jamlokasilakakdPropinsi, jamlokasilakakdKabupaten, jamlokasilakakdKecamatan, skdpnosurat, skdpkodeDPJP, noTelp, user As String)
Dim prm As String
Dim JsonRequest As String

    If jamlakaLantas = "0" Then
      jamsuplesi = ""
    End If

'JsonRequest = "{""request"":{""t_sep"":{""noKartu"":""" & nokartu & """,""tglSep"":""" & TglSEP & """ ,""ppkPelayanan"":""" & PPKPelayanan & """,""jnsPelayanan"":""" & jnspelayanan & """ ,""klsRawat"":""" & KlsRawat & """,""noMR"":""" & NoMr & """ ,""rujukan"":{""asalRujukan"":""" & asalRujukan & """,""tglRujukan"":""" & tglrujukan & """ ,""noRujukan"":""" & NoRujukan & """,""ppkRujukan"":""" & ppkRujukan & """},""catatan"":""" & Catatan & """,""diagAwal"":""" & DiagAwal & """,""poli"":{""tujuan"":""" & PoliTujuan & """,""eksekutif"":""" & polieksekutif & """},""cob"":{""cob"":""" & cob & """},""katarak"":{""katarak"":""" & katarak & """},""jaminan"":{""lakaLantas"":""" & jamlakaLantas & """,""penjamin"":{""penjamin"":" & _
'           """" & jampenjamin & """,""tglKejadian"":""" & jamtglKejadian & """,""keterangan"":""" & jamketerangan & """ ,""suplesi"":{""suplesi"":""" & IIf(jamsuplesi = "", 0, jamsuplesi) & """ ,""noSepSuplesi"":""" & jamnoSepSuplesi & """ ,""lokasiLaka"":{""kdPropinsi"":""" & jamlokasilakakdPropinsi & """ ,""kdKabupaten"":""" & jamlokasilakakdKabupaten & """,""kdKecamatan"":""" & jamlokasilakakdKecamatan & """}}}},""skdp"":{""noSurat"":""" & skdpnoSurat & """ ,""kodeDPJP"":""" & skdpkodeDPJP & """},""noTelp"":""" & noTelp & """,""user"":""" & user & """}}}"
           

    prm = "{" & Chr(34) & "request" & Chr(34) & ":{"
    prm = prm & Chr(34) & "t_sep" & Chr(34) & ":{"
    prm = prm & Chr(34) & "noKartu" & Chr(34) & ":" & Chr(34) & nokartu & Chr(34) & ","
    prm = prm & Chr(34) & "tglSep" & Chr(34) & ":" & Chr(34) & tglSEP & Chr(34) & ","
    prm = prm & Chr(34) & "ppkPelayanan" & Chr(34) & ":" & Chr(34) & PPKPelayanan & Chr(34) & ","
    prm = prm & Chr(34) & "jnsPelayanan" & Chr(34) & ":" & Chr(34) & jnspelayanan & Chr(34) & ","
    prm = prm & Chr(34) & "klsRawat" & Chr(34) & ":" & Chr(34) & klsrawat & Chr(34) & ","
    prm = prm & Chr(34) & "noMR" & Chr(34) & ":" & Chr(34) & nomr & Chr(34) & ","
    prm = prm & Chr(34) & "rujukan" & Chr(34) & ":{"
        prm = prm & Chr(34) & "asalRujukan" & Chr(34) & ":" & Chr(34) & asalRujukan & Chr(34) & ","
        prm = prm & Chr(34) & "tglRujukan" & Chr(34) & ":" & Chr(34) & tglrujukan & Chr(34) & ","
        prm = prm & Chr(34) & "noRujukan" & Chr(34) & ":" & Chr(34) & NoRujukan & Chr(34) & ","
        prm = prm & Chr(34) & "ppkRujukan" & Chr(34) & ":" & Chr(34) & ppkRujukan & Chr(34)
    prm = prm & "},"
    prm = prm & Chr(34) & "catatan" & Chr(34) & ":" & Chr(34) & catatan & Chr(34) & ","
    prm = prm & Chr(34) & "diagAwal" & Chr(34) & ":" & Chr(34) & diagawal & Chr(34) & ","
    prm = prm & Chr(34) & "poli" & Chr(34) & ":{"
        prm = prm & Chr(34) & "tujuan" & Chr(34) & ":" & Chr(34) & politujuan & Chr(34) & ","
        prm = prm & Chr(34) & "eksekutif" & Chr(34) & ":" & Chr(34) & polieksekutif & Chr(34)
    prm = prm & "},"
    prm = prm & Chr(34) & "cob" & Chr(34) & ":{"
        prm = prm & Chr(34) & "cob" & Chr(34) & ":" & Chr(34) & cob & Chr(34)
    prm = prm & "},"
    prm = prm & Chr(34) & "katarak" & Chr(34) & ":{"
        prm = prm & Chr(34) & "katarak" & Chr(34) & ":" & Chr(34) & katarak & Chr(34)
    prm = prm & "},"
    prm = prm & Chr(34) & "jaminan" & Chr(34) & ":{"
        prm = prm & Chr(34) & "lakaLantas" & Chr(34) & ":" & Chr(34) & jamlakaLantas & Chr(34) & ","
        prm = prm & Chr(34) & "penjamin" & Chr(34) & ":{"
            prm = prm & Chr(34) & "penjamin" & Chr(34) & ":" & Chr(34) & jampenjamin & Chr(34) & ","
            prm = prm & Chr(34) & "tglKejadian" & Chr(34) & ":" & Chr(34) & jamtglKejadian & Chr(34) & ","
            prm = prm & Chr(34) & "keterangan" & Chr(34) & ":" & Chr(34) & jamketerangan & Chr(34) & ","
            prm = prm & Chr(34) & "suplesi" & Chr(34) & ":{"
                prm = prm & Chr(34) & "suplesi" & Chr(34) & ":" & Chr(34) & IIf(jamsuplesi = "", 0, jamsuplesi) & Chr(34) & ","
                prm = prm & Chr(34) & "noSepSuplesi" & Chr(34) & ":" & Chr(34) & jamnoSepSuplesi & Chr(34) & ","
                prm = prm & Chr(34) & "lokasiLaka" & Chr(34) & ":{"
                    prm = prm & Chr(34) & "kdPropinsi" & Chr(34) & ":" & Chr(34) & jamlokasilakakdPropinsi & Chr(34) & ","
                    prm = prm & Chr(34) & "kdKabupaten" & Chr(34) & ":" & Chr(34) & jamlokasilakakdKabupaten & Chr(34) & ","
                    prm = prm & Chr(34) & "kdKecamatan" & Chr(34) & ":" & Chr(34) & jamlokasilakakdKecamatan & Chr(34)
                prm = prm & "}"
            prm = prm & "}"
        prm = prm & "}"
    prm = prm & "},"
    prm = prm & Chr(34) & "skdp" & Chr(34) & ":{"
        prm = prm & Chr(34) & "noSurat" & Chr(34) & ":" & Chr(34) & skdpnosurat & Chr(34) & ","
        prm = prm & Chr(34) & "kodeDPJP" & Chr(34) & ":" & Chr(34) & skdpkodeDPJP & Chr(34)
    prm = prm & "},"
    prm = prm & Chr(34) & "noTelp" & Chr(34) & ":" & Chr(34) & noTelp & Chr(34) & ","
    prm = prm & Chr(34) & "user" & Chr(34) & ":" & Chr(34) & user & Chr(34)
    prm = prm & "}"
    prm = prm & "}"
    prm = prm & "}"

   Me.RequestJson = prm
   Call panggil(req, "SEP/1.1/insert", "POST", prm, True)
   
   Me.status = req.status
   If req.status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & ServerCode & vbNewLine & " Keterangan " & ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
        Else
           Set Hasil = New Dictionary
           With Hasil
                .Add "catatan", p.Item("response").Item("sep").Item("catatan")
                .Add "diagnosa", p.Item("response").Item("sep").Item("diagnosa")
                .Add "jnspelayanan", p.Item("response").Item("sep").Item("jnsPelayanan")
                .Add "kelasrawat", p.Item("response").Item("sep").Item("kelasRawat")
                .Add "nosep", p.Item("response").Item("sep").Item("noSep")
                .Add "penjamin", p.Item("response").Item("sep").Item("penjamin")
                .Add "asuransi", p.Item("response").Item("sep").Item("peserta").Item("asuransi")
                .Add "hakkelas", p.Item("response").Item("sep").Item("peserta").Item("hakKelas")
                .Add "jnspeserta", p.Item("response").Item("sep").Item("peserta").Item("jnsPeserta")
                .Add "kelamin", p.Item("response").Item("sep").Item("peserta").Item("kelamin")
                .Add "nama", p.Item("response").Item("sep").Item("peserta").Item("nama")
                .Add "nokartu", p.Item("response").Item("sep").Item("peserta").Item("noKartu")
                .Add "nomr", p.Item("response").Item("sep").Item("peserta").Item("noMr")
                .Add "tgllahir", p.Item("response").Item("sep").Item("peserta").Item("tglLahir")
                .Add "dinsos", p.Item("response").Item("sep").Item("informasi").Item("Dinsos")
                .Add "prolanisprb", p.Item("response").Item("sep").Item("informasi").Item("prolanisPRB")
                .Add "nosktm", p.Item("response").Item("sep").Item("informasi").Item("noSKTM")
                .Add "poli", p.Item("response").Item("sep").Item("poli")
                .Add "polieksekutif", p.Item("response").Item("polieksekutif")
                .Add "tglsep", p.Item("response").Item("sep").Item("tglsep")
           End With
       End If
    Else
        Call MsgBox("Error Ndaaa")
   End If
End Function
Public Function HapusSEP(NoSEP As String, user As String)
    Dim prm As String
    
    prm = "{" & Chr(34) & "request" & Chr(34) & ":{"
    prm = prm & Chr(34) & "t_sep" & Chr(34) & ":{"
    prm = prm & Chr(34) & "noSep" & Chr(34) & ":" & Chr(34) & NoSEP & Chr(34) & ","
    prm = prm & Chr(34) & "user" & Chr(34) & ":" & Chr(34) & Left(Replace(user, " ", ""), 20) & Chr(34)
    prm = prm & "}"
    prm = prm & "}"
    prm = prm & "}"
    
    
    Me.RequestJson = prm
    Call panggil(req, "SEP/Delete", "DELETE", prm)
    Me.status = req.status
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode = "200" Then
            Set Hasil = New Dictionary
            With Hasil
                .Add "nosep", p.Item("response")
            End With
        Else
            Call MsgBox("Respon server: Kode: " & Me.ServerCode & " ;Pesan: " & Me.ServerMessage, vbOKOnly, "WARNING")
        End If
    Else
        Call MsgBox(req.status & " " & req.ResponseText, vbOKOnly, "WARNING")
    End If
End Function
Public Function CariPotensiSuplesi(nokartu, tanggal As String)
    Dim jml As Integer
    Dim Urlnya As String
    
    Urlnya = "/sep/JasaRaharja/Suplesi/" & nokartu & "/tglPelayanan/" & tanggal
    Call panggil(req, Urlnya)
    Me.status = req.status
    If req.status = "200" Then
          Me.HasilJson = req.ResponseText
          Set p = JSON.parse(req.ResponseText)
          Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
          Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
          If Me.ServerCode = "200" Then
            jml = p("response")("jaminan").Count
            ReDim Hasilarry(jml, 6) As String
            For i = 0 To jml
                potsuplesinoSep.Add i, p("response")("jaminan")(i)("noRegister")
                potsuplesinoSepAwal.Add i, p("response")("jaminan")(i)("noSep")
                potsuplesinoSuratJaminan.Add i, p("response")("jaminan")(i)("noSuratJaminan")
                potsuplesitglKejadian.Add i, p("response")("jaminan")(i)("tglKejadian")
                potsuplesitglSep.Add i, p("response")("jaminan")(i)("tglSep")
            Next i
          Else
               Call MsgBox(Result)
          End If
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function Pengajuan(nokartu, tanggal, jnspelayanan, Keterangan, user As String) As Boolean
    Dim jml As Integer
    Dim JsonRequest As String
    Pengajuan = False
    
    Me.Url = "Sep/pengajuanSEP"
    JsonRequest = "{""request"": {""t_sep"": {""noKartu"": """ & nokartu & """,""tglSep"": """ & tanggal & """,""jnsPelayanan"": """ & jnspelayanan & """,""keterangan"": """ & Keterangan & """,""user"": """ & user & """}}}"
    Me.RequestJson = JsonRequest
    
    Call panggil(req, Me.Url, "POST", JsonRequest)
    Me.status = req.status
    If req.status = "200" Then
          Me.HasilJson = req.ResponseText
          Set p = JSON.parse(req.ResponseText)
          Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
          Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
          If Me.ServerCode = "200" Then
              Set Hasil = New Dictionary
              Hasil.Add "noKartu", p.Item("response")
              Pengajuan = True
          Else
               Call MsgBox(Result)
          End If
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function Approval(nokartu, tanggal, jnspelayanan, Keterangan, user As String) As Boolean
    Dim jml As Integer
    Dim JsonRequest As String
    Approval = False
    
    Me.Url = "Sep/aprovalSEP"
    JsonRequest = "{""request"": {""t_sep"": {""noKartu"": """ & nokartu & """,""tglSep"": """ & tanggal & """,""jnsPelayanan"": """ & jnspelayanan & """,""keterangan"": """ & Keterangan & """,""user"": """ & user & """}}}"
    Me.RequestJson = JsonRequest
    
    Call panggil(req, Me.Url, "POST", JsonRequest)
    Me.status = req.status
    If req.status = "200" Then
          Me.HasilJson = req.ResponseText
          Set p = JSON.parse(req.ResponseText)
          Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
          Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
          If Me.ServerCode = "200" Then
              Set Hasil = New Dictionary
              Hasil.Add "noKartu", p.Item("response")
              Approval = True
          Else
               Call MsgBox(Result)
          End If
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function UpdateTglPulang(NoSEP, TglPulang As String)
    Dim setting As csetting
    Dim callws As callws
    Dim JsonRequest As String
    Dim PPKPelayanan As String
    
    Set setting = New csetting
    setting.GetData
    PPKPelayanan = setting.NoPPK
    Me.Url = "Sep/updtglplg"
    JsonRequest = "{""request"":{""t_sep"":{""noSep"":""" & NoSEP & """,""tglPulang"":""" & TglPulang & """,""ppkPelayanan"":""" & PPKPelayanan & """}}}"
    Me.RequestJson = JsonRequest
    Call panggil(req, Me.Url, "PUT", JsonRequest)
    If req.status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
    End If
    
End Function

'################################ END SEP ##############################################################
'#######################################################################################################

'################################ RUJUKAN ##############################################################
'#######################################################################################################
Public Function BuatRujukan(NoSEP, tglrujukan, PPKdirujuk, jnspelayanan, catatan, diagRujukan, tiperujukan, PoliRujukan, user As String)
Dim JsonRequest As String

JsonRequest = "{""request"": {""t_rujukan"": {""noSep"": """ & NoSEP & """ ,""tglRujukan"": """ & tglrujukan & """,""ppkDirujuk"": """ & PPKdirujuk & """,""jnsPelayanan"": """ & jnspelayanan & """,""catatan"": """ & catatan & """,""diagRujukan"": """ & diagRujukan & """,""tipeRujukan"": """ & tiperujukan & """,""poliRujukan"": """ & PoliRujukan & """,""user"": """ & user & """}}}"
Me.RequestJson = JsonRequest

  ' Exit Function
   Call panggil(req, "/Rujukan/insert", "POST", JsonRequest)
   Me.status = req.status
   If req.status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & ServerCode & vbNewLine & " Keterangan " & ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
        Else
           Set Hasil = New Dictionary
           With Hasil
                .Add "kodeAsalRujukan", p.Item("response").Item("rujukan").Item("AsalRujukan").Item("kode")
                .Add "namaAsalRujukan", p.Item("response").Item("rujukan").Item("AsalRujukan").Item("nama")
                .Add "kodeDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("kode")
                .Add "namaDiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("nama")
                .Add "noRujukan", p.Item("response").Item("rujukan").Item("noRujukan")
                .Add "asuransi", p.Item("response").Item("rujukan").Item("peserta").Item("asuransi")
                .Add "hakKelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas")
                .Add "jnsPeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jnsPeserta")
                .Add "kelamin", p.Item("response").Item("rujukan").Item("peserta").Item("kelamin")
                .Add "nama", p.Item("response").Item("rujukan").Item("peserta").Item("nama")
                .Add "noKartu", p.Item("response").Item("rujukan").Item("peserta").Item("noKartu")
                .Add "nomr", p.Item("response").Item("rujukan").Item("peserta").Item("noMr")
                .Add "tglLahir", p.Item("response").Item("rujukan").Item("peserta").Item("tglLahir")
                .Add "kodePoli", p.Item("response").Item("rujukan").Item("poliTujuan").Item("kode")
                .Add "namaPoli", p.Item("response").Item("rujukan").Item("poliTujuan").Item("nama")
                .Add "tglRujukan", p.Item("response").Item("rujukan").Item("tglRujukan")
                .Add "kodeTujuanRujukan", p.Item("response").Item("rujukan").Item("tujuanRujukan").Item("kode")
                .Add "namaTujukanRujukan", p.Item("response").Item("rujukan").Item("tujuanRujukan").Item("nama")
           End With
       End If
    Else
        Call MsgBox("Error Ndaaa")
   End If
End Function
Public Function CariRujukankeRSWN(tglrujukan As String, isRS As Boolean)
    If isRS = False Then
        Call panggil(req, "Rujukan/List/TglRujukan/" & tglrujukan)
    Else
        Call panggil(req, "Rujukan/RS/List/TglRujukan/" & tglrujukan)
    End If
    Me.status = req.status
    If req.status = "200" Then
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       If Me.ServerMessage = "Gagal Proses Cari Data Rujukan" Then
            Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
            Me.ServerCode = "201"
       End If
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
            Me.HasilJson = req.ResponseText
            i = 1
            Set RTglRujukan = New Dictionary
            Set RNoPeserta = New Dictionary
            Set RNoRujukan = New Dictionary
            Set RNoCM = New Dictionary
            Set RNama = New Dictionary
            Set RKdPoli = New Dictionary
            Set RNmPoli = New Dictionary
            Set RKodeDiagnosa = New Dictionary
            Set RNamaDiagnosa = New Dictionary
            Set RNoPPKPerujuk = New Dictionary
            Set RNmPPkPerujuk = New Dictionary
            For i = 1 To p("response")("rujukan").Count
               Me.RTglRujukan.Add i, p("response")("rujukan")(i)("tglKunjungan")
               Me.RNoPeserta.Add i, p("response")("rujukan")(i)("peserta")("noKartu")
               Me.RNoRujukan.Add i, p("response")("rujukan")(i)("noKunjungan")
               Me.RNoCM.Add i, p("response")("rujukan")(i)("peserta")("mr")("noMR")
               Me.RNama.Add i, p("response")("rujukan")(i)("peserta")("nama")
               Me.RKdPoli.Add i, p("response")("rujukan")(i)("poliRujukan")("kode")
               Me.RNmPoli.Add i, p("response")("rujukan")(i)("poliRujukan")("nama")
               Me.RKodeDiagnosa.Add i, p("response")("rujukan")(i)("diagnosa")("kode")
               Me.RNamaDiagnosa.Add i, p("response")("rujukan")(i)("diagnosa")("nama")
               Me.RNoPPKPerujuk.Add i, p("response")("rujukan")(i)("provPerujuk")("kode")
               Me.RNmPPkPerujuk.Add i, p("response")("rujukan")(i)("provPerujuk")("nama")
            Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function

'################################ END RUJUKAN ##########################################################
'#######################################################################################################

'################################ MONITORING ##########################################################
'#######################################################################################################
Public Function historipelayananpasien(nokartu, TglAwal, TglAkhir As String)
    Call panggil(req, "monitoring/HistoriPelayanan/NoKartu/" & nokartu & "/tglAwal/" & TglAwal & "/tglAkhir/" & TglAkhir, "GET")
     If req.status = "200" Then
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       If Me.ServerMessage = "Gagal Proses Cari Data Rujukan" Then
            Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
            Me.ServerCode = "201"
       End If
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       Me.HasilJson = req.ResponseText
       i = 0
       Set historysepDiagnosa = New Dictionary
       Set historysepJnspelayanan = New Dictionary
       Set historysepKelasrawat = New Dictionary
       Set historysepNamapeserta = New Dictionary
       Set historysepNokartu = New Dictionary
       Set historysepNosep = New Dictionary
       Set historysepNorujukan = New Dictionary
       Set historysepPoli = New Dictionary
       Set historysepPpkpelayanan = New Dictionary
       Set historysepTglpulangsep = New Dictionary
       Set historysepTglsep = New Dictionary
       For i = 1 To p("response")("histori").Count
           Me.historysepDiagnosa.Add i, p("response")("histori")(i)("diagnosa")
           Me.historysepJnspelayanan.Add i, p("response")("histori")(i)("jnsPelayanan")
           Me.historysepKelasrawat.Add i, p("response")("histori")(i)("kelasRawat")
           Me.historysepNamapeserta.Add i, p("response")("histori")(i)("namaPeserta")
           Me.historysepNokartu.Add i, p("response")("histori")(i)("noKartu")
           Me.historysepNosep.Add i, p("response")("histori")(i)("noSep")
           Me.historysepNorujukan.Add i, p("response")("histori")(i)("noRujukan")
           Me.historysepPoli.Add i, p("response")("histori")(i)("poli")
           Me.historysepPpkpelayanan.Add i, p("response")("histori")(i)("ppkPelayanan")
           Me.historysepTglpulangsep.Add i, p("response")("histori")(i)("tglPlgSep")
           Me.historysepTglsep.Add i, p("response")("histori")(i)("tglSep")
       Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function datakunjungan(tanggal, jnspelayanan As String)
    Call panggil(req, "Monitoring/Kunjungan/Tanggal/" & tanggal & "/JnsPelayanan/" & jnspelayanan, "GET")
     If req.status = "200" Then
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       If Me.ServerMessage = "Gagal Proses Cari Data Rujukan" Then
            Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
            Me.ServerCode = "201"
       End If
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       Me.HasilJson = req.ResponseText
       i = 0
       Set historysepDiagnosa = New Dictionary
       Set historysepJnspelayanan = New Dictionary
       Set historysepKelasrawat = New Dictionary
       Set historysepNamapeserta = New Dictionary
       Set historysepNokartu = New Dictionary
       Set historysepNosep = New Dictionary
       Set historysepNorujukan = New Dictionary
       Set historysepPoli = New Dictionary
       Set historysepPpkpelayanan = New Dictionary
       Set historysepTglpulangsep = New Dictionary
       Set historysepTglsep = New Dictionary
       For i = 1 To p("response")("sep").Count
           Me.historysepDiagnosa.Add i, p("response")("sep")(i)("diagnosa")
           Me.historysepJnspelayanan.Add i, p("response")("sep")(i)("jnsPelayanan")
           Me.historysepKelasrawat.Add i, p("response")("sep")(i)("kelasRawat")
           Me.historysepNamapeserta.Add i, p("response")("sep")(i)("nama")
           Me.historysepNokartu.Add i, p("response")("sep")(i)("noKartu")
           Me.historysepNosep.Add i, p("response")("sep")(i)("noSep")
           Me.historysepNorujukan.Add i, p("response")("sep")(i)("noRujukan")
           Me.historysepPoli.Add i, p("response")("sep")(i)("poli")
           Me.historysepTglpulangsep.Add i, p("response")("sep")(i)("tglPlgSep")
           Me.historysepTglsep.Add i, p("response")("sep")(i)("tglSep")
       Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function dataklaim(tanggal, jnspelayanan, status As String)
    Call panggil(req, "Monitoring/Klaim/Tanggal/" & tanggal & "/JnsPelayanan/" & jnspelayanan & "/Status/" & status, "GET")
     If req.status = "200" Then
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       If Me.ServerMessage = "Gagal Proses Cari Data Rujukan" Then
            Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
            Me.ServerCode = "201"
       End If
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       Me.HasilJson = req.ResponseText
       i = 0
        Set dataklaimKodeinacbg = New Dictionary
        Set dataklaimNamainacbg = New Dictionary
        Set dataklaimBiayabypengajuan = New Dictionary
        Set dataklaimBiayabySetujui = New Dictionary
        Set dataklaimBiayabyTarifGruper = New Dictionary
        Set dataklaimBiayabyTarifRS = New Dictionary
        Set dataklaimBiayabyTopup = New Dictionary
        Set dataklaimKelasrawat = New Dictionary
        Set dataklaimNoFPK = New Dictionary
        Set dataklaimNosep = New Dictionary
        Set dataklaimPesertanama = New Dictionary
        Set dataklaimPesertanokartu = New Dictionary
        Set dataklaimPesertanomr = New Dictionary
        Set dataklaimPoli = New Dictionary
        Set dataklaimstatus = New Dictionary
        Set dataklaimTglpulang = New Dictionary
        Set dataklaimTglsep = New Dictionary
        For i = 1 To p("response")("klaim").Count
           Me.dataklaimKodeinacbg.Add i, p("response")("klaim")(i)("Inacbg")("kode")
           Me.dataklaimNamainacbg.Add i, p("response")("klaim")(i)("Inacbg")("nama")
           Me.dataklaimBiayabypengajuan.Add i, p("response")("klaim")(i)("biaya")("byPengajuan")
           Me.dataklaimBiayabySetujui.Add i, p("response")("klaim")(i)("biaya")("bySetujui")
           Me.dataklaimBiayabyTarifGruper.Add i, p("response")("klaim")(i)("biaya")("byTarifGruper")
           Me.dataklaimBiayabyTarifRS.Add i, p("response")("klaim")(i)("biaya")("byTarifRS")
           Me.dataklaimBiayabyTopup.Add i, p("response")("klaim")(i)("biaya")("byTopup")
           Me.dataklaimKelasrawat.Add i, p("response")("klaim")(i)("kelasRawat")
           Me.dataklaimNoFPK.Add i, p("response")("klaim")(i)("noFPK")
           Me.dataklaimNosep.Add i, p("response")("klaim")(i)("noSEP")
           Me.dataklaimPesertanama.Add i, p("response")("klaim")(i)("peserta")("nama")
           Me.dataklaimPesertanokartu.Add i, p("response")("klaim")(i)("peserta")("noKartu")
           Me.dataklaimPesertanomr.Add i, p("response")("klaim")(i)("peserta")("noMR")
           Me.dataklaimPoli.Add i, p("response")("klaim")(i)("poli")
           Me.dataklaimstatus.Add i, p("response")("klaim")(i)("status")
           Me.dataklaimTglpulang.Add i, p("response")("klaim")(i)("tglPulang")
           Me.dataklaimTglsep.Add i, p("response")("klaim")(i)("tglSep")
       Next i
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function


'################################ UTILITY ##################################################################
'#######################################################################################################
Private Function panggil(req As callws, Url As String, Optional Method As String, Optional body As String, Optional BypassUrlSetting As Boolean)
   Set req = New callws
   req.EndPoint = Url
   If Method = "" Then
        req.Method = "GET"
        Call req.Send
   ElseIf Method = "POST" Or Method = "DELETE" Or Method = "PUT" Then
        req.Method = Method
        Call req.Send(body, False)
   Else
        req.Method = Method
        Call req.Send
   End If
End Function

Public Function PesanError(strPesan, strKode As String)
    Me.ServerMessage = strPesan
    Me.ServerCode = strKode
    Call MsgBox(strKode & vbNewLine & strPesan, vbOKOnly, "PERHATIAN")
End Function
'################################ END ##################################################################
'#######################################################################################################


