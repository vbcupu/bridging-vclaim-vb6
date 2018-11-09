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
' COPYRIGHT:    @2018 martin eko setiawan for www.martinsetiawan.com
'
' DESCRIPTION:
' Class Setting data untuk aplikasi bridging VClaim dengan visual basic 6
' message. Baca keseluruhan class ini untuk lebih detailnya
' Saya berbagi class ini bukan karena karena saya ingin berbagi kepada teman-teman programmer rumah sakit yang kesulitan dalam bridging vb6 dengan vclaim
' Pada dasarnya saya masih belajar dan terus belajar, jadi maaf bila source ini masih banyak kekurangan, acak-acakan dan banyak bug-nya

' Ini adalah class yang gratis tis, tanpa biaya, Dengan Catatan:
' 1. Kode ini boleh didistribusikan ulang ke source code dan aplikasi lain.
' 2. Class ini hanya rintisan, minimal dengan modal class ini, kamu bisa mengembangkan modul lainnya karena prinsip kerja baik request maupun pengiriman parameter hampir sama.
' 3. Tolong author nama saya jangan dihapus ya, kalau mau ditambah nggak apa-apa, kasih aja di bawah atau diatas nama saya.
' 4. Karena Source Code ini saya berikan dengan gratis, maka saya berharap source code ini tidak dipakai oleh perusahaan developer berbayar. Karena perusahaan developer harusnya bisa membuat/meresearch API sendiri.
' 5. Saya tidak menjamin aplikasi ini bebas bug, tapi aplikasi ini saya pakai di kerjaan juga, jadi InsyaAllah bebas bug.
' 6. Sekali lagi saya berharap kamu menghormati tentang Copyright, jangan mengklaim yang tidak murni buatanmu sebagai buatanmu.
'     Diaplikasi ini saya juga memakai class-class dan module-module buatan orang lain, dan saya tidak merubah sedikitpun dari source yang saya dapat.

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

Public url As String
Public EndPoint As String
Public Method As String
Public Status As String
Public ResponseText As String
Public ConsID As String
Public SecretKey As String

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
Private req As CallWS
Private Result As String
Private p As Object
Private i As Integer
'############################ CARI PASIEN ############################################################################################
'############################ CARI PASIEN ############################################################################################
Public Function CariPesertaByNoKartu(NoKartu As String)
   Call panggil(req, "Peserta/nokartu/" & NoKartu & "/tglSEP/" & Format(Now, "yyyy-MM-dd"))
   If req.Status = "200" Then
        Me.HasilJson = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
        Set Hasil = New Dictionary
        With Hasil
            .Add "COBAsuransi", p.Item("response").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "COBNoAsuransi", p.Item("response").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "COBPesertaTglTAT", p.Item("response").Item("peserta").Item("cob").Item("tglTAT")
            .Add "COBPesertaTglTMT", p.Item("response").Item("peserta").Item("cob").Item("tglTMT")
            .Add "hakkelas", p.Item("response").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "hakkelasKode", p.Item("response").Item("peserta").Item("hakKelas").Item("kode")
            .Add "informasiDinsos", p.Item("response").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "jenisPeserta", p.Item("response").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "jenisPesertaKode", p.Item("response").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "noMR", p.Item("response").Item("peserta").Item("mr").Item("noMR")
            .Add "noMRnoTelepon", p.Item("response").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("peserta").Item("pisa")
            .Add "provUmum", p.Item("response").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "provuUmum", p.Item("response").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("peserta").Item("sex")
            .Add "statusPeserta", p.Item("response").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "statusPesertaKode", p.Item("response").Item("peserta").Item("statusPeserta").Item("kode")
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
Public Function CariPesertaByNIK(nik As String)
   Call panggil(req, "Peserta/nik/" & nik & "/tglSEP/" & Format(Now, "yyyy-MM-dd"))
   If req.Status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "COBAsuransi", p.Item("response").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "COBNoAsuransi", p.Item("response").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "COBPesertaTglTAT", p.Item("response").Item("peserta").Item("cob").Item("tglTAT")
            .Add "COBPesertaTglTMT", p.Item("response").Item("peserta").Item("cob").Item("tglTMT")
            .Add "hakkelas", p.Item("response").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "hakkelasKode", p.Item("response").Item("peserta").Item("hakKelas").Item("kode")
            .Add "informasiDinsos", p.Item("response").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasiNoSKTM", p.Item("response").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiProlanisPRB", p.Item("response").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "jenisPeserta", p.Item("response").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "jenisPesertaKode", p.Item("response").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "noMR", p.Item("response").Item("peserta").Item("mr").Item("noMR")
            .Add "noMRnoTelepon", p.Item("response").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("peserta").Item("nik")
            .Add "noKartu", p.Item("response").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("peserta").Item("pisa")
            .Add "provUmum", p.Item("response").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "provuUmum", p.Item("response").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "jk", p.Item("response").Item("peserta").Item("sex")
            .Add "statusPeserta", p.Item("response").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "statusPesertaKode", p.Item("response").Item("peserta").Item("statusPeserta").Item("kode")
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
Public Function CariByRujukanPcare(norujukan As String)
   Call panggil(req, "Rujukan/" & norujukan)
   If req.Status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "kodediagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("kode")
            .Add "namadiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("nama")
            .Add "keluhan", p.Item("response").Item("rujukan").Item("keluhan")
            .Add "nokunjungan", p.Item("response").Item("rujukan").Item("noKunjungan")
            .Add "kodepelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("kode")
            .Add "namapelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("nama")
            .Add "cobnmasuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobnoasuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobtgltat", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobtgltmt", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTMT")
            .Add "kodehakkelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("kode")
            .Add "keteranganhakkelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "informasidinsos", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasinoSKTM", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiprolanisPRB", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "kodejenispeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("kode")
            .Add "keteranganjenispeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenisPeserta").Item("keterangan")
            .Add "nomr", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrtelepon", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("rujukan").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("rujukan").Item("peserta").Item("nik")
            .Add "nokartu", p.Item("response").Item("rujukan").Item("peserta").Item("noKartu")
            .Add "pisa", p.Item("response").Item("rujukan").Item("peserta").Item("pisa")
            .Add "provumumkdprovider", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "provumumnmprovider", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "JK", p.Item("response").Item("rujukan").Item("peserta").Item("sex")
            .Add "statuspeserta", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("keterangan")
            .Add "statuspesertakode", p.Item("response").Item("rujukan").Item("peserta").Item("statusPeserta").Item("kode")
            .Add "tglcetakkartu", p.Item("response").Item("rujukan").Item("peserta").Item("tglCetakKartu")
            .Add "tgllahir", p.Item("response").Item("rujukan").Item("peserta").Item("tglLahir")
            .Add "tgltat", p.Item("response").Item("rujukan").Item("peserta").Item("tglTAT")
            .Add "tgltmt", p.Item("response").Item("rujukan").Item("peserta").Item("tglTMT")
            .Add "umursaatpelayanan", p.Item("response").Item("rujukan").Item("peserta").Item("umurSaatPelayanan")
            .Add "umursekarang", p.Item("response").Item("rujukan").Item("peserta").Item("umurSekarang")
            .Add "kodepolirujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("kode")
            .Add "namapolirujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("nama")
            .Add "kodeprovperujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("kode")
            .Add "namaprovperujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("nama")
            .Add "tgldirujuk", p.Item("response").Item("rujukan").Item("tglKunjungan")
       End With
    Else
        Call PesanError("Server BPJS Tidak Dapat Dihubungi", "401")
    End If
End Function
Public Function CariByRujukanRS(norujukan As String)
   Call panggil(req, "/Rujukan/RS/" & norujukan)
   If req.Status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Set Hasil = New Dictionary
       If Me.ServerCode <> "200" Then Call PesanError(Me.ServerMessage, Me.ServerCode): Exit Function
       With Hasil
            .Add "kodediagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("kode")
            .Add "namadiagnosa", p.Item("response").Item("rujukan").Item("diagnosa").Item("nama")
            .Add "keluhan", p.Item("response").Item("rujukan").Item("keluhan")
            .Add "nokunjungan", p.Item("response").Item("rujukan").Item("noKunjungan")
            .Add "kodepelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("kode")
            .Add "namapelayanan", p.Item("response").Item("rujukan").Item("pelayanan").Item("nama")
            .Add "cobnmasuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("nmAsuransi")
            .Add "cobnoasuransi", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("noAsuransi")
            .Add "cobtgltat", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTAT")
            .Add "cobtgltmt", p.Item("response").Item("rujukan").Item("peserta").Item("cob").Item("tglTMT")
            .Add "kodehakkelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("kode")
            .Add "keteranganhakkelas", p.Item("response").Item("rujukan").Item("peserta").Item("hakKelas").Item("keterangan")
            .Add "informasidinsos", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("dinsos")
            .Add "informasinoSKTM", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("noSKTM")
            .Add "informasiprolanisPRB", p.Item("response").Item("rujukan").Item("peserta").Item("informasi").Item("prolanisPRB")
            .Add "kodejenispeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenispeserta").Item("kode")
            .Add "keteranganjenispeserta", p.Item("response").Item("rujukan").Item("peserta").Item("jenispeserta").Item("keterangan")
            .Add "nomr", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noMR")
            .Add "nomrtelepon", p.Item("response").Item("rujukan").Item("peserta").Item("mr").Item("noTelepon")
            .Add "nama", p.Item("response").Item("rujukan").Item("peserta").Item("nama")
            .Add "nik", p.Item("response").Item("rujukan").Item("peserta").Item("nik")
            .Add "nokartu", p.Item("response").Item("rujukan").Item("peserta").Item("nokartu")
            .Add "pisa", p.Item("response").Item("rujukan").Item("peserta").Item("pisa")
            .Add "provumumkdprovider", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("kdProvider")
            .Add "provumumnmprovider", p.Item("response").Item("rujukan").Item("peserta").Item("provUmum").Item("nmProvider")
            .Add "JK", p.Item("response").Item("rujukan").Item("peserta").Item("sex")
            .Add "statuspeserta", p.Item("response").Item("rujukan").Item("peserta").Item("StatusPeserta").Item("keterangan")
            .Add "statuspesertakode", p.Item("response").Item("rujukan").Item("peserta").Item("StatusPeserta").Item("kode")
            .Add "tglcetakkartu", p.Item("response").Item("rujukan").Item("peserta").Item("tglCetakKartu")
            .Add "tgllahir", p.Item("response").Item("rujukan").Item("peserta").Item("tglLahir")
            .Add "tgltat", p.Item("response").Item("rujukan").Item("peserta").Item("tglTAT")
            .Add "tgltmt", p.Item("response").Item("rujukan").Item("peserta").Item("tglTMT")
            .Add "umursaatpelayanan", p.Item("response").Item("rujukan").Item("peserta").Item("umurSaatPelayanan")
            .Add "umursekarang", p.Item("response").Item("rujukan").Item("peserta").Item("umurSekarang")
            .Add "kodepolirujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("kode")
            .Add "namapolirujukan", p.Item("response").Item("rujukan").Item("poliRujukan").Item("nama")
            .Add "kodeprovperujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("kode")
            .Add "namaprovperujuk", p.Item("response").Item("rujukan").Item("provPerujuk").Item("nama")
            .Add "tgldirujuk", p.Item("response").Item("rujukan").Item("tglKunjungan")
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
  If req.Status = "200" Then
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
    If req.Status = "200" Then
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
    If req.Status = "200" Then
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
    If req.Status = "200" Then
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
    If req.Status = "200" Then
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
    If req.Status = "200" Then
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
    If req.Status = "200" Then
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
    If req.Status = "200" Then
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
Public Function DetailSEP(nosep As String)
    Call panggil(req, "SEP/" & nosep)
    If req.Status = "200" Then
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
Public Function BuatSEP(NoKartu, tglSep, ppkPelayanan, jnsPelayanan, klsRawat, noMR, asalRujukan, tglRujukan, norujukan, ppkRujukan, catatan, diagAwal, politujuan, polieksekutif, cob, katarak, jamlakaLantas, jampenjamin, jamtglKejadian, jamketerangan, jamsuplesi, jamnoSepSuplesi, jamlokasilakakdPropinsi, jamlokasilakakdKabupaten, jamlokasilakakdKecamatan, skdpnoSurat, skdpkodeDPJP, noTelp, user As String)
Dim jsonRequest As String
If jamlakaLantas = "0" Then
    jamsuplesi = ""
End If

jsonRequest = "{""request"":{""t_sep"":{""noKartu"":""" & NoKartu & """,""tglSep"":""" & tglSep & """ ,""ppkPelayanan"":""" & ppkPelayanan & """,""jnsPelayanan"":""" & jnsPelayanan & """ ,""klsRawat"":""" & klsRawat & """,""noMR"":""" & noMR & """ ,""rujukan"":{""asalRujukan"":""" & asalRujukan & """,""tglRujukan"":""" & tglRujukan & """ ,""noRujukan"":""" & norujukan & """,""ppkRujukan"":""" & ppkRujukan & """},""catatan"":""" & catatan & """,""diagAwal"":""" & diagAwal & """,""poli"":{""tujuan"":""" & politujuan & """,""eksekutif"":""" & polieksekutif & """},""cob"":{""cob"":""" & cob & """},""katarak"":{""katarak"":""" & katarak & """},""jaminan"":{""lakaLantas"":""" & jamlakaLantas & """,""penjamin"":{""penjamin"":" & _
           """" & jampenjamin & """,""tglKejadian"":""" & jamtglKejadian & """,""keterangan"":""" & jamketerangan & """ ,""suplesi"":{""suplesi"":""" & IIf(jamsuplesi = "", 0, jamsuplesi) & """ ,""noSepSuplesi"":""" & jamnoSepSuplesi & """ ,""lokasiLaka"":{""kdPropinsi"":""" & jamlokasilakakdPropinsi & """ ,""kdKabupaten"":""" & jamlokasilakakdKabupaten & """,""kdKecamatan"":""" & jamlokasilakakdKecamatan & """}}}},""skdp"":{""noSurat"":""" & skdpnoSurat & """ ,""kodeDPJP"":""" & skdpkodeDPJP & """},""noTelp"":""" & noTelp & """,""user"":""" & user & """}}}"
  ' Exit Function
   Call panggil(req, "SEP/insert", "POST", jsonRequest)
   
   If req.Status = "200" Then
       Me.HasilJson = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & ServerCode & vbNewLine & " Keterangan " & ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
        Else
            Call MsgBox(req.ResponseText)
       End If
    Else
        Call MsgBox("Error Ndaaa")
   End If
End Function
Public Function HapusSEP(nosep As String)
    Dim jsonRequest As String
    
    jsonRequest = "{" & _
       """request"": {" & _
          """t_sep"": {" & _
             """noSep"": """ & nosep & """," & _
             """user"": ""wsRSUDKota""" & _
          "}" & _
       "}" & _
    "}"
    Call panggil(req, "SEP/Delete", "DELETE", jsonRequest)
    If req.Status = "200" Then
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
        Call MsgBox(req.Status & " " & req.ResponseText, vbOKOnly, "WARNING")
    End If
End Function
Public Function CariPotensiSuplesi(NoKartu, Tanggal As String)
    Dim jml As Integer
    Dim Urlnya As String
    
    Urlnya = "/sep/JasaRaharja/Suplesi/" & NoKartu & "/tglPelayanan/" & Tanggal
    Call panggil(req, Urlnya)
    If req.Status = "200" Then
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


'################################ UTILITY ##################################################################
'#######################################################################################################
Private Function panggil(req As CallWS, url As String, Optional Method As String, Optional body As String)
   Set req = New CallWS
   req.EndPoint = url
   If Method = "" Then
        req.Method = "GET"
        Call req.Send
   ElseIf Method = "POST" Or Method = "DELETE" Or Me.Method = "PUT" Then
        req.Method = Method
        Call req.Send(body)
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
