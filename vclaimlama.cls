VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vclaimLama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' Class:       vclaim
' FILENAME:     vclaim.cls
' AUTHOR:       Martin Eko Setiawan
' CREATED:      19/02/2018
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
Public ConsID As String
Public SecretKey As String
Public HasilKey As String
Public AlamatWebService As String
Public ServerCode As String
Public ServerMessage As String


Public DataPesertaByNIK As Dictionary
Public DataPesertaByNoKartu As Dictionary
Public resKdCariPoli As Dictionary
Public resNamaCariPoli As Dictionary
Public resKdCariFaskes As Dictionary
Public resNamaCariFaskes As Dictionary
Public resKdDiagnosa As Dictionary
Public resNamaDiagnosa As Dictionary
Public resKdDokter As Dictionary
Public resNamaDokter As Dictionary
Public resKdSpesialis As Dictionary
Public resNamaSpesialis As Dictionary
Public resKdPropinsi As Dictionary
Public resNamaPropinsi As Dictionary
Public resKdKota As Dictionary
Public resNamaKota As Dictionary
Public resKdKecamatan As Dictionary
Public resNamaKecamatan As Dictionary
Public jsonKirim As String
Private req As CallWS
Private Result As String
Private p As Object
Private i As Integer
Public Function CariDataByNoKartu(NoKartu As String)
   Call panggil(req, Me.AlamatWebService & "Peserta/nokartu/" & NoKartu & "/tglSEP/" & Format(Now, "yyyy-MM-dd"))
   If req.Status = "200" Then
       Result = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Me.ServerMessage = ServerMessage
       Me.ServerCode = ServerCode
       Set DataPesertaByNoKartu = New Dictionary
       With DataPesertaByNoKartu
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
        Call MsgBox("Error Ndaaa")
   End If
End Function
Public Function CariDataByNIK(NoKartu As String)
   Call panggil(req, Me.AlamatWebService & "Peserta/nik/" & NoKartu & "/tglSEP/" & Format(Now, "yyyy-MM-dd"))
   If req.Status = "200" Then
       Result = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Me.ServerMessage = ServerMessage
       Me.ServerCode = ServerCode
       If Not ServerCode = "200" Then
            Call MsgBox("Pesan Server " & ServerCode & vbNewLine & " Keterangan " & ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       End If
       
       Set DataPesertaByNIK = New Dictionary
       With DataPesertaByNIK
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
        Call MsgBox("Error Ndaaa")
   End If
   
End Function
Public Function CariPoli(NamaPoli As String)
  Call panggil(req, Me.AlamatWebService & "referensi/poli/" & NamaPoli)
  If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(req.ResponseText)
        ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
        Me.ServerMessage = ServerMessage
        Me.ServerCode = ServerCode
        If Not ServerCode = "200" Then
            Call MsgBox("Pesan Server " & ServerCode & vbNewLine & " Keterangan " & ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
        End If
        i = 1
        Set Me.resKdCariPoli = New Dictionary
        Set Me.resNamaCariPoli = New Dictionary
        If ServerCode = "200" Then
            For i = 1 To p("response")("faskes").Count
                resKdCariPoli.Add i, p("response")("poli")(i)("kode")
                resNamaCariPoli.Add i, p("response")("poli")(i)("nama")
            Next i
        End If
    Else
        Call MsgBox("Error Ndaaa")
   End If
End Function
Public Function CariFaskes(NamaFaskes, Faskes As String)
    Call panggil(req, Me.AlamatWebService & "referensi/faskes/" & NamaFaskes & "/" & Faskes)
    If req.Status = "200" Then
       Result = req.ResponseText
       Set p = JSON.parse(req.ResponseText)
       Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & Me.ServerCode & vbNewLine & " Keterangan " & Me.ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       Else
            i = 1
            Set Me.resKdCariFaskes = New Dictionary
            Set Me.resNamaCariFaskes = New Dictionary
            For i = 1 To p("response")("faskes").Count
                resKdCariFaskes.Add i, p("response")("faskes")(i)("kode")
                resNamaCariFaskes.Add i, p("response")("faskes")(i)("nama")
            Next i
    End If
End Function
Public Function CariDiagnosa(NamaDiagnosa As String)
    Call panggil(req, Me.AlamatWebService & "/referensi/diagnosa/" & NamaDiagnosa)
    If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(Result)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & Me.ServerCode & vbNewLine & " Keterangan " & Me.ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       Else
            i = 1
            Set Me.resKdDiagnosa = New Dictionary
            Set Me.resNamaDiagnosa = New Dictionary
            For i = 1 To p("response")("diagnosa").Count
                resKdDiagnosa.Add i, p("response")("diagnosa")(i)("kode")
                resNamaDiagnosa.Add i, p("response")("diagnosa")(i)("nama")
            Next i
        End If
    End If
End Function
Public Function refDokterDPJP(JenisPelayanan As String, TglPelayanan As String, Spesialis As String)
    Call panggil(req, Me.AlamatWebService & "/referensi/dokter/pelayanan/" & JenisPelayanan & "/tglPelayanan/" & TglPelayanan & "/Spesialis/" & Spesialis)
    If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(Result)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & Me.ServerCode & vbNewLine & " Keterangan " & Me.ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       Else
            i = 1
            Set Me.resKdDokter = New Dictionary
            Set Me.resNamaDokter = New Dictionary
            For i = 1 To p("response")("list").Count
                resKdDokter.Add i, p("response")("list")(i)("kode")
                resNamaDokter.Add i, p("response")("list")(i)("nama")
            Next i
        End If
    End If
End Function
Public Function refSpesialis()
    Call panggil(req, Me.AlamatWebService & "/referensi/spesialistik")
    If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(Result)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & Me.ServerCode & vbNewLine & " Keterangan " & Me.ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       Else
            i = 1
            Set Me.resKdSpesialis = New Dictionary
            Set Me.resNamaSpesialis = New Dictionary
            For i = 1 To p("response")("list").Count
                resKdSpesialis.Add i, p("response")("list")(i)("kode")
                resNamaSpesialis.Add i, p("response")("list")(i)("nama")
            Next i
        End If
    End If
End Function
Public Function refPropinsi()
    Call panggil(req, Me.AlamatWebService & "/referensi/propinsi")
    If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(Result)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & Me.ServerCode & vbNewLine & " Keterangan " & Me.ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       Else
            i = 1
            Set Me.resKdPropinsi = New Dictionary
            Set Me.resNamaPropinsi = New Dictionary
            For i = 1 To p("response")("list").Count
                resKdPropinsi.Add i, p("response")("list")(i)("kode")
                resNamaPropinsi.Add i, p("response")("list")(i)("nama")
            Next i
        End If
    End If
End Function
Public Function refkota(kdPropinsi As String)
    Call panggil(req, Me.AlamatWebService & "/referensi/kabupaten/propinsi/" & kdPropinsi)
    If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(Result)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & Me.ServerCode & vbNewLine & " Keterangan " & Me.ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       Else
            i = 1
            Set Me.resKdKota = New Dictionary
            Set Me.resNamaKota = New Dictionary
            For i = 1 To p("response")("list").Count
                resKdKota.Add i, p("response")("list")(i)("kode")
                resNamaKota.Add i, p("response")("list")(i)("nama")
            Next i
        End If
    End If
End Function
Public Function refkecamatan(kdKota As String)
    Call panggil(req, Me.AlamatWebService & "/referensi/kecamatan/kabupaten/" & kdKota)
    If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(Result)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       If Not Me.ServerCode = "200" Then
            Call MsgBox("Pesan Server " & Me.ServerCode & vbNewLine & " Keterangan " & Me.ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
       Else
            i = 1
            Set Me.resKdKecamatan = New Dictionary
            Set Me.resNamaKecamatan = New Dictionary
            For i = 1 To p("response")("list").Count
                resKdKecamatan.Add i, p("response")("list")(i)("kode")
                resNamaKecamatan.Add i, p("response")("list")(i)("nama")
            Next i
        End If
    End If
End Function

Public Function CreateSEP(NoKartu, tglSep, ppkPelayanan, jnsPelayanan, klsRawat, noMR, asalRujukan, tglRujukan, norujukan, ppkRujukan, catatan, diagAwal, politujuan, polieksekutif, cob, jaminanLakalantas, jaminanPenjamin, jaminanLokasiLaka, noTelp, user As String)
Dim jsonRequest As String
Dim req As WinHttp.WinHttpRequest
Dim TglNow As Date
Dim TglSekarang As String
Dim url As String
Dim Result As String
Dim ServerMessage As String
Dim ServerCode As String
Dim p As Object


jsonRequest = "{" & _
           """request"": {" & _
              """t_sep"": {" & _
                 """noKartu"": """ & NoKartu & """," & _
                 """tglSep"": """ & tglSep & """, " & _
                 """ppkPelayanan"": """ & ppkPelayanan & """, " & _
                 """jnsPelayanan"": """ & jnsPelayanan & """, " & _
                 """klsRawat"": """ & klsRawat & """, " & _
                 """noMR"": """ & noMR & """, " & _
                 """rujukan"": { " & _
                    """asalRujukan"": """ & asalRujukan & """, " & _
                    """tglRujukan"": """ & tglRujukan & """, " & _
                    """noRujukan"": """ & norujukan & """, " & _
                    """ppkRujukan"": """ & ppkRujukan & """}," & _
                 """catatan"": """ & catatan & """, " & _
                 """diagAwal"": """ & diagAwal & """," & _
                 """poli"": { " & _
                 """tujuan"": """ & politujuan & """," & _
                 """eksekutif"": """ & polieksekutif & """ },""cob"": { ""cob"": """ & cob & """},""jaminan"": { ""lakaLantas"": """ & jaminanLakalantas & """, ""penjamin"": """ & jaminanPenjamin & """," & _
                 """lokasiLaka"": """ & jaminanLokasiLaka & """ },""noTelp"": """ & noTelp & """,""user"": """ & user & """}}}"
   
   Call panggil(req, Me.AlamatWebService & "/SEP/insert", "POST", jsonRequest)
   
   If req.Status = "200" Then
       Result = req.ResponseText
       
       Set p = JSON.parse(req.ResponseText)
       ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
       ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
       Me.ServerMessage = ServerMessage
       Me.ServerCode = ServerCode
       If Not ServerCode = "200" Then
            Call MsgBox("Pesan Server " & ServerCode & vbNewLine & " Keterangan " & ServerMessage, vbOKOnly, "PERHATIAN ALIAS WARNING BRIDGING SEP BPJS")
            Exit Function
        Else
            Call MsgBox(Result)
       End If
    Else
        Call MsgBox("Error Ndaaa")
   End If
End Function
Private Function panggil(req As CallWS, url As String, Optional Method As String, Optional body As String)
   Set req = New CallWS
   req.EndPoint = url
   If Method = "" Then
        req.Method = "GET"
        Call req.Send
   Else
        req.Method = Metho
        Call req.Send
   End If
'   If Method = "POST" Then
'        req.Send (body)
'    Else
'        req.Send
'    End If
End Function
