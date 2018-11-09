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
Public Hasil As Dictionary
Public ServerMessage As String
Public ServerCode As String
Private req As CallWS
Private Result As String
Private p As Object
Private i As Integer
Public Function CariPesertaByNoKartu(NoKartu As String)
   Call panggil(req, "Peserta/nokartu/" & NoKartu & "/tglSEP/" & Format(Now, "yyyy-MM-dd"))
   If req.Status = "200" Then
        Result = req.ResponseText
        Set p = JSON.parse(Result)
        Me.ServerMessage = IIf(IsNull(p.Item("metaData").Item("message")) = True, "", p.Item("metaData").Item("message"))
        Me.ServerCode = IIf(IsNull(p.Item("metaData").Item("code")) = True, "", p.Item("metaData").Item("code"))
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
        Call PesanError("SERVER BPJS ERROR", "401")
   End If
End Function
Private Function panggil(req As CallWS, url As String, Optional Method As String, Optional body As String)
   Set req = New CallWS
   req.EndPoint = url
   If Method = "" Then
        req.Method = "GET"
   Else
        req.Method = Method
   End If
   If Method = "POST" Then
        req.Send body
    Else
        req.Send
    End If
End Function
Public Function PesanError(strPesan, strKode As String)
    Me.ServerMessage = strPesan
    Me.ServerCode = strKode
    Call MsgBox(strKode & vbNewLine & strPesan, vbOKOnly, "PERHATIAN")
End Function

