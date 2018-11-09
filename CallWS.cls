VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' Class:        CallWS.cls
' FILENAME:     CallWS.cls
' AUTHOR:       Martin Eko Setiawan
' CREATED:      19/02/2018
' COPYRIGHT:    @ 2018 martin eko setiawan for www.martinsetiawan.com
'
' DESCRIPTION:
' Class CallWS adalah file helper untuk meringkaskan request ke API BPJS
' message. Baca keseluruhan class ini untuk lebih detailnya
'
' Ini adalah class yang gratis tis, tanpa biaya, Dengan Catatan:

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
Public ServerCode As String
Public ServerMessage As String

Public Function Send(Optional body As String, Optional bypass As Boolean)
   Dim TglNow As Date
   Dim tglSekarang As String
   Dim HasilKey As String
   TglNow = Now()
   tglSekarang = ConvertToTglUnix(xtime(Now))
   If bypass = True Then
        Call CallSetting(True)
   Else
        Call CallSetting(False)
   End If
   HasilKey = CreateSignature(Me.ConsId, Me.SecretKey, tglSekarang)
   Dim req  As WinHttp.WinHttpRequest
   Set req = New WinHttp.WinHttpRequest
  
   
 
    Select Case Me.Method
        Case "GET"
            req.Open "GET", Url, False
        Case "POST"
            req.Open "POST", Url, False
        Case "PUT"
            req.Open "PUT", Url, False
        Case "DELETE"
            req.Open "DELETE", Url, False
   End Select
   
   With req
      '  .Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 13056
        .SetRequestHeader "X-Cons-ID", ConsId
        .SetRequestHeader "X-Timestamp", tglSekarang
        .SetRequestHeader "X-Signature", HasilKey
        If Len(body) < 1 Then
            .Send
        Else
            .SetRequestHeader "Content-Type", "Application/x-www-form-urlencoded"
            .Send Trim(body)
        End If
   End With
   
   Me.status = req.status
   If Not req.status = "200" Then
        Call MsgBox("BPJS Error")
        Me.ResponseText = ""
    Else
        Me.ResponseText = req.ResponseText
    End If

End Function
Public Function Send2(Optional body As String)
   Dim TglNow As Date
   Dim tglSekarang As String
   Dim HasilKey As String
   TglNow = Now()
   tglSekarang = ConvertToTglUnix(xtime(Now))
   Call CallSetting
   HasilKey = CreateSignature(Me.ConsId, Me.SecretKey, tglSekarang)
   
   Dim req As WinHttp.WinHttpRequest
   Set req = New WinHttp.WinHttpRequest
   With req
        .Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 13056
        req.Open "POST", Url, False
        req.SetRequestHeader "X-Cons-ID", ConsId
        req.SetRequestHeader "X-Timestamp", tglSekarang
        req.SetRequestHeader "X-Signature", HasilKey
        req.SetRequestHeader "Content-Type", "Application/x-www-form-urlencoded"
        req.Send body
   End With
   
   
   
   Me.status = req.status
   If Not req.status = "200" Then
        Call MsgBox("BPJS Error")
        Me.ResponseText = ""
    Else
        Me.ResponseText = req.ResponseText
    End If

End Function
Private Function CreateSignature(ConsId As String, scretKey As String, timeStamp As String) As String
    Dim fconsId As String
    Dim fSecretKey As String
    Dim SignatureKey As String
    Dim HS256 As HS256
        
    Set HS256 = New HS256
    fconsId = ConsId
    fSecretKey = scretKey
    
    Dim bitToSign() As Byte
    
    bitToSign = HS256.ToUTF8(fconsId & "&" & timeStamp)
    
    Dim bitSecretKey() As Byte
    bitSecretKey = HS256.ToUTF8(fSecretKey)
    
    HS256.InitHmac bitSecretKey
    
    Dim bitSignature() As Byte
    bitSignature = HS256.HmacSha256(bitToSign)
    
    Dim base64Signature As String
    base64Signature = HS256.Encode(bitSignature, edfBase64, efNoFolding)
    CreateSignature = base64Signature
    
End Function
Private Function ConvertToTglUnix(tgl As Date) As String
    ConvertToTglUnix = (tgl - DateSerial(1970, 1, 1)) * 86400
End Function
Private Function xtime(time As Date) As String
Dim the_date As Date
Dim system_time As SYSTEMTIME
Dim local_file_time As FILETIME
Dim utc_file_time As FILETIME

    ' Get the local time.
    the_date = CDate(time)

    ' Convert it into a SYSTEMTIME.
    DateToSystemTime the_date, system_time

    ' Convert it to a FILETIME.
    SystemTimeToFileTime system_time, local_file_time

    ' Convert it to a UTC time.
    LocalFileTimeToFileTime local_file_time, utc_file_time

    ' Convert it to a SYSTEMTIME.
    FileTimeToSystemTime utc_file_time, system_time

    ' Convert it to a Date.
    SystemTimeToDate system_time, the_date

    ' Display the result.
    xtime = Format$(the_date)
End Function

' Convert a Date into a SYSTEMTIME.
Private Sub DateToSystemTime(ByVal the_date As Date, ByRef system_time As SYSTEMTIME)
    With system_time
        .wYear = Year(the_date)
        .wMonth = Month(the_date)
        .wDay = Day(the_date)
        .wHour = Hour(the_date)
        .wMinute = Minute(the_date)
        .wSecond = Second(the_date)
    End With
End Sub
'Convert a SYSTEMTIME into a Date.
Private Sub SystemTimeToDate(system_time As SYSTEMTIME, ByRef the_date As Date)
    With system_time
        '' The following doesn't work internationally.
        'the_date = CDate( _
        '    Format$(.wMonth) & "/" & _
        '    Format$(.wDay) & "/" & _
        '    Format$(.wYear) & " " & _
        '    Format$(.wHour) & ":" & _
        '    Format$(.wMinute, "00") & ":" & _
        '    Format$(.wSecond, "00"))

        ' Thanks to Justin Starnes for the
        ' following internationalized version.
        the_date = DateSerial(.wYear, .wMonth, .wDay) + _
                   TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Sub
Public Function CallSetting(Optional bypass As Boolean)
    Dim setting As csetting
    Set setting = New csetting
    Call setting.GetData
    Me.ConsId = setting.ConsId
    Me.SecretKey = setting.SecretKey
    If bypass = True Then
        Me.Url = Me.EndPoint
    Else
        Me.Url = setting.urlWebService & Me.EndPoint
    End If
End Function


