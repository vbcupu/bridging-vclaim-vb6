VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "csetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' Class:        csetting.cls
' FILENAME:     csetting.cls
' AUTHOR:       Martin Eko Setiawan
' CREATED:      19/02/2018
' COPYRIGHT:    @ 2018 martin eko setiawan for www.martinsetiawan.com
'
' DESCRIPTION:
' Class Setting data untuk aplikasi bridging VClaim dengan visual basic 6
' message. Baca keseluruhan class ini untuk lebih detailnya
'
' Ini adalah class yang gratis tis, tanpa biaya, Dengan Catatan:

' 1. Kode ini boleh didistribusikan ulang ke source code dan aplikasi lain,
' 2. Tolong author nama saya jangan dihapus ya, kalau mau ditambah nggak apa-apa, kasih aja di bawah atau diatas nama saya
' 3. Karena Source Code ini saya berikan dengan gratis, maka saya berharap source code ini tidak dipakai oleh perusahaan developer berbayar. Karena perusahaan developer harusnya bisa membuat/meresearch API sendiri.
' 4. Saya tidak menjamin aplikasi ini bebas bug, tapi aplikasi ini saya pakai di kerjaan juga, jadi InsyaAllah bebas bug

' Bila kamu mengalami kesulitan dalam implementasi class ini, silahkan berkonsultasi kapan saja, dengan catatan:
' 1. Jangan meminta jawaban secepatnya, karena mungkin saya sedang sibuk.
' 2. Jangan meminta saya bisa menyelesaikan masalahmu, kalau mungkin masalahnya tidak bisa saya selesaikan. Bagaimanapun juga saya cuma programmer biasa yang masih kelas cupu


' Web Site:  www.martinsetiawan.com
' E-mail:vbcupu@gmail.com/martinsetiawan@gmail.com
' Whatsapp: Silahkan lihat di web

'*******************************************************************************
Public ConsId As String
Public SecretKey As String
Public urlWebService As String
Public NoPPK As String
Public NamaPPK As String


Public Function GetData()
    Call getKoneksi
    Call msubrec(rs, "Select * From setting")
    If rs.EOF = True Or rs.BOF = True Then Call MsgBox("Setting Masih Kosong"): Exit Function
    ConsId = rs.fields("consid").Value
    SecretKey = rs.fields("secretkey").Value
    urlWebService = rs.fields("urlapibpjs").Value
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    NoPPK = rs.fields("noppkrs").Value
    NamaPPK = rs.fields("namappkrs").Value
End Function

