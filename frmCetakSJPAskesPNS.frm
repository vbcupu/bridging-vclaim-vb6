VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSJPAskesPNS 
   Caption         =   "Medifirst-2000"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCetakSJPAskesPNS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakSJPAskesPNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NoSEP As String
Public NoPendaftaran As String
Public NoCM As String
Public kdPenjamin As String
Public CetakLangsung As Boolean
Dim Report1 As New crCetakSJPBPJS

Private Sub Form_Load()

'End If
'mBPJSMode = True
'Exit Sub

'hell:
'    Set rs = Nothing
'    strSQL = ""
'    Set frmCetakSJPAskesPNS = Nothing
'    Screen.MousePointer = vbDefault
'    'Call msubPesanError
   ' Unload Me
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
strCetak = ""
'Set frmCetakSJPAskesPNS = Nothing
Set rs = Nothing
strSQL = ""
Set frmCetakSJPAskesPNS = Nothing
'Exit Sub
End Sub
Public Sub CetakSEP()
    Call getKoneksi
    Me.WindowState = 2
    strSQL = ""
    
    If CetakLangsung = True Then
        strSQL = "Select * From detailbridgingVclaim where hasilnosep = '" & Me.NoSEP & "'"
        Call msubrec(rs, strSQL)
        If rs.EOF = False Or rs.BOF = False Then
           With Report1
               .txtNoSEP.SetText rs.fields("hasilnoSEP").Value
               .txtTanggalSJP.SetText rs.fields("hasiltglsep").Value
               .txtNamaPasien.SetText rs.fields("hasilpesertanama").Value & " (" & rs.fields("hasilpesertanomr").Value & ")"
               .txtTglLahir.SetText rs.fields("hasilpesertatgllahir").Value
               .txtJK.SetText rs.fields("hasilpesertakelamin")
               .txtPoliTujuan.SetText rs.fields("hasilpoli")
               .txtAsalFaskesTK1.SetText IIf(rs.fields("rujukannmppkrujukan").Value = "", "RSUD KRMT WONGSONEGORO", rs.fields("rujukannmppkrujukan").Value)
               .txtDiagnosaAwal.SetText rs.fields("hasildiagnosa").Value
               .txtCatatan.SetText rs.fields("hasilcatatan").Value
               .txtNomorKartuAskes.SetText rs.fields("nokartu").Value
               .txtPeserta.SetText rs.fields("hasilpesertajnspeserta").Value
               .txtCOB.SetText IIf(IsNull(rs.fields("hasilcob").Value) = True, "", rs.fields("hasilcob").Value)
               .txtJnsRawat.SetText rs.fields("hasiljnspelayanan").Value
               .txtKlsRawat.SetText rs.fields("hasilpesertahakKelas").Value
               .txtNoTelpon.SetText rs.fields("notelpon").Value
               If IsNull(rs.fields("hasilpenjamin").Value) = False Then
                    If Not rs.fields("hasilpenjamin") = "-" Then
                        .txtlakaYa.SetText "Ya"
                        .txtPenjaminLaka.SetText rs.fields("hasilpenjamin").Value
                    End If
               End If
               
               .PrintOut False
               Unload Me
            End With
        Else
            Call MsgBox("SEP Tidak ditemukan")
        End If
    Else
        strSQL = "Select * From detailbridgingVclaim where hasilnosep = '" & Me.NoSEP & "'"
        Call msubrec(rs, strSQL)
        If rs.EOF = False Or rs.BOF = False Then
           With Report1
               .txtNoSEP.SetText rs.fields("hasilnoSEP").Value
               .txtTanggalSJP.SetText rs.fields("hasiltglsep").Value
               .txtNamaPasien.SetText rs.fields("hasilpesertanama").Value & " (" & rs.fields("hasilpesertanomr").Value & ")"
               .txtTglLahir.SetText rs.fields("hasilpesertatgllahir").Value
               .txtJK.SetText rs.fields("hasilpesertajnsKelamin")
               .txtPoliTujuan.SetText rs.fields("hasilpoli")
               .txtAsalFaskesTK1.SetText IIf(rs.fields("rujukannmppkrujukan").Value = "", "RSUD KRMT WONGSONEGORO", rs.fields("rujukannmppkrujukan").Value)
               .txtDiagnosaAwal.SetText rs.fields("hasildiagnosa").Value
               .txtCatatan.SetText rs.fields("hasilcatatan").Value
               .txtNomorKartuAskes.SetText rs.fields("nokartu").Value
               .txtPeserta.SetText rs.fields("hasilpesertajnspeserta").Value
               .txtCOB.SetText IIf(IsNull(rs.fields("hasilcob").Value) = True, "", rs.fields("hasilcob").Value)
               .txtJnsRawat.SetText rs.fields("hasiljnspelayanan").Value
               .txtKlsRawat.SetText rs.fields("hasilpesertahakKelas").Value
               .txtNoTelpon.SetText rs.fields("notelpon").Value
               If IsNull(rs.fields("hasilpenjamin").Value) = False Then
                    If Not rs.fields("hasilpenjamin") = "" Then
                        .txtlakaYa.SetText "Ya"
                        .txtPenjaminLaka.SetText rs.fields("hasilpenjamin").Value
                    End If
               End If
            End With
            
            With CRViewer1
                .EnableExportButton = True
                .EnableGroupTree = True
                .ReportSource = Report1
                .ViewReport
                .Zoom 1
            End With
            Screen.MousePointer = vbDefault
        Else
          Call MsgBox("SEP Tidak ditemukan")
        End If
    End If
End Sub
'Public Sub CetakSJP()
'
'
'    Call openConnection
'    Me.WindowState = 2
'    strSQL = ""
'
'    If strCetak = "Langsung" Then
'
'
'    strSQL = "SELECT dbo.AsuransiPasien.IdPenjamin, dbo.AsuransiPasien.IdAsuransi, dbo.AsuransiPasien.NoCM, dbo.AsuransiPasien.NamaPeserta, " & _
'             "dbo.AsuransiPasien.IDPeserta, dbo.AsuransiPasien.KdGolongan, dbo.AsuransiPasien.TglLahir, dbo.AsuransiPasien.Alamat, " & _
'             "dbo.AsuransiPasien.IdPenjamin, dbo.Penjamin.NamaPenjamin AS NamaPerusahaan, dbo.PemakaianAsuransi.NoPendaftaran, dbo.PemakaianAsuransi.KdKelasDiTanggung, dbo.PemakaianAsuransi.TglSJP, dbo.PemakaianAsuransi.NoSJP " & _
'             "FROM dbo.AsuransiPasien INNER JOIN " & _
'             "dbo.Penjamin ON dbo.AsuransiPasien.IdPenjamin = dbo.Penjamin.IdPenjamin INNER JOIN " & _
'             "dbo.PemakaianAsuransi ON dbo.AsuransiPasien.IdPenjamin = dbo.PemakaianAsuransi.IdPenjamin AND " & _
'             "dbo.AsuransiPasien.IdAsuransi = dbo.PemakaianAsuransi.IdAsuransi And dbo.AsuransiPasien.NoCM = dbo.PemakaianAsuransi.NoCM " & _
'             "WHERE (dbo.AsuransiPasien.NoCM = '" & Me.NoCM & "') AND (dbo.AsuransiPasien.IdPenjamin = '" & Me.kdPenjamin & "') " & _
'             "AND  (dbo.PemakaianAsuransi.NoPendaftaran LIKE '%" & Me.NoPendaftaran & "%')"
'
'    Set rs = Nothing
'    Call msubRecFO(rs, strSQL)
'
'    If rs.fields("idPenjamin") = "0000000144" Then
'        If mBPJSMode = True Then
'            strSQL = "Select * From DetailRequestSEPBPJS where NoPendaftaran = '" & Me.NoPendaftaran & "'"
'            Call msubRecFO(rs, strSQL)
'            If rs.EOF = False Or rs.BOF = False Then
'                With Report1
'                    .txtNoSEP.SetText rs.fields("hasilnoSEP").Value
'                    .txtTanggalSJP.SetText rs.fields("hasiltglsep").Value
'                    .txtNamaPasien.SetText rs.fields("hasilpesertanama").Value & " (" & rs.fields("hasilpesertanomr").Value & ")"
'                    .txtTglLahir.SetText rs.fields("hasilpesertatgllahir").Value
'                    .txtJK.SetText rs.fields("hasilpesertakelamin")
'                    .txtPoliTujuan.SetText rs.fields("hasilpoli")
'                    .txtAsalFaskesTK1.SetText IIf(rs.fields("rujukannmppkrujukan").Value = "", "RSUD KRMT WONGSONEGORO", rs.fields("rujukannmppkrujukan").Value)
'                    .txtDiagnosaAwal.SetText rs.fields("hasildiagnosa").Value
'                    .txtCatatan.SetText rs.fields("hasilcatatan").Value
'                    .txtNomorKartuAskes.SetText rs.fields("nokartu").Value
'                    .txtPeserta.SetText rs.fields("hasilpesertajnspeserta").Value
'                    .txtCOB.SetText IIf(IsNull(rs.fields("hasilcob").Value) = True, "", rs.fields("hasilcob").Value)
'                    .txtJnsRawat.SetText rs.fields("hasiljnspelayanan").Value
'                    .txtKlsRawat.SetText rs.fields("hasilpesertahakKelas").Value
'                    .PrintOut False
'                    Screen.MousePointer = vbDefault
'                    Unload Me
'                End With
'             Else
'                Call MsgBox("Data SEP Tidak DiTemukan")
'                Exit Sub
'             End If
'             Unload Me
'             Exit Sub
'         End If
'         Else
'         With Report
'        .txtNoCM.SetText rs("NoCM")
'        .txtTanggalSJP.SetText rs("TglSJP") '(typAsuransi.dTglSJP)
'        .txtNamaPasien.SetText rs("NamaPeserta") '(frmUbahJenisPasien.txtNamaPasien)
'        .txtNomorKartuAskes.SetText rs("IdAsuransi") '(typAsuransi.strIdAsuransi)
'        .txtNoUrut.SetText rs("NoSJP") '(typAsuransi.strNoSJP)
'        'If strCetak = "frmRegistrasiAll" Then
'                Set rsum = New ADODB.Recordset
'                strSQL = "select dbo.S_HitungUmur ('" & Format(rs.fields("TglLahir").Value, "yyyy/MM/dd") & "', '" & Format(rs.fields("TglSJP").Value, "yyyy/MM/dd") & "')"
'                rsum.Open strSQL, dbconn
'                .txtUmur.SetText rsum.fields(0).Value
'                Set rsum = Nothing
'                Set rsum = New ADODB.Recordset
'                strSQL = "select JenisKelamin from Pasien where nocm = '" & rs.fields("nocm").Value & "'"
'               rsum.Open strSQL, dbconn
'               If rsum.fields(0).Value = "L" Then
'                   .txtJK.SetText "Laki-laki"
'               ElseIf rsum.fields(0).Value = "P" Then
'                   .txtJK.SetText "Perempuan"
'               End If
'
'            .txtAlamat.SetText rs("Alamat")
'            .txtBuktilayanan.SetText (" ")
'             .PrintOut False
'            'If strCetak = "frmRegistrasiAll" Then .PrintOut False
'
'        End With
'        Screen.MousePointer = vbDefault
'        Unload Me
'   End If
'
'
'Else
'    strSQL = "SELECT dbo.AsuransiPasien.IdPenjamin, dbo.AsuransiPasien.IdAsuransi, dbo.AsuransiPasien.NoCM, dbo.AsuransiPasien.NamaPeserta, " & _
'             "dbo.AsuransiPasien.IDPeserta, dbo.AsuransiPasien.KdGolongan, dbo.AsuransiPasien.TglLahir, dbo.AsuransiPasien.Alamat, " & _
'             "dbo.AsuransiPasien.IdPenjamin, dbo.Penjamin.NamaPenjamin AS NamaPerusahaan, dbo.PemakaianAsuransi.NoPendaftaran, dbo.PemakaianAsuransi.KdKelasDiTanggung, dbo.PemakaianAsuransi.TglSJP, dbo.PemakaianAsuransi.NoSJP, dbo.Pasien.TglLahir " & _
'             "FROM dbo.AsuransiPasien INNER JOIN " & _
'             "dbo.Penjamin ON dbo.AsuransiPasien.IdPenjamin = dbo.Penjamin.IdPenjamin INNER JOIN " & _
'             "dbo.PemakaianAsuransi ON dbo.AsuransiPasien.IdPenjamin = dbo.PemakaianAsuransi.IdPenjamin AND " & _
'             "dbo.AsuransiPasien.IdAsuransi = dbo.PemakaianAsuransi.IdAsuransi And dbo.AsuransiPasien.NoCM = dbo.PemakaianAsuransi.NoCM Inner Join Pasien on AsuransiPasien.NoCM = Pasien.NoCM " & _
'             "WHERE (dbo.AsuransiPasien.NoCM = '" & Me.NoCM & "') AND (dbo.AsuransiPasien.IdPenjamin = '" & Me.kdPenjamin & "') " & _
'             "AND  (dbo.PemakaianAsuransi.NoPendaftaran LIKE '%" & Me.NoPendaftaran & "%')"
'
'    Set rs = Nothing
'    Call msubRecFO(rs, strSQL)
'
'    If rs.EOF = True Or rs.BOF = True Then Exit Sub
'    If rs.fields("idPenjamin") = "0000000144" Then
'
'            strSQL = "Select * From DetailRequestSEPBPJS where NoPendaftaran = '" & Me.NoPendaftaran & "'"
'            Call msubRecFO(rs, strSQL)
'            If rs.EOF = False Or rs.BOF = False Then
'                With Report1
'                    .txtNoSEP.SetText rs.fields("hasilnoSEP").Value
'                    .txtTanggalSJP.SetText rs.fields("hasiltglsep").Value
'                    .txtNamaPasien.SetText rs.fields("hasilpesertanama").Value & " (" & rs.fields("hasilpesertanomr").Value & ")"
'                    .txtTglLahir.SetText rs.fields("hasilpesertatgllahir").Value
'                    .txtJK.SetText rs.fields("hasilpesertakelamin")
'                    .txtPoliTujuan.SetText rs.fields("hasilpoli")
'                    .txtAsalFaskesTK1.SetText IIf(rs.fields("rujukannmppkrujukan").Value = "", "RSUD KRMT WONGSONEGORO", rs.fields("rujukannmppkrujukan").Value)
'                    .txtDiagnosaAwal.SetText rs.fields("hasildiagnosa").Value
'                    .txtCatatan.SetText rs.fields("hasilcatatan").Value
'                    .txtNomorKartuAskes.SetText rs.fields("nokartu").Value
'                    .txtPeserta.SetText rs.fields("hasilpesertajnspeserta").Value
'                    .txtCOB.SetText IIf(IsNull(rs.fields("hasilcob").Value) = True, "", rs.fields("hasilcob").Value)
'                    .txtJnsRawat.SetText rs.fields("hasiljnspelayanan").Value
'                    .txtKlsRawat.SetText rs.fields("hasilpesertahakKelas").Value
'                    .txtNoTelpon.SetText rs.fields("notelpon").Value
'                    .PrintOut False
'                End With
'           Else
'                Call MsgBox("Data SEP Tidak DiTemukan")
'                Exit Sub
'           End If
'
'     With CRViewer1
'        .EnableExportButton = True
'        .EnableGroupTree = True
'        .ReportSource = Report1
'        .ViewReport
'        .Zoom 1
'    End With
'    Screen.MousePointer = vbDefault
'   End If
'    Screen.MousePointer = vbDefault
'   '  Unload Me
'  ' End If
'End If
'
'
'End Sub
