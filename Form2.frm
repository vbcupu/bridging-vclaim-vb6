VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRefDPJP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg2 
      Height          =   1935
      Left            =   9360
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3413
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdTUTUP 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   8760
      TabIndex        =   7
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame fraHasilFaskes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HASIL PENCARIAN"
      Height          =   4695
      Left            =   0
      TabIndex        =   3
      Top             =   2640
      Width           =   10215
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
         Height          =   4215
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7435
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CARIDIAGNOSA"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin MSDataListLib.DataCombo dcJenisPelayanan 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtNamaPoli 
         Height          =   405
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   121503745
         CurrentDate     =   43341
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Poli"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tanggal Kontrol"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jenis Pelayanan"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmRefDPJP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormPengirim As String
Dim vclaim As vclaim
Dim KodePoli As String


Private Sub dcSpesialis_GotFocus()
    
End Sub

Private Sub cbPoli_Change()
    
End Sub

Private Sub cbPoli_GotFocus()
    cbPoli.Clear
    Call vclaim.refp
End Sub

Private Sub cmdCari_Click()
    Dim letak As Integer
    Call vclaim.refDokterDPJP(dcJenisPelayanan.BoundText, Format(DTPicker1.Value, "yyyy-MM-dd"), KodePoli)
    If vclaim.ServerCode = "200" Then
        fg1.rows = vclaim.HasilKode.Count
        fg1.Cols = 3
        For i = 1 To vclaim.HasilKode.Count - 1
            fg1.TextMatrix(i, 1) = vclaim.HasilKode.Item(i)
            fg1.TextMatrix(i, 2) = vclaim.HasilKet.Item(i)
        Next i
    End If
    
    fg1.ColWidth(0) = 500
    fg1.ColWidth(1) = 1000
    fg1.ColWidth(2) = 5000
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub fg1_DblClick()
    If Me.FormPengirim = "frmUbahJenisPasienBPJS" Then
        frmUbahJenisPasienBPJS.txtKdDPJP.Text = fg1.TextMatrix(fg1.Row, 1)
        frmUbahJenisPasienBPJS.txtNamaDPJP.Text = fg1.TextMatrix(fg1.Row, 2)
        Unload Me
    End If
End Sub

Private Sub fg2_DblClick()
     KodePoli = fg2.TextMatrix(fg2.Row, 1)
     txtNamaPoli.Text = fg2.TextMatrix(fg2.Row, 2)
     fg2.Visible = False
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Now
    Set vclaim = New vclaim
    Call centerForm(Me, MDIForm1)
    'cbJenisPelayanan.Clear
    'cbJenisPelayanan.AddItem "1-Rawat Inap"
    'cbJenisPelayanan.AddItem "2-Rawat Jalan"
    Call msubdcSource(dcJenisPelayanan, rs, "Select * From JnspelayananBPJS")
    fillDPJP
End Sub
Private Sub fillDPJP()
   ' vclaim.refSpesialis
   ' If vclaim.ServerCode = "200" Then
   '     cbSpesialis.Clear
   '     For i = 1 To vclaim.HasilKode.Count - 1
   '         cbSpesialis.AddItem vclaim.HasilKode.Item(i) & "-" & vclaim.HasilKet.Items(i)
   '     Next i
   ' End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRefDPJP = Nothing
End Sub

Private Sub txtNamaPoli_Change()
    If Len(txtNamaPoli.Text) > 3 Then
        Call vclaim.refPoli(txtNamaPoli.Text)
        If vclaim.HasilKode.Count <> 0 Then
            fg2.Left = 1920
            fg2.Top = 1800
            fg2.Visible = True
            fg2.Cols = 3
            fg2.TextMatrix(0, 1) = "KODE"
            fg2.TextMatrix(0, 2) = "NAMA POLI"
            fg2.ColWidth(0) = 300
            fg2.ColWidth(1) = 1000
            fg2.ColWidth(2) = 2500
            fg2.rows = vclaim.HasilKode.Count + 1
            For i = 1 To vclaim.HasilKode.Count
                fg2.TextMatrix(i, 1) = vclaim.HasilKode.Item(i)
                fg2.TextMatrix(i, 2) = vclaim.HasilKet.Item(i)
            Next i
        End If
            
    End If
End Sub
