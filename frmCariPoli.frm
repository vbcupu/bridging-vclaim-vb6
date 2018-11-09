VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCariPoli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARI POLI"
   ClientHeight    =   6750
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6285
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   6240
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
      Height          =   5415
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9551
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdNamaPoli 
      Caption         =   "Cari"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtNamaPoili 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "NAMA POLI"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmCariPoli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormPengirim As String
Dim vclaim As New vclaim
Dim setting As New csetting
Dim i As Integer
Private Sub cmdNamaPoli_Click()
    Set vclaim = New vclaim
    Call vclaim.refPoli(txtNamaPoili.Text)
    If vclaim.ServerCode = "200" Then
        fg1.rows = vclaim.HasilKode.Count + 1
        fg1.Cols = 3
        fg1.ColWidth(1) = 1500
        fg1.ColWidth(2) = 6000
    
        For i = 1 To vclaim.HasilKode.Count
            fg1.TextMatrix(i, 1) = vclaim.HasilKode.Item(i)
            fg1.TextMatrix(i, 2) = vclaim.HasilKet.Item(i)
        Next i
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub



Private Sub fg1_DblClick()
 If fg1.rows = 2 Then
        If fg1.TextMatrix(1, 1) = "" Then Exit Sub
End If
 If FormPengirim = "frmUbahJenisPasienBPJS" Then
    frmUbahJenisPasienBPJS.txtKdPoli.Text = fg1.TextMatrix(fg1.Row, 1)
    frmUbahJenisPasienBPJS.txtNamaPoli.Text = fg1.TextMatrix(fg1.Row, 2)
    Unload Me
 ElseIf FormPengirim = "frmBuatRujukanBPJS" Then
    frmBuatRujukanBPJS.txtKdPoli.Text = fg1.TextMatrix(fg1.Row, 1)
    frmBuatRujukanBPJS.txtNamaPoli.Text = fg1.TextMatrix(fg1.Row, 2)
End If

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
    fg1.rows = 2
    fg1.Cols = 3
    fg1.ColWidth(1) = 1500
    fg1.ColWidth(2) = 6000
    fg1.TextMatrix(0, 1) = "KODE POLI"
    fg1.TextMatrix(0, 2) = "NAMA POLI"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCariPoli = Nothing
End Sub

Private Sub txtNamaPoili_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    cmdNamaPoli.SetFocus
    'End If
    If Len(txtNamaPoili.Text) > 3 Then
        Call cmdNamaPoli_Click
    End If
End Sub
