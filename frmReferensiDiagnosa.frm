VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReferensiDiagnosa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Referensi DIagnosa"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "CARIDIAGNOSA"
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtNamaDiagnosa 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   495
         Left            =   6840
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Faskes"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraHasilFaskes 
      Caption         =   "HASIL PENCARIAN"
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   10215
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8705
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmReferensiDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormPengirim As String
Dim vclaim As vclaim

Private Sub cmdCari_Click()
    Set vclaim = New vclaim
    vclaim.CariDiagnosa (txtNamaDiagnosa.Text)
    If vclaim.ServerCode = "200" Then
        fg1.rows = vclaim.HasilKode.Count
        fg1.Cols = 3
        For i = 1 To vclaim.HasilKode.Count - 1
            fg1.TextMatrix(i, 1) = vclaim.HasilKode.Item(i)
            fg1.TextMatrix(i, 2) = vclaim.HasilKet.Item(i)
        Next i
    End If
    
    fg1.ColWidth(0) = 200
    fg1.ColWidth(1) = 1000
    fg1.ColWidth(2) = 10000
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub fg1_DblClick()
    If FormPengirim = "frmUbahJenisPasienBPJS" Then
        frmUbahJenisPasienBPJS.txtKDDiagnosa.Text = fg1.TextMatrix(fg1.Row, 1)
        frmUbahJenisPasienBPJS.txtNamaDiagnosa.Text = fg1.TextMatrix(fg1.Row, 2)
        Unload Me
    ElseIf FormPengirim = "frmBuatRujukanBPJS" Then
        frmBuatRujukanBPJS.txtKDDiagnosa.Text = fg1.TextMatrix(fg1.Row, 1)
        frmBuatRujukanBPJS.txtNamaDiagnosa.Text = fg1.TextMatrix(fg1.Row, 2)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReferensiDiagnosa = Nothing
End Sub

Private Sub txtNamaDiagnosa_Change()
    If Len(txtNamaDiagnosa.Text) > 3 Then
        Call cmdCari_Click
    End If
End Sub
