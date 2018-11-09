VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCariFaskes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARI FASKES"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Frame fraHasilFaskes 
      Caption         =   "HASIL PENCARIAN"
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   6975
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8705
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CARI FASKES"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtNamaFaskes 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   4815
      End
      Begin VB.ComboBox cbFaskes 
         Height          =   360
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Faskes"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "JENIS FASKES"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCariFaskes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormPengirim As String
Dim vclaim As vclaim
Dim setting As csetting
Private Sub cmdCari_Click()
Call vclaim.CariFaskes(txtNamaFaskes.Text, IIf(cbFaskes.Text = "FASKES 1", "1", "2"))
If vclaim.ServerCode = "200" Then
    fg1.Cols = 3
    fg1.rows = vclaim.HasilKode.Count + 1
    For i = 1 To fg1.rows - 1
       fg1.TextMatrix(i, 1) = vclaim.HasilKode.Item(i)
       fg1.TextMatrix(i, 2) = vclaim.HasilKet.Item(i)
    Next i
End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub fg1_DblClick()
    If Me.FormPengirim = "frmCreateSEP" Then
        frmCreateSEP.txtPPkPelayananan.Text = fg1.TextMatrix(fg1.Row, 1)
        frmCreateSEP.txtnmPPKPelayanan.Text = fg1.TextMatrix(fg1.Row, 2)
        Unload Me
    ElseIf FormPengirim = "frmBuatRujukanBPJS" Then
        frmBuatRujukanBPJS.txtKdFaskes.Text = fg1.TextMatrix(fg1.Row, 1)
        frmBuatRujukanBPJS.txtNamaFaskes.Text = fg1.TextMatrix(fg1.Row, 2)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set vclaim = New vclaim
    cbFaskes.AddItem "FASKES 1"
    cbFaskes.AddItem "FASKES 2"
    Call centerForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCariFaskes = Nothing
End Sub
