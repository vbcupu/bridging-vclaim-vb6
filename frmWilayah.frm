VERSION 5.00
Begin VB.Form frmWilayah 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "KECAMATAN"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   8175
      Begin VB.ComboBox cbKecamatan 
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "Kode Kecamatan: "
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "KOTA"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   8175
      Begin VB.ComboBox cbKota 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Kode Kota: "
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PROPINSI"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.ComboBox cbPropinsi 
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Propinsi: "
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmWilayah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim
Public FormPengirim As String
Dim kdPropinsi As String
Dim nmPropinsi As String
Dim kdKota As String
Dim nmKota As String
Dim kdKecamatan As String
Dim nmKecamatan As String

Private Sub cbKecamatan_Click()
    If Not cbKecamatan.Text = "" Then
        Label3.Caption = "Kode Kecamatan: " & Left(cbKecamatan.Text, 4)
        kdKecamatan = Left(cbKecamatan.Text, 4)
        nmKecamatan = Right(cbKecamatan.Text, Len(cbKecamatan.Text) - 5)
    End If
End Sub

Private Sub cbKecamatan_GotFocus()
    cbKecamatan.Clear
    vclaim.refkecamatan (Left(cbKota.Text, 4))
    If vclaim.ServerCode = "200" Then
        For i = 1 To vclaim.HasilKode.Count
            cbKecamatan.AddItem vclaim.HasilKode.Item(i) & "-" & vclaim.HasilKet.Item(i)
        Next i
    End If
End Sub

Private Sub cbKota_Click()
    If Not cbKota.Text = "" Then
        Label2.Caption = "Kode Kota: " & Left(cbKota.Text, 4)
        kdKota = Left(cbKota.Text, 4)
        nmKota = Right(cbKota.Text, Len(cbKota.Text) - 5)
    End If
End Sub

Private Sub cbKota_GotFocus()
    cbKota.Clear
    vclaim.refkota (Left(cbPropinsi.Text, 2))
    If vclaim.ServerCode = "200" Then
        For i = 1 To vclaim.HasilKode.Count
            cbKota.AddItem vclaim.HasilKode.Item(i) & "-" & vclaim.HasilKet.Item(i)
        Next i
    End If
End Sub

Private Sub cbPropinsi_Click()
    If Not cbPropinsi.Text = "" Then
        Label1.Caption = "Kode Propinsi: " & Left(cbPropinsi.Text, 2)
        kdPropinsi = Left(cbPropinsi.Text, 2)
        nmPropinsi = Right(cbPropinsi.Text, Len(cbPropinsi.Text) - 3)
    End If
End Sub

Private Sub cbPropinsi_GotFocus()
    cbPropinsi.Clear
    vclaim.refPropinsi
    If vclaim.ServerCode = "200" Then
        For i = 1 To vclaim.HasilKode.Count
            cbPropinsi.AddItem vclaim.HasilKode.Item(i) & "-" & vclaim.HasilKet.Item(i)
        Next i
    End If
End Sub

Private Sub cmdTutup_Click()
    If FormPengirim = "frmCreateSEP" Then
        frmCreateSEP.txtKdPropinsi.Text = kdPropinsi
        frmCreateSEP.txtNamaPropinsi.Text = nmPropinsi
        frmCreateSEP.txtKdKabupaten.Text = kdKota
        frmCreateSEP.txtNamaKabupaten.Text = nmKota
        frmCreateSEP.txtKdKecamatan1.Text = kdKecamatan
        frmCreateSEP.txtNamaKecamatan1.Text = nmKecamatan
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    Set vclaim = New vclaim
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmWilayah = Nothing
End Sub
