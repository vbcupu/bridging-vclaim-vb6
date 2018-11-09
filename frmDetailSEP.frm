VERSION 5.00
Begin VB.Form frmDetailSEP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail SEP"
   ClientHeight    =   5475
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCari 
      Caption         =   "CARI"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdHapusSEP 
      Caption         =   "Hapus SEP"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtDeskripsi 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   6735
   End
   Begin VB.TextBox txtNopSEP 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "No SEP"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmDetailSEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim

Private Sub cmdCari_Click()
    Set vclaim = New vclaim
    Call vclaim.DetailSEP(txtNopSEP.Text)
    If vclaim.ServerCode = "200" Then
        txtDeskripsi.Text = vclaim.HasilJson
    End If
End Sub

Private Sub cmdHapusSEP_Click()
    Set vclaim = New vclaim
    Call vclaim.HapusSEP(txtNopSEP.Text)
    If vclaim.ServerCode = "200" Then
        Call MsgBox("SEP Berhasil Di Hapus")
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDetailSEP = Nothing
End Sub
