VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCariPotensiSuplesi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
      Height          =   3255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5741
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "Cari"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtNoKartu 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "No. Kartu"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmCariPotensiSuplesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormPengirim As String
Dim vclaim As vclaim
Private Sub cmdCari_Click()
    Set vclaim = New vclaim
    Call vclaim.CariPotensiSuplesi(txtNoKartu.Text, Format(Now, "yyyy-MM-dd"))
    If vclaim.ServerCode = "200" Then
        fg1.Cols = 7
        fg1.rows = vclaim.potsuplesiNoRegister.Count + 1
        For i = 1 To vclaim.potsuplesiNoRegister.Count
            fg1.TextMatrix(i, 1) = vclaim.potsuplesiNoRegister.Item(i)
            fg1.TextMatrix(i, 2) = vclaim.potsuplesinoSep.Item(i)
            fg1.TextMatrix(i, 3) = vclaim.potsuplesinoSepAwal.Item(i)
            fg1.TextMatrix(i, 4) = vclaim.potsuplesinoSuratJaminan.Item(i)
            fg1.TextMatrix(i, 5) = vclaim.potsuplesitglKejadian.Item(i)
            fg1.TextMatrix(i, 6) = vclaim.potsuplesitglSep.Item(i)
        Next i
    End If
End Sub
