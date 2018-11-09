VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHistoryPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HISTORY PASIEN"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10110
   Begin VB.Frame Frame1 
      Caption         =   "DATA KUNJUNGAN"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.TextBox txtNoKartu 
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "TUTUP"
         Height          =   375
         Left            =   8640
         TabIndex        =   2
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "LIHAT DATA"
         Height          =   495
         Left            =   7440
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
         Height          =   5175
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9128
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker dtpTglMulai 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   115998721
         CurrentDate     =   43410
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   115998721
         CurrentDate     =   43410
      End
      Begin VB.Label Label3 
         Caption         =   "NO KARTU"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TANGGAL"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmHistoryPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim
Private Sub cmdCari_Click()
    Set vclaim = New vclaim
    Call vclaim.historipelayananpasien(txtNoKartu.Text, Format(dtpTglMulai.Value, "yyyy-MM-dd"), Format(dtpTglAkhir, "yyyy-MM-dd"))
     If vclaim.ServerCode = "200" Then
        fg1.rows = vclaim.historysepNosep.Count + 1
        fg1.Cols = 11
        For i = 1 To vclaim.historysepNosep.Count
             fg1.TextMatrix(i, 1) = vclaim.historysepDiagnosa.Item(i)
            fg1.TextMatrix(i, 2) = vclaim.historysepJnspelayanan.Item(i)
            fg1.TextMatrix(i, 3) = vclaim.historysepKelasrawat.Item(i)
            fg1.TextMatrix(i, 4) = vclaim.historysepNamapeserta.Item(i)
            fg1.TextMatrix(i, 5) = vclaim.historysepNokartu.Item(i)
            fg1.TextMatrix(i, 6) = vclaim.historysepNosep.Item(i)
            fg1.TextMatrix(i, 7) = vclaim.historysepNorujukan.Item(i)
            fg1.TextMatrix(i, 8) = vclaim.historysepPoli(i)
            fg1.TextMatrix(i, 9) = IIf(IsNull(vclaim.historysepTglpulangsep.Item(i)), "", vclaim.historysepTglpulangsep.Item(i))
            fg1.TextMatrix(i, 10) = vclaim.historysepTglsep.Item(i)
        Next i
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHistoryPasien = Nothing
End Sub
