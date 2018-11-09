VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmrujukanbpjs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RUJUKAN KE RSWN"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   16500
   Begin VB.CommandButton Command1 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   15120
      TabIndex        =   7
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   8175
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   16575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
         Height          =   7695
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   13573
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CARI RUJUKAN TANGGAL"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16575
      Begin VB.CommandButton cmdCari 
         Caption         =   "CARI"
         Height          =   495
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptFaskes2 
         Caption         =   "FASKES 2"
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optFaskes1 
         Caption         =   "FASKES 1"
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   116129793
         CurrentDate     =   43388
      End
   End
End
Attribute VB_Name = "frmrujukanbpjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCari_Click()
    Dim vclaim As New vclaim
    Set vclaim = New vclaim
    If optFaskes1.Value = True Then
        Call vclaim.CariRujukankeRSWN(Format(dtpTanggal.Value, "yyyy-mm-dd"), False)
    Else
        Call vclaim.CariRujukankeRSWN(Format(dtpTanggal.Value, "yyyy-mm-dd"), True)
    End If
    
    If vclaim.ServerCode = "200" Then
        fg1.rows = vclaim.RNoRujukan.Count + 1
        fg1.Cols = 10
        
        For i = 0 To vclaim.RNoRujukan.Count
            If i = vclaim.RNoRujukan.Count Then Exit For
            fg1.TextMatrix(i + 1, 1) = vclaim.RNoPeserta.Items(i)
            fg1.TextMatrix(i + 1, 2) = vclaim.RNoRujukan.Items(i)
            fg1.TextMatrix(i + 1, 3) = IIf(IsNull(vclaim.RNoCM.Items(i)) = True, "", vclaim.RNoCM.Items(i))
            fg1.TextMatrix(i + 1, 4) = vclaim.RNama.Items(i)
            fg1.TextMatrix(i + 1, 5) = vclaim.RKodeDiagnosa.Items(i)
            fg1.TextMatrix(i + 1, 6) = IIf(IsNull(vclaim.RNamaDiagnosa.Items(i)) = True, "", vclaim.RNamaDiagnosa.Items(i))
            fg1.TextMatrix(i + 1, 7) = vclaim.RNoPPKPerujuk.Items(i)
            fg1.TextMatrix(i + 1, 8) = vclaim.RNmPPkPerujuk.Items(i)
        Next i
    End If
    
    If Not vclaim.ServerCode = "200" Then Exit Sub
    fg1.TextMatrix(0, 1) = "NoPeserta"
    fg1.TextMatrix(0, 2) = "No Rujukan"
    fg1.TextMatrix(0, 3) = "No CM"
    fg1.TextMatrix(0, 4) = "Nama Peserta"
    fg1.TextMatrix(0, 5) = "Kode Diagnosa"
    fg1.TextMatrix(0, 6) = "NamaDiagnosa"
    fg1.TextMatrix(0, 7) = "PPK Perujuk"
    fg1.TextMatrix(0, 8) = "Nama PPK"
    fg1.ColWidth(1) = 1500
    fg1.ColWidth(2) = 2000
    fg1.ColWidth(3) = 1000
    fg1.ColWidth(4) = 2500
    fg1.ColWidth(5) = 1500
    fg1.ColWidth(6) = 3000
    fg1.ColWidth(7) = 1500
    fg1.ColWidth(8) = 3000
    
    
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
End Sub
