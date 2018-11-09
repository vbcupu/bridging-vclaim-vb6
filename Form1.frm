VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPropinsi 
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   6240
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg1 
      Height          =   5895
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   10398
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmPropinsi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim setting As cSetting
Dim vclaim As vclaim
Private Sub cmdNamaPoli_Click()
    
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set vclaim = New vclaim
    Call CenterForm(Me)
    
    vclaim.refPropinsi
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

Private Sub Form_Unload(Cancel As Integer)
    Set frmPropinsi = Nothing
End Sub
