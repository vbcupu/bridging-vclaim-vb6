VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRiwayatBridging 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIWAYAT SEP"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "DAFTAR SEP BRIDGING"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      Begin VB.CommandButton Command1 
         Caption         =   "CARI"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   5400
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpTgl 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   113704961
         CurrentDate     =   43413
      End
      Begin VB.OptionButton OptLangsung 
         BackColor       =   &H80000004&
         Caption         =   "Cetak Langsung"
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   4920
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optLihatViewer 
         BackColor       =   &H80000004&
         Caption         =   "Lihat di Viewer"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Cetak"
         Height          =   735
         Left            =   5760
         TabIndex        =   5
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "TUTUP"
         Height          =   495
         Left            =   11160
         TabIndex        =   4
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton cmdHapusSEP 
         Caption         =   "HAPUS SEP"
         Height          =   495
         Left            =   9240
         TabIndex        =   3
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox txtNoPilihSEP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   4920
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid dg1 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   8070
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmRiwayatBridging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim

Private Sub cmdCetak_Click()
    frmCetakSJPAskesPNS.NoSEP = dg1.Columns("nosep").Value
    If OptLangsung.Value = True Then
        frmCetakSJPAskesPNS.CetakLangsung = True
    Else
        frmCetakSJPAskesPNS.CetakLangsung = False
    End If
    frmCetakSJPAskesPNS.CetakSEP
End Sub

Private Sub cmdHapusSEP_Click()
    Set vclaim = New vclaim
    Call vclaim.HapusSEP(txtNoPilihSEP.Text, "userws")
    If vclaim.ServerCode = "200" Then
        strSQL = "Update detailbridgingvclaim set ishapus = 1 where hasilnosep = '" & txtNoPilihSEP.Text & "'"
        dbconn.Execute strSQL
        Call Load
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Call Load
End Sub

Private Sub dg1_Click()
    If dg1.ApproxCount <> 0 Then
        txtNoPilihSEP.Text = dg1.Columns("NoSEP").Value
    End If
End Sub

Private Sub dtpTgl_Change()
    Call Load
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
    Call Load
End Sub
Public Sub Load()
    strSQL = "Select hasiljnspelayanan as JenisPelayanan, tglSEP, hasilnosep as NoSEP, noKartu, hasilpesertanomr as NoCM, hasilpesertanama as Nama, hasilpesertahakkelas as Kelas From detailbridgingvclaim where  (((detailbridgingVclaim.tglcreate)=#" & Format(dtpTgl, "MM/dd/yyyy") & "#)) And ishapus = 0"
    Call msubrec(rs, strSQL)
    Set dg1.DataSource = rs
    dg1.Columns("JenisPelayanan").Width = 1000
    dg1.Columns("tglSEP").Width = 1200
    dg1.Columns("NoSEP").Width = 2000
    dg1.Columns("noKartu").Width = 2000
    dg1.Columns("NoCM").Width = 800
    dg1.Columns("Nama").Width = 2000
    dg1.Columns("kelas").Width = 1500
End Sub
