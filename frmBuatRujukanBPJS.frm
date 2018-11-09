VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBuatRujukanBPJS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUAT RUJUKAN BPJS"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8715
   Begin VB.Frame Frame1 
      Caption         =   "RUJUKAN BPJS"
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtuser 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   25
         Top             =   4200
         Width           =   5895
      End
      Begin VB.CommandButton cmdCariPoli 
         Caption         =   "CARI POLI"
         Height          =   495
         Left            =   7560
         TabIndex        =   24
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtNamaPoli 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2280
         TabIndex        =   23
         Top             =   3720
         Width           =   5175
      End
      Begin VB.TextBox txtKdPoli 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   22
         Top             =   3720
         Width           =   615
      End
      Begin MSDataListLib.DataCombo dcTipeRujukan 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   2760
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtcatatan 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CARI DIAG"
         Height          =   495
         Left            =   7560
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtNamaDiagnosa 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2400
         TabIndex        =   14
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txtKDDiagnosa 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   13
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "TUTUP"
         Height          =   495
         Left            =   5400
         TabIndex        =   11
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   4080
         TabIndex        =   10
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdCariFaskes 
         Caption         =   "CARI FASKES"
         Height          =   495
         Left            =   7560
         TabIndex        =   9
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtNamaFaskes 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2880
         TabIndex        =   8
         Top             =   3120
         Width           =   4575
      End
      Begin VB.TextBox txtKdFaskes 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpTgldiRujuk 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   116064257
         CurrentDate     =   43368
      End
      Begin VB.TextBox txtNoSEP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo dcJenisPelayanan 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label9 
         Caption         =   "USER"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Poli Tujuan"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Tipe Rujukan"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Catatan"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Diagnosa"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Dirujuk Ke"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Jenis Pelayanan"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal dirujuk"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "No SEP"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBuatRujukanBPJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NoPendaftaran As String
Public NoSEP As String
Dim vclaim As vclaim

Private Sub cmdCariFaskes_Click()
    frmCariFaskes.FormPengirim = Me.Name
    frmCariFaskes.Show
End Sub

Private Sub cmdCariPoli_Click()
    frmCariPoli.FormPengirim = Me.Name
    frmCariPoli.Show
End Sub

Private Sub Command1_Click()
    frmReferensiDiagnosa.FormPengirim = Me.Name
    frmReferensiDiagnosa.Show
End Sub

Private Sub cmdSimpan_Click()
    Set vclaim = New vclaim
    Dim user As String
    Dim adoCommand As ADODB.Command
    user = txtuser.Text
    
    Call vclaim.BuatRujukan(txtNoSEP.Text, Format(dtpTgldiRujuk.Value, "yyyy-MM-dd"), txtKdFaskes.Text, dcJenisPelayanan.BoundText, txtcatatan.Text, Trim(txtKDDiagnosa.Text), dcTipeRujukan.BoundText, txtKdPoli.Text, user)
    If vclaim.status = "200" Then
        If Not (vclaim.ServerCode = "200") Then Exit Sub
        Call MsgBox("Rujukan Berhasil di Buat, Data Rujukan :" & vbNewLine & vclaim.HasilJson, vbOKOnly, "RESULT")
        strSQL = "insert into rujukankeluarBPJS_logrequest values (" & _
        "'" & vclaim.RequestJson & "'," & _
        "'" & Me.NoSEP & "'," & _
        "'" & vclaim.Hasil.Item("kodeAsalRujukan") & "'," & _
        "'" & vclaim.Hasil.Item("namaAsalRujukan") & "'," & _
        "'" & vclaim.Hasil.Item("kodeDiagnosa") & "'," & _
        "'" & vclaim.Hasil.Item("namaDiagnosa") & "'," & _
        "'" & vclaim.Hasil.Item("noRujukan") & "'," & _
        "'" & vclaim.Hasil.Item("asuransi") & "'," & _
        "'" & IIf(IsNull(vclaim.Hasil.Item("hakKelas")) = True, "", vclaim.Hasil.Item("hakKelas")) & "'," & _
        "'" & vclaim.Hasil.Item("jnsPeserta") & "'," & _
        "'" & vclaim.Hasil.Item("kelamin") & "'," & _
        "'" & vclaim.Hasil.Item("nama") & "'," & _
        "'" & vclaim.Hasil.Item("noKartu") & "'," & _
        "'" & vclaim.Hasil.Item("nomr") & "'," & _
        "'" & vclaim.Hasil.Item("tglLahir") & "'," & _
        "'" & vclaim.Hasil.Item("kodePoli") & "'," & _
        "'" & vclaim.Hasil.Item("namaPoli") & "'," & _
        "'" & vclaim.Hasil.Item("tglRujukan") & "'," & _
        "'" & vclaim.Hasil.Item("kodeTujuanRujukan") & "'," & _
        "'" & vclaim.Hasil.Item("namaTujukanRujukan") & "')"
        dbconn.Execute strSQL
    Else
        Call MsgBox("Rujukan Gagal di Buat", vbOKOnly, "WARNING")
    End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Form_Load()
     Call centerForm(Me, MDIForm1)
     Call msubdcSource(dcTipeRujukan, rs, "Select kode, nama From TipeRujukanBPJS")
     Call msubdcSource(dcJenisPelayanan, rs, "select * from jnspelayananbpjs")
     txtNoSEP.Text = Me.NoSEP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBuatRujukanBPJS = Nothing
End Sub
