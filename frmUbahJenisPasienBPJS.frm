VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUbahJenisPasienBPJS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubah Jenis Pasien BPJS"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   15840
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DATA PASIEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8295
         Begin MSDataListLib.DataCombo dcKelasPelayanan 
            Height          =   315
            Left            =   5520
            TabIndex        =   36
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcHubungan 
            Height          =   315
            Left            =   2160
            TabIndex        =   35
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtNIK 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   20
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtTglLahir 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5520
            TabIndex        =   19
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtAlamat 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Top             =   960
            Width           =   4815
         End
         Begin VB.TextBox txtNamaPasien 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3600
            TabIndex        =   16
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtNoCM 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtNoPendaftaran 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            TabIndex        =   13
            Top             =   240
            Width           =   3855
         End
         Begin MSDataListLib.DataCombo dcJenisPasien 
            Height          =   315
            Left            =   2160
            TabIndex        =   41
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFFF&
            Caption         =   "KELAS RAWAT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   117
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TGL LAHIR"
            Height          =   255
            Left            =   4440
            TabIndex        =   116
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HUBUNGAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label26 
            BackColor       =   &H00FFFFFF&
            Caption         =   "JENIS PASIEN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label25 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NIK/NAMA PASIEN"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ALAMAT"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NOCM/NOPENDAFTARAN"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdCreateSEP 
         Caption         =   "SIMPAN DAN BUAT SEP (BPJS)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         TabIndex        =   112
         Top             =   8400
         Width           =   2175
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "TUTUP"
         Height          =   495
         Left            =   12360
         TabIndex        =   111
         Top             =   8400
         Width           =   975
      End
      Begin VB.CommandButton cmdRubahSEP 
         Caption         =   "RUBAH SEP"
         Enabled         =   0   'False
         Height          =   495
         Left            =   15120
         TabIndex        =   110
         Top             =   7200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtNoTelpon 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Height          =   375
         Left            =   10080
         TabIndex        =   108
         Top             =   7080
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LAKALANTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   8520
         TabIndex        =   73
         Top             =   3120
         Width           =   6975
         Begin VB.CommandButton cmdCall 
            Caption         =   "---"
            Height          =   375
            Left            =   6240
            TabIndex        =   113
            Top             =   960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdCariKemungkinanSuplesi 
            BackColor       =   &H8000000A&
            Caption         =   "---"
            Height          =   315
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   1800
            Width           =   735
         End
         Begin VB.CommandButton cmdCariKecamatan 
            BackColor       =   &H8000000A&
            Caption         =   "---"
            Height          =   255
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox chkSuplesi 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   104
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtNoSuplesi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   1680
            TabIndex        =   101
            Top             =   2160
            Width           =   5175
         End
         Begin VB.TextBox txtPenjamin 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   4440
            TabIndex        =   100
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdCariKabupaten 
            BackColor       =   &H8000000A&
            Caption         =   "---"
            Height          =   255
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   3000
            Width           =   615
         End
         Begin VB.CommandButton cmdCariPropinsi 
            BackColor       =   &H8000000A&
            Caption         =   "---"
            Height          =   255
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox txtKdPropinsi 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   97
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtKdKabupaten 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   96
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtNamaPropinsi 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   2400
            TabIndex        =   95
            Top             =   2640
            Width           =   3015
         End
         Begin VB.TextBox txtNamaKabupaten 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   2400
            TabIndex        =   94
            Top             =   3000
            Width           =   3015
         End
         Begin VB.TextBox txtKdKecamatan1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            TabIndex        =   93
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox txtNamaKecamatan1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   2400
            TabIndex        =   92
            Top             =   3360
            Width           =   3015
         End
         Begin VB.Frame Frame11 
            Caption         =   "Lokasi Laka"
            Height          =   1335
            Left            =   10560
            TabIndex        =   85
            Top             =   2520
            Width           =   6255
            Begin VB.Label Label43 
               Caption         =   "kdPropinsi"
               Height          =   255
               Left            =   240
               TabIndex        =   88
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label42 
               Caption         =   "KdKabupaten"
               Height          =   255
               Left            =   240
               TabIndex        =   87
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label41 
               Caption         =   "KdKecamatan"
               Height          =   255
               Left            =   240
               TabIndex        =   86
               Top             =   960
               Width           =   1095
            End
         End
         Begin VB.TextBox txtKeteranganLaka 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   1680
            TabIndex        =   83
            Top             =   1440
            Width           =   5175
         End
         Begin VB.CheckBox chkAsabri 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ASABRI"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   80
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkTaspen 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "TASPEN"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   79
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chkBPJSKetenagakerjaan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "BPJS KETENAGAKERJAAN"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   78
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkJasaRaharja 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "JASA RAHARJA"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   77
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkLaka 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            TabIndex        =   75
            Top             =   360
            Width           =   375
         End
         Begin MSComCtl2.DTPicker dtpTglKejadian 
            Height          =   315
            Left            =   3000
            TabIndex        =   82
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   116129793
            CurrentDate     =   43341
         End
         Begin VB.Label Label44 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NO SEP SUPLESI"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label45 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SUPLESI"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "PROPINSI"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "KABUPATEN"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "KECAMATAN"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "KETERANGAN"
            Height          =   315
            Left            =   120
            TabIndex        =   84
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "TANGGAL"
            Height          =   255
            Left            =   2040
            TabIndex        =   81
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "PENJAMIN"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "KASUS LAKA"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NOSEP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8400
         TabIndex        =   37
         Top             =   2520
         Width           =   7095
         Begin VB.TextBox txtNoSEP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   240
            TabIndex        =   38
            Text            =   "0"
            Top             =   240
            Width           =   6735
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CARI PASIEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   8400
         TabIndex        =   22
         Top             =   360
         Width           =   7095
         Begin TabDlg.SSTab SSTab1 
            Height          =   1815
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3201
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "NO RUJUKAN"
            TabPicture(0)   =   "frmUbahJenisPasienBPJS.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label27"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label30"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "dcCariAsalRujukan1"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "txtCariNoRujukan"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "cmdCariByNoRujukan"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Command1"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).ControlCount=   6
            TabCaption(1)   =   "NO KARTU"
            TabPicture(1)   =   "frmUbahJenisPasienBPJS.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label28"
            Tab(1).Control(1)=   "Label37"
            Tab(1).Control(2)=   "txtCariNoKartu"
            Tab(1).Control(3)=   "cmdCariByNoKartu"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "NIK"
            TabPicture(2)   =   "frmUbahJenisPasienBPJS.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label29"
            Tab(2).Control(1)=   "txtCariNIK"
            Tab(2).Control(2)=   "cmdCariByNoNIK"
            Tab(2).ControlCount=   3
            TabCaption(3)   =   "CARI RUJUKAN BY NO KARTU"
            TabPicture(3)   =   "frmUbahJenisPasienBPJS.frx":0054
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "cmdCariRujukanByNokartu1rec"
            Tab(3).Control(1)=   "cmdCariRujukanByNokartumulti"
            Tab(3).Control(2)=   "txtCariRujukanByNokartu"
            Tab(3).Control(3)=   "dcAsalRujukanCariByNoKartu"
            Tab(3).Control(4)=   "Label48"
            Tab(3).Control(5)=   "Label49"
            Tab(3).ControlCount=   6
            Begin VB.CommandButton Command1 
               Caption         =   "Command1"
               Height          =   255
               Left            =   3600
               TabIndex        =   136
               Top             =   1320
               Width           =   615
            End
            Begin VB.CommandButton cmdCariRujukanByNokartu1rec 
               Caption         =   "1 Record"
               Height          =   375
               Left            =   -71880
               TabIndex        =   132
               Top             =   540
               Width           =   1815
            End
            Begin VB.CommandButton cmdCariRujukanByNokartumulti 
               Caption         =   "Multi Record"
               Height          =   375
               Left            =   -71880
               TabIndex        =   131
               Top             =   1140
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox txtCariRujukanByNokartu 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   -74640
               TabIndex        =   130
               Top             =   1140
               Width           =   2655
            End
            Begin VB.CommandButton cmdCariByNoNIK 
               Caption         =   "CARI"
               Height          =   495
               Left            =   -73200
               TabIndex        =   33
               Top             =   1020
               Width           =   1455
            End
            Begin VB.TextBox txtCariNIK 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   -73560
               TabIndex        =   31
               Top             =   540
               Width           =   3135
            End
            Begin VB.CommandButton cmdCariByNoKartu 
               Caption         =   "CARI"
               Height          =   375
               Left            =   -72720
               TabIndex        =   30
               Top             =   1380
               Width           =   1215
            End
            Begin VB.TextBox txtCariNoKartu 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   -73200
               TabIndex        =   28
               Top             =   900
               Width           =   2775
            End
            Begin VB.CommandButton cmdCariByNoRujukan 
               Caption         =   "CARI"
               Height          =   375
               Left            =   1920
               TabIndex        =   27
               Top             =   1260
               Width           =   1455
            End
            Begin VB.TextBox txtCariNoRujukan 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1440
               TabIndex        =   25
               Top             =   780
               Width           =   2775
            End
            Begin MSDataListLib.DataCombo dcCariAsalRujukan1 
               Height          =   315
               Left            =   1440
               TabIndex        =   118
               Top             =   420
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648384
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo dcAsalRujukanCariByNoKartu 
               Height          =   315
               Left            =   -73800
               TabIndex        =   133
               Top             =   780
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin VB.Label Label48 
               Caption         =   "Rujukan By No Kartu"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74640
               TabIndex        =   135
               Top             =   420
               Width           =   2295
            End
            Begin VB.Label Label49 
               Caption         =   "FASKES"
               Height          =   255
               Left            =   -74640
               TabIndex        =   134
               Top             =   780
               Width           =   1095
            End
            Begin VB.Label Label37 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "SELALU MANFAATKAN CARI PASIEN UNTUK MEMASTIKAN STATUS KEPESERTAAN"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   -74400
               TabIndex        =   42
               Top             =   420
               Width           =   4575
            End
            Begin VB.Label Label30 
               Caption         =   "FASKES"
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   420
               Width           =   1095
            End
            Begin VB.Label Label29 
               Caption         =   "No NIK"
               Height          =   375
               Left            =   -74760
               TabIndex        =   32
               Top             =   540
               Width           =   1095
            End
            Begin VB.Label Label28 
               Caption         =   "NO KARTU"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74400
               TabIndex        =   29
               Top             =   900
               Width           =   1095
            End
            Begin VB.Label Label27 
               Caption         =   "NO RUJUKAN"
               Height          =   375
               Left            =   240
               TabIndex        =   26
               Top             =   780
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H8000000E&
         Height          =   5655
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   8295
         Begin VB.CheckBox chkPPKRS 
            Caption         =   "PPK = RS"
            Height          =   375
            Left            =   7200
            TabIndex        =   121
            Top             =   4320
            Width           =   1095
         End
         Begin VB.Frame fraDokter 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DAFTAR DOKTER"
            Height          =   2655
            Left            =   8160
            TabIndex        =   119
            Top             =   2520
            Visible         =   0   'False
            Width           =   6735
            Begin MSDataGridLib.DataGrid dgdokter 
               Height          =   2175
               Left            =   120
               TabIndex        =   120
               Top             =   240
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   3836
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
         Begin VB.TextBox txtNamaPPKPerujuk 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            TabIndex        =   114
            Top             =   4320
            Width           =   4095
         End
         Begin VB.CommandButton cmdCariDokter 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "---"
            Height          =   315
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox txtNamaDPJP 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   69
            Top             =   5160
            Width           =   4455
         End
         Begin VB.TextBox txtKdDPJP 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   68
            Top             =   5160
            Width           =   735
         End
         Begin VB.TextBox txtNoKontrol 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   67
            Top             =   4680
            Width           =   5175
         End
         Begin VB.TextBox txtNoPPKPerujuk 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   63
            Top             =   4320
            Width           =   1095
         End
         Begin VB.TextBox txtNoRujukan 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   59
            Text            =   "0"
            Top             =   3240
            Width           =   2775
         End
         Begin VB.CheckBox chkKatarak 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3840
            TabIndex        =   58
            Top             =   2520
            Width           =   375
         End
         Begin VB.CheckBox chkCOB 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   56
            Top             =   2520
            Width           =   375
         End
         Begin VB.CheckBox chkPoliEksekutif 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1560
            TabIndex        =   55
            Top             =   2520
            Width           =   375
         End
         Begin VB.TextBox txtKDPoli 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   52
            Top             =   2880
            Width           =   735
         End
         Begin VB.CommandButton cmdCariPoli 
            BackColor       =   &H8000000A&
            Caption         =   "---"
            Height          =   315
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox txtNamaPoli 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2280
            TabIndex        =   50
            Top             =   2880
            Width           =   4575
         End
         Begin VB.TextBox txtKdDiagnosa 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   47
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdCariDiagnosa 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "---"
            Height          =   285
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtNamaDiagnosa 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   45
            Top             =   2160
            Width           =   4215
         End
         Begin VB.TextBox txtCatatan 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   43
            Text            =   "-"
            Top             =   1800
            Width           =   5415
         End
         Begin VB.TextBox txtNamaPPKPelayanan 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   5
            Text            =   "RSUD KRMT WONGSONEGORO"
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtNoPPK 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Text            =   "1101R024"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtNoKartu 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   240
            Width           =   5415
         End
         Begin MSComCtl2.DTPicker dtpTglSEP 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   49152
            CalendarTitleBackColor=   49152
            Format          =   116129793
            CurrentDate     =   43144
         End
         Begin MSDataListLib.DataCombo dcJenisPelayanan 
            Height          =   315
            Left            =   1560
            TabIndex        =   39
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632319
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcKelasdiTanggung 
            Height          =   315
            Left            =   3240
            TabIndex        =   40
            Top             =   1320
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632319
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtpTglRujukan 
            Height          =   315
            Left            =   4800
            TabIndex        =   61
            Top             =   3240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   116129793
            CurrentDate     =   43144
         End
         Begin MSDataListLib.DataCombo dcAsalRujukanBPJS 
            Height          =   315
            Left            =   1560
            TabIndex        =   66
            Top             =   3600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648384
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "DPJP"
            Height          =   315
            Left            =   120
            TabIndex        =   72
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "NO KONTROL"
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "ASAL RUJUKAN"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "KODE PPK PERUJUK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   4320
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "TGL"
            Height          =   285
            Left            =   4440
            TabIndex        =   62
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "NO RUJUKAN"
            Height          =   285
            Left            =   120
            TabIndex        =   60
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label46 
            BackColor       =   &H00FFFFFF&
            Caption         =   "KATARAK"
            Height          =   255
            Left            =   3000
            TabIndex        =   57
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFFF&
            Caption         =   "COB"
            Height          =   255
            Left            =   2040
            TabIndex        =   54
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "POLI"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label33 
            BackColor       =   &H00FFFFFF&
            Caption         =   "POLI EKSEKUTIF"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DIAG AWAL"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "CATATAN"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000E&
            Caption         =   "Kelas Rawat (Di Tanggung)"
            Height          =   315
            Left            =   3240
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            Caption         =   "Jenis Pelayanan"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            Caption         =   "PPK PELAYANAN"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000E&
            Caption         =   "Tgl SEP"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            Caption         =   "No Kartu"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10080
         TabIndex        =   1
         Text            =   "userws"
         Top             =   7560
         Width           =   3855
      End
      Begin VB.Frame Frame2 
         Caption         =   "TIPE TRANSAKSI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   122
         Top             =   2280
         Width           =   8295
         Begin VB.OptionButton OptKontrolPostRanap 
            Caption         =   "POST RANAP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   129
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton OptRanapBayi 
            Caption         =   "RANAP BAYI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   128
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton optSEPIGD 
            Caption         =   "SEP IGD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   127
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton OptRanap 
            Caption         =   "RANAP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   126
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptIGDTanpaSEP 
            Caption         =   "IGD TANPA SEP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   125
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton OptKlinik 
            Caption         =   "KLINIK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptKontrol 
            Caption         =   "KONTROL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   123
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "USER"
         Height          =   255
         Left            =   8880
         TabIndex        =   109
         Top             =   7560
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "TELEPON"
         Height          =   255
         Left            =   8760
         TabIndex        =   107
         Top             =   7080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmUbahJenisPasienBPJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vclaim As vclaim

Private Sub cmdCariByNoKartu_Click()
 Dim Ins As Integer
    If Len(txtCariNoKartu.Text) > 13 Then
       Call MsgBox("No KArtu Tidak Boleh Lebih Dari 13 Karakter", vbOKOnly, "PERHATIAN")
       Exit Sub
    End If
    If Len(txtCariNoKartu.Text) < 13 Then
       Call MsgBox("No KArtu Tidak Boleh Kurang Dari 13 Karakter", vbOKOnly, "PERHATIAN")
       Exit Sub
    End If
    Set vclaim = New vclaim
    Call vclaim.CariPesertaByNoKartu(txtCariNoKartu.Text)
    If vclaim.ServerCode = "200" Then
        'If Not Me.FormPengirim = "" Then
        '    txtNoKartu.Text = vclaim.Hasil.Item("noKartu")
        '    If Trim(UCase(txtNamaPasien.Text)) <> Trim(UCase(vclaim.Hasil.Item("nama"))) Then
        '        If MsgBox("Ada Perbedaan Nama di data BPJS dan Rumah Sakit, Nama di BPJS :" & vclaim.Hasil.Item("nama") & " Data di Rumah Sakit: " & txtNamaPasien.Text & " Apakah Transaksi akan dihentikan?", vbYesNo) = vbNo Then
        '            txtNamaPasien.Text = vclaim.Hasil.Item("nama")
        '        End If
        '    End If
       '
       '     If Trim(txtNIK.Text) <> Trim(vclaim.Hasil.Item("nik")) Then
       '         If MsgBox("Ada Perbedaan NOMOR NIK di data BPJS dan Rumah Sakit, Nama di BPJS : " & vclaim.Hasil.Item("nik") & " Data di Rumah Sakit: " & IIf(txtNIK.Text = "", "ZONK", txtNIK.Text) & " Apakah Transaksi akan dihentikan?", vbYesNo) = vbNo Then
       '             txtNIK.Text = vclaim.Hasil.Item("nik")
       '         End If
       '     End If
            If Left(Trim(vclaim.Hasil.Item("kodeJenisPeserta")), 3) = "PBI" Then
                dcJenisPasien.BoundText = "02"
                
            Else
                dcJenisPasien.BoundText = "03"
            End If
       '
       ' Else
   '         txtNamaPasien.Text = vclaim.Hasil.Item("nama")
   '         txtNIK.Text = vclaim.Hasil.Item("nik")
  ' End If
        txtNoKartu.Text = vclaim.Hasil.Item("noKartu")
        dcKelasdiTanggung.BoundText = vclaim.Hasil.Item("kodeHakKelas")
        dcHubungan.BoundText = vclaim.Hasil.Item("pisa")
        Hubungan = vclaim.Hasil.Item("pisa")
        dcAsalRujukanBPJS.BoundText = "1"
        txtNoPPKPerujuk.Text = vclaim.Hasil.Item("kodeProvUmum")
        txtNamaPPKPerujuk.Text = vclaim.Hasil.Item("namaProvUmum")
        txtNoCM.Text = IIf(IsNull(vclaim.Hasil.Item("nomr")) = True, "", vclaim.Hasil.Item("nomr"))
        txtNamaPasien.Text = vclaim.Hasil.Item("nama")
        txtNIK.Text = vclaim.Hasil.Item("nik")

    Else
        If Not vclaim.ServerMessage = "" Then
            Call MsgBox(vclaim.ServerMessage)
        End If
    End If
End Sub

Private Sub cmdCariByNoNIK_Click()
   If Len(txtCariNIK.Text) > 16 Then
       Call MsgBox("No KArtu Tidak Boleh Lebih Dari 16 Karakter", vbOKOnly, "PERHATIAN")
       Exit Sub
    End If
    If Len(txtCariNIK.Text) < 16 Then
       Call MsgBox("No KArtu Tidak Boleh Kurang Dari 16 Karakter", vbOKOnly, "PERHATIAN")
       Exit Sub
    End If
  
        Set vclaim = New vclaim
    
  
    
        Call vclaim.CariPesertaByNIK(txtCariNIK.Text)
        If vclaim.ServerCode = "200" Then
            txtNamaPasien.Text = vclaim.Hasil.Item("nama")
            txtNoCM.Text = vclaim.Hasil.Item("mr")
            txtNIK.Text = vclaim.Hasil.Item("nik")
            txtTglLahir.Text = vclaim.Hasil.Item("tglLahir")
            dcHubungan.BoundText = vclaim.Hasil.Item("pisa")
            Hubungan = vclaim.Hasil.Item("pisa")
            dcJenisPasien.BoundText = IIf(Left(vclaim.Hasil.Item("namaJenisPeserta"), 3) = "PBI", "02", "03")
            txtNoKartu.Text = vclaim.Hasil.Item("noKartu")
            txtNoKartu.BackColor = vbYellow
            txtNIK.Text = vclaim.Hasil.Item("nik")
            txtTglLahir.Text = vclaim.Hasil.Item("tglLahir")
            dcKelasdiTanggung.BoundText = vclaim.Hasil.Item("kodeHakKelas")
            txtNoPPKPerujuk.Text = vclaim.Hasil.Item("kodeProvUmum")
            txtNamaPPKPerujuk.Text = vclaim.Hasil.Item("namaProvUmum")
        Else
            Call MsgBox(vclaim.ServerMessage)
        End If

   
End Sub

Private Sub cmdCariByNoRujukan_Click()
 Dim KdDetailAsalRujukan As String
    Dim KdAsalRujukan As String

    Set vclaim = New vclaim
    If dcCariAsalRujukan1.BoundText = "" Then Call MsgBox("Pilih Faskes terlebih Dahulu"): Exit Sub
    If dcCariAsalRujukan1.BoundText = "1" Then
        Call vclaim.CariByRujukanPcare(txtCariNoRujukan.Text)
    Else
        Call vclaim.CariByRujukanRS(txtCariNoRujukan.Text)
    End If
    If vclaim.ServerCode = "200" Then
        Dim a As String
        a = vclaim.HasilJson
        txtNoKartu.Text = vclaim.Hasil.Item("noKartu")
        txtNoCM.Text = vclaim.Hasil.Item("nomr")
        txtNamaPasien.Text = vclaim.Hasil.Item("nama")
        txtNIK.Text = vclaim.Hasil.Item("nik")
        dcKelasdiTanggung.BoundText = vclaim.Hasil.Item("kodeHakKelas")
        dcJenisPasien.BoundText = IIf(Left(vclaim.Hasil.Item("namaJenisPeserta"), 3) = "PBI", "02", "03")
        dcHubungan.BoundText = vclaim.Hasil.Item("pisa")
        txtNoRujukan.Text = txtCariNoRujukan.Text
        dcJenisPelayanan.BoundText = vclaim.Hasil.Item("kodePelayanan")
        txtKdPoli.Text = vclaim.Hasil.Item("kodePoliRujukan")
        txtNamaPoli.Text = vclaim.Hasil.Item("namaPoliRujukan")
        txtNoKartu.Text = vclaim.Hasil.Item("noKartu")
        txtKDDiagnosa.Text = vclaim.Hasil.Item("kodeDiagnosa")
        txtNamaDiagnosa.Text = vclaim.Hasil.Item("namaDiagnosa")
        txtNoPPKPerujuk.Text = vclaim.Hasil.Item("kodeAsalRujukan")
        Hubungan = vclaim.Hasil.Item("pisa")
        dtpTglRujukan.Value = Format(CDate(vclaim.Hasil.Item("tglDirujuk")), "dd/MM/yyyy")
        txtKdPoli.Text = vclaim.Hasil.Item("kodePoliRujukan")
        txtNamaPoli.Text = vclaim.Hasil.Item("namaPoliRujukan")
        txtKDDiagnosa.Text = vclaim.Hasil.Item("kodeDiagnosa")
        txtNamaDiagnosa.Text = vclaim.Hasil.Item("namaDiagnosa")
        Hubungan = vclaim.Hasil.Item("pisa")
        dcAsalRujukanBPJS.BoundText = dcCariAsalRujukan1.BoundText
        txtNoPPKPerujuk.Text = vclaim.Hasil.Item("kodeProvPerujuk")
        txtNamaPPKPerujuk.Text = vclaim.Hasil.Item("namaProvPerujuk")
        txtNoKartu.BackColor = vbYellow
    Else
        Call MsgBox(vclaim.ServerMessage)
    End If
    If dcCariAsalRujukan1.BoundText = "1" Then
        dcAsalRujukanBPJS.BoundText = "1"
    Else
        dcAsalRujukanBPJS.BoundText = "2"
    End If
End Sub

Private Sub cmdCariDiagnosa_Click()
    frmReferensiDiagnosa.FormPengirim = Me.Name
    frmReferensiDiagnosa.Show
End Sub

Private Sub cmdCariDokter_Click()
    frmRefDPJP.FormPengirim = Me.Name
    frmRefDPJP.Show
End Sub

Private Sub cmdCariPoli_Click()
    frmCariPoli.FormPengirim = Me.Name
    frmCariPoli.Show
End Sub

Private Sub cmdCariRujukanByNokartu1rec_Click()
    Set vclaim = New vclaim
    
    If dcAsalRujukanCariByNoKartu.BoundText = "1" Then
        Call vclaim.CariByRujukanByNoKartu1recPcare(txtCariRujukanByNokartu.Text)
    Else
        Call vclaim.CariByRujukanByNoKartu1recRS(txtCariRujukanByNokartu.Text)
    End If
    If vclaim.ServerCode = "200" Then
        Dim a As String
        a = vclaim.HasilJson
        txtNoKartu.Text = vclaim.Hasil.Item("nokartu")
        txtNamaPasien.Text = vclaim.Hasil.Item("nama")
        txtNIK.Text = vclaim.Hasil.Item("nik")
        txtNoKartu.Text = vclaim.Hasil.Item("noKartu")
        dcKelasdiTanggung.BoundText = vclaim.Hasil.Item("kodeHakKelas")
        dcJenisPasien.BoundText = IIf(Left(vclaim.Hasil.Item("namaJenisPeserta"), 3) = "PBI", "20", "19")
        dcHubungan.BoundText = vclaim.Hasil.Item("pisa")
        txtNoRujukan.Text = vclaim.Hasil.Item("noKunjungan")
        dcJenisPelayanan.BoundText = vclaim.Hasil.Item("kodePelayanan")
        txtKdPoli.Text = vclaim.Hasil.Item("kodePoliRujukan")
        txtNamaPoli.Text = vclaim.Hasil.Item("namaPoliRujukan")
        txtNoKartu.Text = vclaim.Hasil.Item("noKartu")
        txtKDDiagnosa.Text = vclaim.Hasil.Item("kodeDiagnosa")
        txtNamaDiagnosa.Text = vclaim.Hasil.Item("namaDiagnosa")
        txtNoPPKPerujuk.Text = vclaim.Hasil.Item("kodeAsalRujukan")
        dtpTglRujukan.Value = Format(CDate(vclaim.Hasil.Item("tglDirujuk")), "dd/MM/yyyy")
        txtKdPoli.Text = vclaim.Hasil.Item("kodePoliRujukan")
        txtNamaPoli.Text = vclaim.Hasil.Item("namaPoliRujukan")
        dcAsalRujukanBPJS.BoundText = vclaim.Hasil.Item("kodePelayanan")
        txtNoPPKPerujuk.Text = vclaim.Hasil.Item("kodeProvPerujuk")
        txtNamaPPKPerujuk.Text = vclaim.Hasil.Item("namaProvPerujuk")
        txtNoRujukan.Text = vclaim.Hasil.Item("noKunjungan")
        dtpTglRujukan.Value = Format(vclaim.Hasil.Item("tglDirujuk"), "dd/MM/yyyy")
        Hubungan = vclaim.Hasil.Item("pisa")
      
        txtNoKartu.BackColor = vbYellow
    Else
        Call MsgBox(vclaim.ServerMessage)
    End If
End Sub

Private Sub cmdCreateSEP_Click()
  If OptKlinik.Value = True Or OptKontrol.Value = True Or optSEPIGD.Value = True Or OptRanap.Value = True Or OptKontrolPostRanap.Value = True Then
        Call BuatSEP
  Else
        'Call SimpanJaminan
        cmdCreateSEP.Enabled = False
        cmdTutup.SetFocus
  End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Call SimpanLog("123", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIForm1)
    Call msubdcSource(dcCariAsalRujukan1, rs, "Select * From asalRujukanBPJS")
    Call msubdcSource(dcAsalRujukanBPJS, rs, "Select * From asalRujukanBPJS")
    Call msubdcSource(dcAsalRujukanCariByNoKartu, rs, "Select * From asalRujukanBPJS")
    Call msubdcSource(dcKelasPelayanan, rs, "Select * From kelasPelayananBPJS")
    Call msubdcSource(dcKelasdiTanggung, rs, "Select * From kelasPelayananBPJS")
    Call msubdcSource(dcHubungan, rs, "select * From hubunganbpjs")
    Call msubdcSource(dcJenisPasien, rs, "select * From jenispasien")
    Call msubdcSource(dcJenisPelayanan, rs, "select * from jnspelayananbpjs")
    dtpTglSEP.Value = Now
    dtpTglKejadian.Value = Now
    dtpTglRujukan.Value = Now
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUbahJenisPasienBPJS = Nothing
End Sub
Private Sub BuatSEP()
Dim nokartu, tglSEP, NoPPK, JenisPelayanan, KelasDiTanggung, NoCM, asalRujukan, tglrujukan, NoRujukan, PPKPerujuk, catatan, Diagnosa, Poli, polieksekutif, cob, katarak, Lakalantas, penjamin, TglKejadian, Keterangan, Suplesi, NoSEPSuplesi, kdPropinsi, kdKabupaten, kdKecamatan, NoKontrol, DPJP, notelpon, UserName As String
    If Len(txtNoSEP.Text) > 10 Then Call MsgBox("No SEP Sudah Ada, Pembuatan SEP Dibatalkan"): Exit Sub
    If Len(txtNoTelpon.Text) < 8 Then Call MsgBox("No Telpon Belum di Isi/Belum diisi Dengan Benar"): Exit Sub
    
    nokartu = txtNoKartu.Text
    tglSEP = Format(dtpTglSEP.Value, "yyyy-mm-dd")
    NoPPK = txtNoPPK.Text
    JenisPelayanan = dcJenisPelayanan.BoundText
    KelasDiTanggung = dcKelasdiTanggung.BoundText
    NoCM = txtNoCM.Text
    
    'isi rujukan
    If Not txtNoRujukan.Text = "" Then
        asalRujukan = dcAsalRujukanBPJS.BoundText
        tglrujukan = Format(dtpTglRujukan.Value, "yyyy-mm-dd")
        PPKPerujuk = txtNoPPKPerujuk.Text
        catatan = txtcatatan.Text
        NoRujukan = txtNoRujukan.Text
        
    Else
        asalRujukan = ""
        tglrujukan = ""
        PPKPerujuk = ""
        catatan = ""
    End If
    
    Diagnosa = Trim(txtKDDiagnosa.Text)
    Poli = Trim(txtKdPoli.Text)
    polieksekutif = IIf(chkPoliEksekutif.Value = 0, "0", "1")
    cob = IIf(chkCOB.Value = 0, "0", "1")
    katarak = IIf(chkKatarak.Value = 0, "0", "1")
    
    'Lakalantas
    If chkLaka.Value = 1 Then
        Lakalantas = "1"
        penjamin = txtPenjamin.Text
        TglKejadian = Format(dtpTglKejadian.Value, "yyyy-mm-dd")
        Keterangan = txtKeteranganLaka.Text
        If chkSuplesi.Value = 1 Then
            Suplesi = "1"
            NoSEPSuplesi = txtNoSuplesi.Text
        Else
            Suplesi = "0"
            NoSEPSuplesi = ""
        End If
        kdPropinsi = txtKdPropinsi.Text
        kdKabupaten = txtKdKabupaten.Text
        kdKecamatan = txtKdKecamatan1.Text
    Else
        Lakalantas = "0"
        penjamin = ""
        TglKejadian = ""
        Keterangan = ""
        Suplesi = "0"
        NoSEPSuplesi = 0
        kdPropinsi = txtKdPropinsi.Text
        kdKabupaten = txtKdKabupaten.Text
        kdKecamatan = txtKdKecamatan1.Text
    End If
    NoKontrol = txtNoKontrol.Text
    DPJP = txtKdDPJP.Text
    notelpon = txtNoTelpon.Text
    UserName = txtuser.Text
    Set vclaim = New vclaim
    Call vclaim.BuatSEP(nokartu, tglSEP, NoPPK, JenisPelayanan, KelasDiTanggung, NoCM, asalRujukan, tglrujukan, NoRujukan, PPKPerujuk, catatan, Diagnosa, Poli, polieksekutif, cob, katarak, Lakalantas, penjamin, TglKejadian, Keterangan, Suplesi, NoSEPSuplesi, kdPropinsi, kdKabupaten, kdKecamatan, NoKontrol, DPJP, notelpon, UserName)
    
    If vclaim.ServerCode = "200" Then
        txtNoSEP.Text = vclaim.Hasil.Item("nosep")
        Call SimpanLog(txtNoPendaftaran.Text, nokartu, tglSEP, NoPPK, JenisPelayanan, KelasDiTanggung, NoCM, asalRujukan, tglrujukan, IIf(NoRujukan = "", "0", NoRujukan), PPKPerujuk, txtNamaPPKPerujuk.Text, catatan, txtKDDiagnosa.Text, Trim(txtNamaDiagnosa.Text), Poli, txtNamaPoli.Text, _
polieksekutif, cob, katarak, Lakalantas, penjamin, TglKejadian, Keterangan, Suplesi, NoSEPSuplesi, kdPropinsi, kdKabupaten, kdKecamatan, IIf(NoKontrol = "", "", NoKontrol), IIf(DPJP = "", "", DPJP), notelpon, UserName, vclaim.Hasil.Item("catatan"), vclaim.Hasil.Item("diagnosa"), vclaim.Hasil.Item("jnspelayanan"), vclaim.Hasil.Item("kelasrawat"), vclaim.Hasil.Item("nosep"), vclaim.Hasil.Item("penjamin"), vclaim.Hasil.Item("asuransi"), vclaim.Hasil.Item("hakkelas"), vclaim.Hasil.Item("jnspeserta"), vclaim.Hasil.Item("kelamin"), vclaim.Hasil.Item("nama"), vclaim.Hasil.Item("nokartu"), vclaim.Hasil.Item("nomr"), IIf(vclaim.Hasil.Item("tgllahir") = "", "", vclaim.Hasil.Item("tgllahir")), IIf(vclaim.Hasil.Item("dinsos") = "", "", vclaim.Hasil.Item("dinsos")), _
IIf(vclaim.Hasil.Item("prolanisprb") = "", "", vclaim.Hasil.Item("prolanisprb")), IIf(vclaim.Hasil.Item("nosktm") = "", "", vclaim.Hasil.Item("nosktm")), vclaim.Hasil.Item("poli"), IIf(vclaim.Hasil.Item("polieksekutif") = "", "", vclaim.Hasil.Item("polieksekutif")), dtpTglSEP.Value)
        
    End If
    

End Sub
Private Function SimpanLog(NoPendaftaran, nokartu, tglSEP, NoPPK, JenisPelayanan, klsrawat, nomr, rujukanasalrujukan, rujukantglrujukan, rujukannorujukan, rujukanppkrujukan, rujukannmppkrujukan, catatan, diagawal, nmdiagAwal, politujuan, nmpolitujuan, eksekutif, cob, katarak, jaminanlakalantas, jaminanpenjaminpenjamin, jaminanpenjamintglkejadian, jaminanpenjaminketerangan, jaminanpenjaminsuplesisuplesi, jaminanpenjaminsuplesinosuplesi, jaminanpenjaminsuplesilokasilakakdpropinsi, jaminanpenjaminsuplesilokasilakakdkabupaten, jaminanpenjaminsuplesilokasilakakdkecamatan, skdpnosurat, skdpdpjp, notelpon, namauser, hasilcatatan, hasildiagnosa, hasiljnspelayanan, hasilkelasrawat, hasilnosep, hasilpenjamin, hasilpesertaasuransi, hasilpesertahakkelas, hasilpesertajnspeserta, hasilpesertajnsKelamin, hasilpesertanama, hasilpesertanokartu, hasilpesertanomr, hasilpesertatgllahir, hasilinformasidinsos, hasilinformasiprolanisPRB, hasilinformasinosktm, hasilpoli, hasilpolieksekutif, hasiltglsep As String)
    Dim rsLog As ADODB.Recordset
    Set rsLog = New ADODB.Recordset
    rsLog.Open "Select * From detailbridgingVclaim", dbconn, adOpenStatic, adLockOptimistic
    With rsLog
        .AddNew
        .fields("NoPendaftaran").Value = NoPendaftaran
        .fields("nokartu").Value = nokartu
        .fields("TglSEP") = tglSEP
        .fields("NoPPK") = NoPPK
        .fields("JenisPelayanan") = JenisPelayanan
        .fields("klsrawat") = klsrawat
        .fields("nomr") = nomr
        .fields("rujukanasalrujukan") = rujukanasalrujukan
        .fields("rujukantglrujukan") = rujukantglrujukan
        .fields("rujukannorujukan") = rujukannorujukan
        .fields("rujukannorujukan") = rujukannorujukan
        .fields("rujukannmppkrujukan") = rujukannmppkrujukan
        .fields("catatan") = catatan
        .fields("diagawal") = diagawal
        .fields("nmdiagAwal") = nmdiagAwal
        .fields("politujuan") = politujuan
        .fields("nmpolitujuan") = nmpolitujuan
        .fields("eksekutif") = eksekutif
        .fields("cob") = cob
        .fields("katarak") = katarak
        .fields("jaminanlakalantas") = jaminanlakalantas
        .fields("jaminanpenjaminpenjamin") = jaminanpenjaminpenjamin
        .fields("jaminanpenjamintglkejadian") = jaminanpenjamintglkejadian
        .fields("jaminanpenjaminketerangan") = jaminanpenjaminketerangan
        .fields("jaminanpenjaminsuplesisuplesi") = jaminanpenjaminsuplesisuplesi
        .fields("jaminanpenjaminsuplesinosuplesi") = jaminanpenjaminsuplesinosuplesi
        .fields("jaminanpenjaminsuplesilokasilakakdpropinsi") = jaminanpenjaminsuplesilokasilakakdpropinsi
        .fields("jaminanpenjaminsuplesilokasilakakdkabupaten") = jaminanpenjaminsuplesilokasilakakdkabupaten
        .fields("jaminanpenjaminsuplesilokasilakakdkecamatan") = jaminanpenjaminsuplesilokasilakakdkecamatan
        .fields("skdpnosurat") = skdpnosurat
        .fields("skdpdpjp") = skdpdpjp
        .fields("notelpon") = notelpon
        .fields("namauser") = namauser
        .fields("hasilcatatan") = hasilcatatan
        .fields("hasildiagnosa") = hasildiagnosa
        .fields("hasiljnspelayanan") = hasiljnspelayanan
        .fields("hasilkelasrawat") = hasilkelasrawat
        .fields("hasilnosep") = hasilnosep
        .fields("hasilpenjamin") = hasilpenjamin
        .fields("hasilpesertaasuransi") = hasilpesertaasuransi
        .fields("hasilpesertahakkelas") = hasilpesertahakkelas
        .fields("hasilpesertajnspeserta") = hasilpesertajnspeserta
        .fields("hasilpesertajnsKelamin") = hasilpesertajnsKelamin
        .fields("hasilpesertanama") = hasilpesertanama
        .fields("hasilpesertanokartu") = hasilpesertanokartu
        .fields("hasilpesertanomr") = hasilpesertanomr
        .fields("hasilpesertatgllahir") = hasilpesertatgllahir
        .fields("hasilinformasidinsos") = hasilinformasidinsos
        .fields("hasilinformasiprolanisPRB") = hasilinformasiprolanisPRB
        .fields("hasilinformasinosktm") = hasilinformasinosktm
        .fields("hasilpoli") = hasilpoli
        .fields("hasilpolieksekutif") = hasilpolieksekutif
        .fields("hasiltglsep") = hasiltglsep
        .Update
    End With
End Function

