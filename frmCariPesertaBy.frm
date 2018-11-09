VERSION 5.00
Begin VB.Form frmCariPesertaBy 
   Caption         =   "Form2"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form2"
   ScaleHeight     =   1605
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCari 
      Appearance      =   0  'Flat
      Caption         =   "CARI"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtParam 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.OptionButton NoRujukan 
      Caption         =   "NoRujukan"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton OptNIK 
      Caption         =   "NIK"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton OptNoKartu 
      Caption         =   "NoKartu"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCariPesertaBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
