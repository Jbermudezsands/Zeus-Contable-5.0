VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmInforme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de usuarios"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "FrmInforme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Usuario Actual"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmInforme.frx":08CA
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmInforme.frx":0948
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblUsuario 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmInforme.frx":09C4
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNivel 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmInforme.frx":0A22
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   3840
      Picture         =   "FrmInforme.frx":0A80
      ScaleHeight     =   705
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
LblNivel.Caption = NivelAcceso
LblUsuario.Caption = NombreUsuario
Me.Picture1.BackColor = RGB(207, 207, 207)
End Sub

Private Sub CmdAceptar_Click()
Unload Me
End Sub



Private Sub xptopbuttons1_Click()
Unload Me
End Sub


