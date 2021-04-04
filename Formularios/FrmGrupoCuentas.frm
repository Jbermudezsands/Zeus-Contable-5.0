VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmGrupoCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Grupo de Cuentas"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "&Siguiente"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
      Begin MSDBCtls.DBCombo DBCodGrupo 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmGrupoCuentas.frx":0000
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "FrmGrupoCuentas.frx":0078
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmGrupoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
