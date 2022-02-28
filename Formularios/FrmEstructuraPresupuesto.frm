VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmEstructuraPresupuesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estructura Presupuestaria"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   15645
   Begin MSAdodcLib.Adodc AdoPresupuestoAnual 
      Height          =   375
      Left            =   480
      Top             =   5880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoPresupuestoAnual"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   14280
      TabIndex        =   200
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Distribuir"
      Height          =   375
      Left            =   12720
      TabIndex        =   197
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton CmdCopiar 
      Caption         =   "Copiar"
      Height          =   375
      Left            =   11160
      TabIndex        =   196
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox TxtTotal1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   10680
      TabIndex        =   120
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox TxtTotal2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   12360
      TabIndex        =   119
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox TxtTotal3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   14040
      TabIndex        =   118
      Top             =   3840
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel Lbl1 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":0000
      TabIndex        =   112
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   98
      Text            =   "0.00"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   97
      Text            =   "0.00"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   96
      Text            =   "0.00"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   95
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   94
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   93
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   92
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   91
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   90
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   89
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   88
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10920
      TabIndex        =   87
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   86
      Text            =   "0.00"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   85
      Text            =   "0.00"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   84
      Text            =   "0.00"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   83
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   82
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   81
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   80
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   79
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text21 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   78
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   77
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   76
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text24 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12600
      TabIndex        =   75
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text25 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   74
      Text            =   "0.00"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text26 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   73
      Text            =   "0.00"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text27 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   72
      Text            =   "0.00"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text28 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   71
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text29 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   70
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text30 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   69
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text31 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   68
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text32 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   67
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text33 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   66
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text34 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   65
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text35 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   64
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text36 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   14280
      TabIndex        =   63
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "0.00"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Txt12 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt11 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt10 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt9 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt7 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "0.00"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Txt2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "0.00"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Txt36 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt35 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt34 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt33 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt32 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt31 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt30 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt29 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt28 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt27 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0.00"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Txt26 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Txt25 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Txt24 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt23 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt22 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt21 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt20 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt19 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt18 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt17 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt16 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt15 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0.00"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Txt14 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Txt13 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   840
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":006E
      TabIndex        =   14
      Top             =   240
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas de Mayor"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   5415
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   4320
         TabIndex        =   198
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton SmartButton1 
         Caption         =   "Editar"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton SmartButton2 
         Caption         =   "&Nuevo"
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuentas"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdBorrarCuentas 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdVerCuenta 
         Caption         =   "Ver Cuenta"
         Height          =   495
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdMover 
         Caption         =   "Mover"
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mover Grupo Ctas"
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin SmartButtonProject.SmartButton CmdCancelar 
         Height          =   735
         Left            =   3720
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         Caption         =   "Cancelar"
         Picture         =   "FrmEstructuraPresupuesto.frx":00E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGridCuentas 
      Bindings        =   "FrmEstructuraPresupuesto.frx":0DBA
      Height          =   1695
      Left            =   2040
      TabIndex        =   0
      Top             =   6720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2990
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Detalle de Cuentas"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureCurrentRow(0)=   "bHQAAOYBAABCTeYBAAAAAAAANgAAACgAAAAPAAAACQAAAAEAGAAAAAAAsAEAAAAAAAAAAAAAAAAA"
      PictureCurrentRow(1)=   "AAAAAAD///////////////////////////////////////////////////////////8AAAD/////"
      PictureCurrentRow(2)=   "//////////////////////////////////////////////////////8AAAD///////8AhgAAhgAA"
      PictureCurrentRow(3)=   "hgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgD///////8AAAD///////8AhgD///+EhoSEhoSEhoSE"
      PictureCurrentRow(4)=   "hoSEhoSEhoSEhoSEhoQAhgD///////8AAAD///////8AhgD////Gx8bGx8bGx8bGx8bGx8bGx8bG"
      PictureCurrentRow(5)=   "x8aEhoQAhgD///////8AAAD///////8AhgD///////////////////////////////////8AhgD/"
      PictureCurrentRow(6)=   "//////8AAAD///////8AhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgD///////8AAAD/"
      PictureCurrentRow(7)=   "//////////////////////////////////////////////////////////8AAAD/////////////"
      PictureCurrentRow(8)=   "//////////////////////////////////////////////8AAAA="
      PictureCurrentRow.vt=   9
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000009&"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&H80000009&"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.alignment=2,.bgcolor=&HC08080&"
      _StyleDefs(20)  =   ":id=22,.fgcolor=&H0&,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(21)  =   ":id=22,.strikethrough=0,.charset=0"
      _StyleDefs(22)  =   ":id=22,.fontname=Viner Hand ITC"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.fgcolor=&H800000&,.bold=-1"
      _StyleDefs(24)  =   ":id=14,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=14,.fontname=Garamond"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(43)  =   "Named:id=33:Normal"
      _StyleDefs(44)  =   ":id=33,.parent=0"
      _StyleDefs(45)  =   "Named:id=34:Heading"
      _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   ":id=34,.wraptext=-1"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=37:Caption"
      _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(54)  =   "Named:id=38:HighlightRow"
      _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(56)  =   "Named:id=39:EvenRow"
      _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(58)  =   "Named:id=40:OddRow"
      _StyleDefs(59)  =   ":id=40,.parent=33"
      _StyleDefs(60)  =   "Named:id=41:RecordSelector"
      _StyleDefs(61)  =   ":id=41,.parent=34"
      _StyleDefs(62)  =   "Named:id=42:FilterBar"
      _StyleDefs(63)  =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5640
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstructuraPresupuesto.frx":0DD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstructuraPresupuesto.frx":1225
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstructuraPresupuesto.frx":1677
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstructuraPresupuesto.frx":1AC9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13320
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstructuraPresupuesto.frx":1F1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstructuraPresupuesto.frx":32A5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7858
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList3"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   375
      Left            =   8640
      Top             =   7080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaCuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   8640
      Top             =   7440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaConsulta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaSaldoCuenta 
      Height          =   375
      Left            =   8640
      Top             =   7800
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaSaldoCuenta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaGrupos 
      Height          =   375
      Left            =   8640
      Top             =   8160
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Grupos"
      Caption         =   "DtaGrupos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   10800
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":4E4F
      TabIndex        =   111
      Top             =   240
      Width           =   4815
   End
   Begin ACTIVESKINLibCtl.SkinLabel Lbl2 
      Height          =   255
      Left            =   7560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":4EC3
      TabIndex        =   113
      Top             =   600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel Lbl3 
      Height          =   255
      Left            =   9120
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":4F31
      TabIndex        =   114
      Top             =   600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel Lbl4 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":4F9F
      TabIndex        =   115
      Top             =   600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel Lbl5 
      Height          =   255
      Left            =   12600
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":500D
      TabIndex        =   116
      Top             =   600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel Lbl6 
      Height          =   255
      Left            =   14280
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":507B
      TabIndex        =   117
      Top             =   600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblTotal1 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":50E9
      TabIndex        =   121
      Top             =   3840
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblTotal2 
      Height          =   255
      Left            =   7560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":514F
      TabIndex        =   122
      Top             =   3840
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblTotal3 
      Height          =   255
      Left            =   9120
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":51B5
      TabIndex        =   123
      Top             =   3840
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":521B
      TabIndex        =   124
      Top             =   840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":527B
      TabIndex        =   125
      Top             =   1080
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":52DB
      TabIndex        =   126
      Top             =   1320
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":533B
      TabIndex        =   127
      Top             =   1560
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":539B
      TabIndex        =   128
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":53FB
      TabIndex        =   129
      Top             =   2040
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":545B
      TabIndex        =   130
      Top             =   2280
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":54BB
      TabIndex        =   131
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":551B
      TabIndex        =   132
      Top             =   2760
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":557B
      TabIndex        =   133
      Top             =   3000
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":55DD
      TabIndex        =   134
      Top             =   3240
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   5640
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":563F
      TabIndex        =   135
      Top             =   3480
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":56A1
      TabIndex        =   136
      Top             =   840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5701
      TabIndex        =   137
      Top             =   1080
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5761
      TabIndex        =   138
      Top             =   1320
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":57C1
      TabIndex        =   139
      Top             =   1560
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5821
      TabIndex        =   140
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5881
      TabIndex        =   141
      Top             =   2040
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":58E1
      TabIndex        =   142
      Top             =   2280
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5941
      TabIndex        =   143
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":59A1
      TabIndex        =   144
      Top             =   2760
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5A01
      TabIndex        =   145
      Top             =   3000
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5A63
      TabIndex        =   146
      Top             =   3240
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
      Height          =   255
      Left            =   7200
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5AC5
      TabIndex        =   147
      Top             =   3480
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5B27
      TabIndex        =   148
      Top             =   840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5B87
      TabIndex        =   149
      Top             =   1080
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5BE7
      TabIndex        =   150
      Top             =   1320
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5C47
      TabIndex        =   151
      Top             =   1560
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5CA7
      TabIndex        =   152
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5D07
      TabIndex        =   153
      Top             =   2040
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5D67
      TabIndex        =   154
      Top             =   2280
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5DC7
      TabIndex        =   155
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5E27
      TabIndex        =   156
      Top             =   2760
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5E87
      TabIndex        =   157
      Top             =   3000
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5EE9
      TabIndex        =   158
      Top             =   3240
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5F4B
      TabIndex        =   159
      Top             =   3480
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":5FAD
      TabIndex        =   160
      Top             =   840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":600D
      TabIndex        =   161
      Top             =   1080
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":606D
      TabIndex        =   162
      Top             =   1320
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":60CD
      TabIndex        =   163
      Top             =   1560
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":612D
      TabIndex        =   164
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":618D
      TabIndex        =   165
      Top             =   2040
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":61ED
      TabIndex        =   166
      Top             =   2280
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":624D
      TabIndex        =   167
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":62AD
      TabIndex        =   168
      Top             =   2760
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel48 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":630D
      TabIndex        =   169
      Top             =   3000
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel49 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":636F
      TabIndex        =   170
      Top             =   3240
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel50 
      Height          =   255
      Left            =   10560
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":63D1
      TabIndex        =   171
      Top             =   3480
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6433
      TabIndex        =   172
      Top             =   840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel52 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6493
      TabIndex        =   173
      Top             =   1080
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel53 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":64F3
      TabIndex        =   174
      Top             =   1320
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel54 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6553
      TabIndex        =   175
      Top             =   1560
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel55 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":65B3
      TabIndex        =   176
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel56 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6613
      TabIndex        =   177
      Top             =   2040
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel57 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6673
      TabIndex        =   178
      Top             =   2280
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel58 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":66D3
      TabIndex        =   179
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel59 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6733
      TabIndex        =   180
      Top             =   2760
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel60 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6793
      TabIndex        =   181
      Top             =   3000
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel61 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":67F5
      TabIndex        =   182
      Top             =   3240
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel62 
      Height          =   255
      Left            =   12240
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6857
      TabIndex        =   183
      Top             =   3480
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel63 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":68B9
      TabIndex        =   184
      Top             =   840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel64 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6919
      TabIndex        =   185
      Top             =   1080
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel65 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6979
      TabIndex        =   186
      Top             =   1320
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel66 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":69D9
      TabIndex        =   187
      Top             =   1560
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel67 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6A39
      TabIndex        =   188
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel68 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6A99
      TabIndex        =   189
      Top             =   2040
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel69 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6AF9
      TabIndex        =   190
      Top             =   2280
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel70 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6B59
      TabIndex        =   191
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel71 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6BB9
      TabIndex        =   192
      Top             =   2760
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel72 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6C19
      TabIndex        =   193
      Top             =   3000
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel73 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6C7B
      TabIndex        =   194
      Top             =   3240
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel74 
      Height          =   255
      Left            =   13920
      OleObjectBlob   =   "FrmEstructuraPresupuesto.frx":6CDD
      TabIndex        =   195
      Top             =   3480
      Width           =   375
   End
   Begin MSAdodcLib.Adodc DtaPeriodos 
      Height          =   330
      Left            =   3480
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaPeriodos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3480
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaCuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6000
      Top             =   6000
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaConsulta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaPresupuesto 
      Height          =   330
      Left            =   6000
      Top             =   6360
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaPresupuesto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   330
      Left            =   8640
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaNacceso"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaPresupuestoAnual 
      Height          =   330
      Left            =   8640
      Top             =   6000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaPresupuestoAnual"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   5760
      TabIndex        =   199
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
      _Version        =   786432
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoPresupuesto 
      Height          =   330
      Left            =   4920
      Top             =   7800
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoPresupuesto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   110
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   109
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   108
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   107
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   106
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   105
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   104
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   103
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   102
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   101
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label75 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   100
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label76 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   12600
      TabIndex        =   99
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   62
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   61
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   60
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   59
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   58
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   57
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   56
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   55
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   54
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   53
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   52
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   9120
      TabIndex        =   51
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "FrmEstructuraPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enfoque As Double, TotalAño1 As Double, TotalAño2 As Double, TotalAño3 As Double

Private Sub CmdBorrar_Click()
   Dim NodX As Node
Dim Respuesta As Integer
Dim Relatives As String, RelationsShips As String
Dim LLave As String, Texto As String, Imagen1 As Integer
Dim Imagen2 As Integer
  If KeyPrincipal = "B" Or KeyPrincipal = "A" Or KeyPrincipal = "C" Or KeyPrincipal = "G" Or KeyPrincipal = "D" Or KeyPrincipal = "O" Then
      MsgBox "No se Puede Borrar el Grupo Principal", vbCritical, "Sistema Contable"
      Exit Sub
   Else
     Me.DtaConsulta.RecordSource = "SELECT EstructuraPresupuesto.KeyGrupo, EstructuraPresupuesto.KeyGrupoSuperior, EstructuraPresupuesto.Child, EstructuraPresupuesto.DescripcionGrupo, EstructuraPresupuesto.Imagen1, EstructuraPresupuesto.Imagen2 From EstructuraPresupuesto Where (((EstructuraPresupuesto.KeyGrupoSuperior) = '" & KeyPrincipal & "'))ORDER BY EstructuraPresupuesto.DescripcionGrupo"
     Me.DtaConsulta.Refresh
        If Not Me.DtaConsulta.Recordset.EOF Then
          MsgBox "Este Grupo tiene SubGrupos, no se puede Borrar", vbCritical, "Sistema Contable"
          Exit Sub
        Else
         Me.DtaConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & KeyPrincipal & "')) ORDER BY Cuentas.CodCuentas"
         Me.DtaConsulta.Refresh
          If Not DtaConsulta.Recordset.EOF Then
           '   MsgBox "Este Grupo tiene Cuentas, no se puede Borrar", vbCritical, "Sistema Contable"
           '   Exit Sub
             Else
               Respuesta = MsgBox("Esta seguro de Borrar este Grupo", vbYesNo, "Sistema Contable")
               If Respuesta = 6 Then
           
                   Me.DtaConsulta.RecordSource = "SELECT EstructuraPresupuesto.KeyGrupo, EstructuraPresupuesto.DescripcionGrupo From EstructuraPresupuesto Where (((EstructuraPresupuesto.KeyGrupo) = '" & KeyPrincipal & "'))"
                   Me.DtaConsulta.Refresh
                   If Not DtaConsulta.Recordset.EOF Then
                     Me.DtaConsulta.Recordset.Delete
                   End If
                   Me.TreeView1.Nodes.Remove (KeyPrincipal)
                   
              End If
          End If
      End If
 End If
End Sub

Private Sub CmdCopiar_Click()
 Dim i As Double, Cc As Double, Monto As Double, c As String
 

Monto = MontoPresupuesto(Enfoque)

c = LlenarTexto(Monto, Enfoque)

End Sub
Public Function TotalAño(iPosicion As Double) As Double
  Dim Total As Double
  
   
   If iPosicion <= 12 Then
     Total = CDbl(Me.Text1.Text) + CDbl(Me.Text2.Text) + CDbl(Me.Text3.Text) + CDbl(Me.Text4.Text) + CDbl(Me.Text5.Text) + CDbl(Me.Text6.Text) + CDbl(Me.Text7.Text) + CDbl(Me.Text8.Text) + CDbl(Me.Text9.Text) + CDbl(Me.Text10.Text) + CDbl(Me.Text11.Text) + CDbl(Me.Text12.Text)
   ElseIf iPosicion > 12 And iPosicion <= 24 Then
     Total = CDbl(Me.Text13.Text) + CDbl(Me.Text14.Text) + CDbl(Me.Text15.Text) + CDbl(Me.Text16.Text) + CDbl(Me.Text17.Text) + CDbl(Me.Text18.Text) + CDbl(Me.Text19.Text) + CDbl(Me.Text20.Text) + CDbl(Me.Text21.Text) + CDbl(Me.Text22.Text) + CDbl(Me.Text23.Text) + CDbl(Me.Text24.Text)
   ElseIf iPosicion > 24 Then
     Total = CDbl(Me.Text25.Text) + CDbl(Me.Text26.Text) + CDbl(Me.Text27.Text) + CDbl(Me.Text28.Text) + CDbl(Me.Text29.Text) + CDbl(Me.Text30.Text) + CDbl(Me.Text31.Text) + CDbl(Me.Text32.Text) + CDbl(Me.Text33.Text) + CDbl(Me.Text34.Text) + CDbl(Me.Text35.Text) + CDbl(Me.Text36.Text)
   End If
   
   TotalAño = Total

End Function



Private Function LlenarTexto(Monto As Double, iPosicion As Double) As Double
  Select Case iPosicion
     Case 1: Me.Text1.Text = Monto: Me.Text2.Text = Monto: Me.Text3.Text = Monto: Me.Text4.Text = Monto: Me.Text5.Text = Monto: Me.Text6.Text = Monto: Me.Text7.Text = Monto: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 2: Me.Text2.Text = Monto: Me.Text3.Text = Monto: Me.Text4.Text = Monto: Me.Text5.Text = Monto: Me.Text6.Text = Monto: Me.Text7.Text = Monto: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 3: Me.Text3.Text = Monto: Me.Text4.Text = Monto: Me.Text5.Text = Monto: Me.Text6.Text = Monto: Me.Text7.Text = Monto: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 4: Me.Text4.Text = Monto: Me.Text5.Text = Monto: Me.Text6.Text = Monto: Me.Text7.Text = Monto: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 5: Me.Text5.Text = Monto: Me.Text6.Text = Monto: Me.Text7.Text = Monto: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 6: Me.Text6.Text = Monto: Me.Text7.Text = Monto: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 7: Me.Text7.Text = Monto: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 8: Me.Text8.Text = Monto: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 9: Me.Text9.Text = Monto: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 10: Me.Text10.Text = Monto: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 11: Me.Text11.Text = Monto: Me.Text12.Text = Monto
     Case 12: Me.Text12.Text = Monto
     
     Case 13: Me.Text13.Text = Monto: Me.Text14.Text = Monto: Me.Text15.Text = Monto: Me.Text16.Text = Monto: Me.Text17.Text = Monto: Me.Text18.Text = Monto: Me.Text19.Text = Monto: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 14: Me.Text14.Text = Monto: Me.Text15.Text = Monto: Me.Text16.Text = Monto: Me.Text17.Text = Monto: Me.Text18.Text = Monto: Me.Text19.Text = Monto: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 15: Me.Text15.Text = Monto: Me.Text16.Text = Monto: Me.Text17.Text = Monto: Me.Text18.Text = Monto: Me.Text19.Text = Monto: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 16: Me.Text16.Text = Monto: Me.Text17.Text = Monto: Me.Text18.Text = Monto: Me.Text19.Text = Monto: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 17: Me.Text17.Text = Monto: Me.Text18.Text = Monto: Me.Text19.Text = Monto: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 18: Me.Text18.Text = Monto: Me.Text19.Text = Monto: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 19: Me.Text19.Text = Monto: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 20: Me.Text20.Text = Monto: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 21: Me.Text21.Text = Monto: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 22: Me.Text22.Text = Monto: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 23: Me.Text23.Text = Monto: Me.Text24.Text = Monto
     Case 24: Me.Text24.Text = Monto

     Case 25: Me.Text25.Text = Monto: Me.Text26.Text = Monto: Me.Text27.Text = Monto: Me.Text28.Text = Monto: Me.Text29.Text = Monto: Me.Text30.Text = Monto: Me.Text31.Text = Monto: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 26: Me.Text26.Text = Monto: Me.Text27.Text = Monto: Me.Text28.Text = Monto: Me.Text29.Text = Monto: Me.Text30.Text = Monto: Me.Text31.Text = Monto: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 27: Me.Text27.Text = Monto: Me.Text28.Text = Monto: Me.Text29.Text = Monto: Me.Text30.Text = Monto: Me.Text31.Text = Monto: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 28: Me.Text28.Text = Monto: Me.Text29.Text = Monto: Me.Text30.Text = Monto: Me.Text31.Text = Monto: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 29: Me.Text29.Text = Monto: Me.Text30.Text = Monto: Me.Text31.Text = Monto: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 30: Me.Text30.Text = Monto: Me.Text31.Text = Monto: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 31: Me.Text31.Text = Monto: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 32: Me.Text32.Text = Monto: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 33: Me.Text33.Text = Monto: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 34: Me.Text34.Text = Monto: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 35: Me.Text35.Text = Monto: Me.Text36.Text = Monto
     Case 36: Me.Text36.Text = Monto

End Select

LlenarTexto = 1
End Function



Public Function MontoPresupuesto(iPosicion As Double) As Double
  Select Case iPosicion
    Case 1: MontoPresupuesto = Me.Text1.Text
    Case 2: MontoPresupuesto = Me.Text2.Text
    Case 3: MontoPresupuesto = Me.Text3.Text
    Case 4: MontoPresupuesto = Me.Text4.Text
    Case 5: MontoPresupuesto = Me.Text5.Text
    Case 6: MontoPresupuesto = Me.Text6.Text
    Case 7: MontoPresupuesto = Me.Text7.Text
    Case 8: MontoPresupuesto = Me.Text8.Text
    Case 9: MontoPresupuesto = Me.Text9.Text
    Case 10: MontoPresupuesto = Me.Text10.Text
    Case 11: MontoPresupuesto = Me.Text11.Text
    Case 12: MontoPresupuesto = Me.Text12.Text
    
    Case 13: MontoPresupuesto = Me.Text13.Text
    Case 13: MontoPresupuesto = Me.Text14.Text
    Case 15: MontoPresupuesto = Me.Text15.Text
    Case 16: MontoPresupuesto = Me.Text16.Text
    Case 17: MontoPresupuesto = Me.Text17.Text
    Case 18: MontoPresupuesto = Me.Text18.Text
    Case 19: MontoPresupuesto = Me.Text19.Text
    Case 20: MontoPresupuesto = Me.Text20.Text
    Case 21: MontoPresupuesto = Me.Text21.Text
    Case 22: MontoPresupuesto = Me.Text22.Text
    Case 23: MontoPresupuesto = Me.Text23.Text
    Case 24: MontoPresupuesto = Me.Text24.Text
    
    Case 25: MontoPresupuesto = Me.Text25.Text
    Case 26: MontoPresupuesto = Me.Text26.Text
    Case 27: MontoPresupuesto = Me.Text27.Text
    Case 28: MontoPresupuesto = Me.Text28.Text
    Case 29: MontoPresupuesto = Me.Text29.Text
    Case 30: MontoPresupuesto = Me.Text30.Text
    Case 31: MontoPresupuesto = Me.Text31.Text
    Case 32: MontoPresupuesto = Me.Text32.Text
    Case 33: MontoPresupuesto = Me.Text33.Text
    Case 34: MontoPresupuesto = Me.Text34.Text
    Case 35: MontoPresupuesto = Me.Text35.Text
    Case 36: MontoPresupuesto = Me.Text36.Text
  End Select
  

End Function


Private Sub CmdGrabar_Click()
Dim NumeroPeriodo As Integer, Periodo As Integer
Dim Saldo As Double
 
 '//////////////////////////////////////////////////////////////
 '//////////////Datos del primer  periodo////////////////////////
 '//////////////////////////////////////////////////////////////
 

 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 1))"
 Me.DtaPeriodos.Refresh
 
 Me.DtaPeriodos.Recordset.MoveLast
 Me.osProgress1.Min = 0
 Me.osProgress1.Max = Me.DtaPeriodos.Recordset.RecordCount
 Me.osProgress1.Value = 0
 Me.DtaPeriodos.Refresh
 
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset("NPeriodo")
  Periodo = Me.DtaPeriodos.Recordset("Periodo")
  'CodigoCuenta = Me.DBCliente.Text
  
  
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & KeyPrincipal & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.DtaPresupuesto.Recordset.AddNew
      Me.DtaPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.DtaPresupuesto.Recordset!CodCuenta = KeyPrincipal
      Select Case Periodo
             Case 1: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text1.Text)
             Case 2: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text2.Text)
             Case 3: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text3.Text)
             Case 4: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text4.Text)
             Case 5: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text5.Text)
             Case 6: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text6.Text)
             Case 7: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text7.Text)
             Case 8: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text8.Text)
             Case 9: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text9.Text)
             Case 10: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text10.Text)
             Case 11: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text11.Text)
             Case 12: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text12.Text)
          End Select
      
     Me.DtaPresupuesto.Recordset.Update
   Else '//////En caso que existan datos del presupuesto solo edito el monto
    'Me.DtaConsulta.Recordset.Edit
      
      Select Case Periodo
             Case 1: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text1.Text)
             Case 2: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text2.Text)
             Case 3: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text3.Text)
             Case 4: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text4.Text)
             Case 5: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text5.Text)
             Case 6: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text6.Text)
             Case 7: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text7.Text)
             Case 8: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text8.Text)
             Case 9: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text9.Text)
             Case 10: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text10.Text)
             Case 11: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text11.Text)
             Case 12: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text12.Text)
          End Select
      
    Me.DtaConsulta.Recordset.Update
   
   End If
   
   Me.osProgress1.Value = Me.osProgress1.Value + 1
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
 
  '//////////////////////////////////////////////////////////////
 '//////////////Datos del Segundo  periodo presupuestado////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 2))"
 Me.DtaPeriodos.Refresh
 
 Me.DtaPeriodos.Recordset.MoveLast
 Me.osProgress1.Min = 0
 Me.osProgress1.Max = Me.DtaPeriodos.Recordset.RecordCount
 Me.osProgress1.Value = 0
 Me.DtaPeriodos.Refresh
 
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
'  CodigoCuenta = Me.DBCliente.Text
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & KeyPrincipal & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.DtaPresupuesto.Recordset.AddNew
      Me.DtaPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.DtaPresupuesto.Recordset!CodCuenta = KeyPrincipal
      Select Case Periodo
             Case 1: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text13.Text)
             Case 2: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text14.Text)
             Case 3: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text15.Text)
             Case 4: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text16.Text)
             Case 5: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text17.Text)
             Case 6: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text18.Text)
             Case 7: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text18.Text)
             Case 8: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text20.Text)
             Case 9: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text21.Text)
             Case 10: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text22.Text)
             Case 11: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text23.Text)
             Case 12: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text24.Text)
          End Select
      
     Me.DtaPresupuesto.Recordset.Update
   Else '//////En caso que existan datos del presupuesto solo edito el monto
    'Me.DtaConsulta.Recordset.Edit
      
      Select Case Periodo
             Case 1: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text13.Text)
             Case 2: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text14.Text)
             Case 3: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text15.Text)
             Case 4: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text16.Text)
             Case 5: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text17.Text)
             Case 6: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text18.Text)
             Case 7: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text19.Text)
             Case 8: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text20.Text)
             Case 9: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text21.Text)
             Case 10: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text22.Text)
             Case 11: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text23.Text)
             Case 12: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text24.Text)
          End Select
      
    Me.DtaConsulta.Recordset.Update
   
   End If
     Me.osProgress1.Value = Me.osProgress1.Value + 1
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
   '//////////////////////////////////////////////////////////////
 '//////////////Datos del tercer  periodo presupuestado////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 3))"
 Me.DtaPeriodos.Refresh
 
 Me.DtaPeriodos.Recordset.MoveLast
 Me.osProgress1.Min = 0
 Me.osProgress1.Max = Me.DtaPeriodos.Recordset.RecordCount
 Me.osProgress1.Value = 0
 Me.DtaPeriodos.Refresh
 
 
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
'  CodigoCuenta = Me.DBCliente.Text
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & KeyPrincipal & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.DtaPresupuesto.Recordset.AddNew
      Me.DtaPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.DtaPresupuesto.Recordset!CodCuenta = KeyPrincipal
      Select Case Periodo
             Case 1: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text25.Text)
             Case 2: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text26.Text)
             Case 3: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text27.Text)
             Case 4: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text28.Text)
             Case 5: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text29.Text)
             Case 6: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text30.Text)
             Case 7: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text31.Text)
             Case 8: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text32.Text)
             Case 9: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text33.Text)
             Case 10: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text34.Text)
             Case 11: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text35.Text)
             Case 12: Me.DtaPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text36.Text)
          End Select
      
     Me.DtaPresupuesto.Recordset.Update
   Else '//////En caso que existan datos del presupuesto solo edito el monto
    'Me.DtaConsulta.Recordset.Edit
      
      Select Case Periodo
             Case 1: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text25.Text)
             Case 2: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text26.Text)
             Case 3: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text27.Text)
             Case 4: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text28.Text)
             Case 5: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text29.Text)
             Case 6: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text30.Text)
             Case 7: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text31.Text)
             Case 8: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text32.Text)
             Case 9: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text33.Text)
             Case 10: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text34.Text)
             Case 11: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text35.Text)
             Case 12: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text36.Text)
          End Select
      
    Me.DtaConsulta.Recordset.Update
   
   End If
   
  Me.osProgress1.Value = Me.osProgress1.Value + 1
       
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
 
End Sub

Private Sub Command2_Click()
 Dim Monto As Double, cadena As String
 
cadena = InputBox("Digito el Monto a Distribuir", "Zeus Contable")
 
 
 If Not IsNumeric(cadena) Then
   MsgBox "Lo digitado no es numerico", vbCritical, "Zeus Contable"
   Exit Sub
 End If
 
 Monto = cadena
 
 Monto = Distribuir(Monto, Enfoque)
 


End Sub
Public Function Distribuir(Monto As Double, iPosicion As Double) As Double
 Dim i As Double, MontoDist As Double, MontoTotal As Double, Diferencia As Double
 
 MontoDist = Format(Monto / 12, "##,##0.00")
 
  If Monto > 0 Then
        For i = 1 To 12
         
         If iPosicion <= 12 Then
           Me.Text1.Text = MontoDist: Me.Text2.Text = MontoDist: Me.Text3.Text = MontoDist: Me.Text4.Text = MontoDist: Me.Text5.Text = MontoDist: Me.Text6.Text = MontoDist: Me.Text7.Text = MontoDist: Me.Text8.Text = MontoDist: Me.Text9.Text = MontoDist: Me.Text10.Text = MontoDist: Me.Text11.Text = MontoDist: Me.Text12.Text = MontoDist
        ElseIf iPosicion > 12 And iPosicion <= 24 Then
          Me.Text13.Text = MontoDist: Me.Text14.Text = MontoDist: Me.Text15.Text = MontoDist: Me.Text16.Text = MontoDist: Me.Text17.Text = MontoDist: Me.Text18.Text = MontoDist: Me.Text19.Text = MontoDist: Me.Text20.Text = MontoDist: Me.Text21.Text = MontoDist: Me.Text22.Text = MontoDist: Me.Text23.Text = MontoDist: Me.Text24.Text = MontoDist
        ElseIf iPosicion > 24 Then
          Me.Text25.Text = MontoDist: Me.Text26.Text = MontoDist: Me.Text27.Text = MontoDist: Me.Text28.Text = MontoDist: Me.Text29.Text = MontoDist: Me.Text30.Text = MontoDist: Me.Text31.Text = MontoDist: Me.Text32.Text = MontoDist: Me.Text33.Text = MontoDist: Me.Text34.Text = MontoDist: Me.Text35.Text = MontoDist: Me.Text36.Text = MontoDist
        End If
         
        Next
  End If
  
   If iPosicion <= 12 Then
     MontoTotal = CDbl(Me.Text1.Text) + CDbl(Me.Text2.Text) + CDbl(Me.Text3.Text) + CDbl(Me.Text4.Text) + CDbl(Me.Text5.Text) + CDbl(Me.Text6.Text) + CDbl(Me.Text7.Text) + CDbl(Me.Text8.Text) + CDbl(Me.Text9.Text) + CDbl(Me.Text10.Text) + CDbl(Me.Text11.Text) + CDbl(Me.Text12.Text)
     Diferencia = Monto - MontoTotal
     Me.Text12.Text = Format(CDbl(Me.Text12.Text) + Diferencia, "##,##0.00")
   ElseIf iPosicion > 12 And iPosicion <= 24 Then
     MontoTotal = CDbl(Me.Text13.Text) + CDbl(Me.Text14.Text) + CDbl(Me.Text15.Text) + CDbl(Me.Text16.Text) + CDbl(Me.Text17.Text) + CDbl(Me.Text18.Text) + CDbl(Me.Text19.Text) + CDbl(Me.Text20.Text) + CDbl(Me.Text21.Text) + CDbl(Me.Text22.Text) + CDbl(Me.Text23.Text) + CDbl(Me.Text24.Text)
     Diferencia = Monto - MontoTotal
     Me.Text24.Text = Format(CDbl(Me.Text24.Text) + Diferencia, "##,##0.00")
   ElseIf iPosicion > 24 Then
     MontoTotal = CDbl(Me.Text25.Text) + CDbl(Me.Text26.Text) + CDbl(Me.Text27.Text) + CDbl(Me.Text28.Text) + CDbl(Me.Text29.Text) + CDbl(Me.Text30.Text) + CDbl(Me.Text31.Text) + CDbl(Me.Text32.Text) + CDbl(Me.Text33.Text) + CDbl(Me.Text34.Text) + CDbl(Me.Text35.Text) + CDbl(Me.Text36.Text)
     Diferencia = Monto - MontoTotal
     Me.Text36.Text = Format(CDbl(Me.Text36.Text) + Diferencia, "##,##0.00")
   End If
   
   
   
   Distribuir = 0

End Function


Private Sub Command3_Click()
On Error GoTo TipoErrs
Dim NumeroPeriodo As Integer, Periodo As Integer
Dim Saldo As Double
 
 Me.Command3.Enabled = False
 
 '//////////////////////////////////////////////////////////////
 '//////////////Datos del primer  periodo////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 1))"
 Me.DtaPeriodos.Refresh
 
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.AdoPresupuesto.Recordset.AddNew
      Me.AdoPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.AdoPresupuesto.Recordset!CodCuenta = CodigoCuenta
      Select Case Periodo
             Case 1: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text1.Text)
             Case 2: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text2.Text)
             Case 3: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text3.Text)
             Case 4: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text4.Text)
             Case 5: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text5.Text)
             Case 6: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text6.Text)
             Case 7: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text7.Text)
             Case 8: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text8.Text)
             Case 9: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text9.Text)
             Case 10: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text10.Text)
             Case 11: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text11.Text)
             Case 12: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text12.Text)
          End Select
      
     Me.AdoPresupuesto.Recordset.Update
   Else '//////En caso que existan datos del presupuesto solo edito el monto
    'Me.DtaConsulta.Recordset.Edit
      
      Select Case Periodo
             Case 1: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text1.Text)
             Case 2: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text2.Text)
             Case 3: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text3.Text)
             Case 4: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text4.Text)
             Case 5: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text5.Text)
             Case 6: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text6.Text)
             Case 7: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text7.Text)
             Case 8: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text8.Text)
             Case 9: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text9.Text)
             Case 10: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text10.Text)
             Case 11: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text11.Text)
             Case 12: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text12.Text)
          End Select
      
    Me.DtaConsulta.Recordset.Update
   
   End If
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
 
  '//////////////////////////////////////////////////////////////
 '//////////////Datos del Segundo  periodo presupuestado////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 2))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.AdoPresupuesto.Recordset.AddNew
      Me.AdoPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.AdoPresupuesto.Recordset!CodCuenta = CodigoCuenta
      Select Case Periodo
             Case 1: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text13.Text)
             Case 2: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text14.Text)
             Case 3: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text15.Text)
             Case 4: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text16.Text)
             Case 5: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text17.Text)
             Case 6: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text18.Text)
             Case 7: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text18.Text)
             Case 8: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text20.Text)
             Case 9: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text21.Text)
             Case 10: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text22.Text)
             Case 11: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text23.Text)
             Case 12: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text24.Text)
          End Select
      
     Me.AdoPresupuesto.Recordset.Update
   Else '//////En caso que existan datos del presupuesto solo edito el monto
    'Me.DtaConsulta.Recordset.Edit
      
      Select Case Periodo
             Case 1: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text13.Text)
             Case 2: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text14.Text)
             Case 3: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text15.Text)
             Case 4: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text16.Text)
             Case 5: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text17.Text)
             Case 6: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text18.Text)
             Case 7: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text19.Text)
             Case 8: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text20.Text)
             Case 9: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text21.Text)
             Case 10: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text22.Text)
             Case 11: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text23.Text)
             Case 12: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text24.Text)
          End Select
      
    Me.DtaConsulta.Recordset.Update
   
   End If
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
   '//////////////////////////////////////////////////////////////
 '//////////////Datos del tercer  periodo presupuestado////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 3))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.AdoPresupuesto.Recordset.AddNew
      Me.AdoPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.AdoPresupuesto.Recordset!CodCuenta = CodigoCuenta
      Select Case Periodo
             Case 1: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text25.Text)
             Case 2: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text26.Text)
             Case 3: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text27.Text)
             Case 4: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text28.Text)
             Case 5: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text29.Text)
             Case 6: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text30.Text)
             Case 7: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text31.Text)
             Case 8: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text32.Text)
             Case 9: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text33.Text)
             Case 10: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text34.Text)
             Case 11: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text35.Text)
             Case 12: Me.AdoPresupuesto.Recordset!MontoPresupuestado = CDbl(Me.Text36.Text)
          End Select
      
     Me.AdoPresupuesto.Recordset.Update
   Else '//////En caso que existan datos del presupuesto solo edito el monto
    'Me.DtaConsulta.Recordset.Edit
      
      Select Case Periodo
             Case 1: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text25.Text)
             Case 2: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text26.Text)
             Case 3: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text27.Text)
             Case 4: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text28.Text)
             Case 5: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text29.Text)
             Case 6: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text30.Text)
             Case 7: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text31.Text)
             Case 8: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text32.Text)
             Case 9: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text33.Text)
             Case 10: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text34.Text)
             Case 11: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text35.Text)
             Case 12: Me.DtaConsulta.Recordset!MontoPresupuestado = CDbl(Me.Text36.Text)
          End Select
      
    Me.DtaConsulta.Recordset.Update
   
   End If
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
 
 
 
 
 
 '//////si existen datos grabo el monto anual////////////
 If Not Me.TxtTotal1.Text = "" Then
 
 Me.AdoPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 1) And ((PresupuestoAnual.CodigoCuenta) = '" & KeyPrincipal & "'))"
 Me.AdoPresupuestoAnual.Refresh
 If AdoPresupuestoAnual.Recordset.EOF Then
  AdoPresupuestoAnual.Recordset.AddNew
    AdoPresupuestoAnual.Recordset!NumeroTabla = 1
    AdoPresupuestoAnual.Recordset!CodigoCuenta = KeyPrincipal
    AdoPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal1.Text
  AdoPresupuestoAnual.Recordset.Update
 Else
  'AdoPresupuestoAnual.Recordset.Edit
    AdoPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal1.Text
  AdoPresupuestoAnual.Recordset.Update
 End If
 End If
 
  If Not Me.TxtTotal2.Text = "" Then
 Me.AdoPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 2) And ((PresupuestoAnual.CodigoCuenta) = '" & KeyPrincipal & "'))"
 Me.AdoPresupuestoAnual.Refresh
 If AdoPresupuestoAnual.Recordset.EOF Then
  AdoPresupuestoAnual.Recordset.AddNew
    AdoPresupuestoAnual.Recordset!NumeroTabla = 2
    AdoPresupuestoAnual.Recordset!CodigoCuenta = KeyPrincipal
    AdoPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal2.Text
  AdoPresupuestoAnual.Recordset.Update
 Else
  'AdoPresupuestoAnual.Recordset.Edit
    AdoPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal2.Text
  AdoPresupuestoAnual.Recordset.Update
 End If
 End If
 
 If Not Me.TxtTotal3.Text = "" Then
 Me.AdoPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 3) And ((PresupuestoAnual.CodigoCuenta) = '" & KeyPrincipal & "'))"
 Me.AdoPresupuestoAnual.Refresh
 If AdoPresupuestoAnual.Recordset.EOF Then
  AdoPresupuestoAnual.Recordset.AddNew
    AdoPresupuestoAnual.Recordset!NumeroTabla = 3
    AdoPresupuestoAnual.Recordset!CodigoCuenta = KeyPrincipal
    AdoPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal3.Text
  AdoPresupuestoAnual.Recordset.Update
 Else
  'AdoPresupuestoAnual.Recordset.Edit
    AdoPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal3.Text
  AdoPresupuestoAnual.Recordset.Update
 End If
 End If
 
 
 MsgBox "Presupuesto Grabado con Exito!!", vbExclamation, "Sistema Contable"
 Me.Command3.Enabled = True
 
 Exit Sub
TipoErrs:
 MsgBox err.Description
 Exit Sub
End Sub


Public Sub Cargar_Presupuesto()

On Error GoTo TipoErrs
Dim NumeroPeriodo As Integer, Periodo As Integer, TipoCuenta As String
Dim Saldo As Double, Total1 As Double, Total2 As Double, Total3 As Double, Total4 As Double, Total5 As Double, Total6 As Double
Dim Monto1 As Double, Monto2 As Double, Monto3 As Double, Monto4 As Double, Monto5 As Double, Monto6 As Double
Dim FechaInicio As Date, FechaFin As Date

'Criterio = "KeyGrupo='" & KeyPrincipal & "'"
'If DtaCuentas.Recordset.RecordCount <> 0 Then DtaCuentas.Recordset.MoveFirst
'Me.DtaCuentas.Recordset.Find (Criterio)
Me.DtaCuentas.RecordSource = "SELECT KeyGrupo, CodGrupo, KeyGrupoSuperior, Child, DescripcionGrupo, Imagen1, Imagen2 From EstructuraPresupuesto WHERE (EstructuraPresupuesto.KeyGrupo='" & KeyPrincipal & "')"
Me.DtaCuentas.Refresh
If DtaCuentas.Recordset.EOF Then

       

'  Me.TxtDescripcion.Text = ""
  Me.LblTotal1.Caption = "0.00"
  Me.LblTotal2.Caption = "0.00"
  Me.LblTotal3.Caption = "0.00"
'  Me.LblTotal4.Caption = "0.00"
'  Me.LblTotal5.Caption = "0.00"
'  Me.LblTotal6.Caption = "0.00"
  Total1 = 0
  Total2 = 0
  Total3 = 0
  Total4 = 0
  Total5 = 0
  Total6 = 0
  
  Saldo = 0
      Me.Txt1.Text = Format(Saldo, "##,##0.00")
      Me.Txt2.Text = Format(Saldo, "##,##0.00")
      Me.Txt3.Text = Format(Saldo, "##,##0.00")
      Me.Txt4.Text = Format(Saldo, "##,##0.00")
      Me.Txt5.Text = Format(Saldo, "##,##0.00")
      Me.Txt6.Text = Format(Saldo, "##,##0.00")
      Me.Txt7.Text = Format(Saldo, "##,##0.00")
      Me.Txt8.Text = Format(Saldo, "##,##0.00")
      Me.Txt9.Text = Format(Saldo, "##,##0.00")
      Me.Txt10.Text = Format(Saldo, "##,##0.00")
      Me.Txt11.Text = Format(Saldo, "##,##0.00")
      Me.Txt12.Text = Format(Saldo, "##,##0.00")
      
      Me.Txt13.Text = Format(Saldo, "##,##0.00")
      Me.Txt14.Text = Format(Saldo, "##,##0.00")
      Me.Txt15.Text = Format(Saldo, "##,##0.00")
      Me.Txt16.Text = Format(Saldo, "##,##0.00")
      Me.Txt17.Text = Format(Saldo, "##,##0.00")
      Me.Txt18.Text = Format(Saldo, "##,##0.00")
      Me.Txt19.Text = Format(Saldo, "##,##0.00")
      Me.Txt20.Text = Format(Saldo, "##,##0.00")
      Me.Txt21.Text = Format(Saldo, "##,##0.00")
      Me.Txt22.Text = Format(Saldo, "##,##0.00")
      Me.Txt23.Text = Format(Saldo, "##,##0.00")
      Me.Txt24.Text = Format(Saldo, "##,##0.00")
      
      Me.Txt25.Text = Format(Saldo, "##,##0.00")
      Me.Txt26.Text = Format(Saldo, "##,##0.00")
      Me.Txt27.Text = Format(Saldo, "##,##0.00")
      Me.Txt28.Text = Format(Saldo, "##,##0.00")
      Me.Txt29.Text = Format(Saldo, "##,##0.00")
      Me.Txt30.Text = Format(Saldo, "##,##0.00")
      Me.Txt31.Text = Format(Saldo, "##,##0.00")
      Me.Txt32.Text = Format(Saldo, "##,##0.00")
      Me.Txt33.Text = Format(Saldo, "##,##0.00")
      Me.Txt34.Text = Format(Saldo, "##,##0.00")
      Me.Txt35.Text = Format(Saldo, "##,##0.00")
      Me.Txt36.Text = Format(Saldo, "##,##0.00")
      
     Me.Text1 = Format(Saldo, "##,##0.00")
     Me.Text2 = Format(Saldo, "##,##0.00")
     Me.Text3 = Format(Saldo, "##,##0.00")
     Me.Text4 = Format(Saldo, "##,##0.00")
     Me.Text5 = Format(Saldo, "##,##0.00")
     Me.Text6 = Format(Saldo, "##,##0.00")
     Me.Text7 = Format(Saldo, "##,##0.00")
     Me.Text8 = Format(Saldo, "##,##0.00")
     Me.Text9 = Format(Saldo, "##,##0.00")
     Me.Text10 = Format(Saldo, "##,##0.00")
     Me.Text11 = Format(Saldo, "##,##0.00")
     Me.Text12 = Format(Saldo, "##,##0.00")
  
     
     Me.Text13 = Format(Saldo, "##,##0.00")
     Me.Text14 = Format(Saldo, "##,##0.00")
     Me.Text15 = Format(Saldo, "##,##0.00")
     Me.Text16 = Format(Saldo, "##,##0.00")
     Me.Text17 = Format(Saldo, "##,##0.00")
     Me.Text18 = Format(Saldo, "##,##0.00")
     Me.Text19 = Format(Saldo, "##,##0.00")
     Me.Text20 = Format(Saldo, "##,##0.00")
     Me.Text21 = Format(Saldo, "##,##0.00")
     Me.Text22 = Format(Saldo, "##,##0.00")
     Me.Text23 = Format(Saldo, "##,##0.00")
     Me.Text24 = Format(Saldo, "##,##0.00")
  
     Me.Text25 = Format(Saldo, "##,##0.00")
     Me.Text26 = Format(Saldo, "##,##0.00")
     Me.Text27 = Format(Saldo, "##,##0.00")
     Me.Text28 = Format(Saldo, "##,##0.00")
     Me.Text29 = Format(Saldo, "##,##0.00")
     Me.Text30 = Format(Saldo, "##,##0.00")
     Me.Text31 = Format(Saldo, "##,##0.00")
     Me.Text32 = Format(Saldo, "##,##0.00")
     Me.Text33 = Format(Saldo, "##,##0.00")
     Me.Text34 = Format(Saldo, "##,##0.00")
     Me.Text35 = Format(Saldo, "##,##0.00")
     Me.Text36 = Format(Saldo, "##,##0.00")
  
  
  
  
Else
    TipoMoneda = "Córdobas"  'DtaCuentas.Recordset("TipoMoneda")
    TipoCuenta = "Gastos" 'DtaCuentas.Recordset("TipoCuenta")
   Select Case TipoMoneda
      Case "Dólares"
         Fecha = Format(Now, "DD/MM/YYYY")
         NumFecha1 = Fecha
         Me.DtaConsulta.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = " & NumFecha1 & "))"
         Me.DtaConsulta.Refresh
         If Not DtaConsulta.Recordset.EOF Then
           MontoTasa = Me.DtaConsulta.Recordset!MontoLibras
         End If
      Case "Libras"
         MontoTasa = 1
   End Select

' If Not IsNull(Me.DtaCuentas.Recordset("DescripcionCuentas")) Then
'  Me.TxtDescripcion.Text = Me.DtaCuentas.Recordset("DescripcionCuentas")
' End If
 
 
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 1))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  FechaIni = "01/" & Month(Me.DtaPeriodos.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaPeriodos.Recordset("FechaPeriodo"))
  FechaFin = Me.DtaPeriodos.Recordset("FechaPeriodo")
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
'  Me.DtaConsulta.RecordSource = "SELECT CodCuentas, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "YYYY-MM-DD") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "YYYY-MM-DD") & "', 102)) GROUP BY FacturaNo HAVING (FacturaNo = '" & CodigoCuenta & "') "
  Me.DtaConsulta.RecordSource = "SELECT SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE   (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "YYYY-MM-DD") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "YYYY-MM-DD") & "', 102)) GROUP BY FacturaNo HAVING (FacturaNo = '" & CodigoCuenta & "')"
  Me.DtaConsulta.Refresh
'InputBox "", "", Me.DtaConsulta.RecordSource
  If Not Me.DtaConsulta.Recordset.EOF Then
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
    Debito = Me.DtaConsulta.Recordset("MDebito")
    Else
     Debito = 0
   End If
  If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
   Credito = Me.DtaConsulta.Recordset("MCredito")
  Else
   Credito = 0
  End If
   Select Case TipoMoneda
     Case "Dólares"
        If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
         Saldo = (Debito - Credito) / MontoTasa
        Else
         Saldo = (Credito - Debito) / MontoTasa
        End If
     Case "Libras"
         If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
            Saldo = (Debito - Credito)
         Else
            Saldo = (Credito - Debito)
         End If
     Case "Córdobas"
         If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
            Saldo = (Debito - Credito)
         Else
            Saldo = (Credito - Debito)
         End If
   End Select
   Total1 = Total1 + Saldo
  Else
   Saldo = 0#
 End If
    Select Case Periodo
      Case 1: Me.Txt1.Text = Format(Saldo, "##,##0.00")
      Case 2: Me.Txt2.Text = Format(Saldo, "##,##0.00")
      Case 3: Me.Txt3.Text = Format(Saldo, "##,##0.00")
      Case 4: Me.Txt4.Text = Format(Saldo, "##,##0.00")
      Case 5: Me.Txt5.Text = Format(Saldo, "##,##0.00")
      Case 6: Me.Txt6.Text = Format(Saldo, "##,##0.00")
      Case 7: Me.Txt7.Text = Format(Saldo, "##,##0.00")
      Case 8: Me.Txt8.Text = Format(Saldo, "##,##0.00")
      Case 9: Me.Txt9.Text = Format(Saldo, "##,##0.00")
      Case 10: Me.Txt10.Text = Format(Saldo, "##,##0.00")
      Case 11: Me.Txt11.Text = Format(Saldo, "##,##0.00")
      Case 12: Me.Txt12.Text = Format(Saldo, "##,##0.00")
 
   
    End Select

  Me.DtaPeriodos.Recordset.MoveNext
 Loop

 Me.LblTotal1.Caption = Format(Total1, "##,##0.00")
 '//////////////////////////////////////////////////////////////
 '//////////////Datos del segundo periodo////////////////////////
 '//////////////////////////////////////////////////////////////
  Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 2))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  FechaIni = "01/" & Month(Me.DtaPeriodos.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaPeriodos.Recordset("FechaPeriodo"))
  FechaFin = Me.DtaPeriodos.Recordset("FechaPeriodo")
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
  
  
  Me.DtaConsulta.RecordSource = "SELECT SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE   (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "YYYY-MM-DD") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "YYYY-MM-DD") & "', 102)) GROUP BY FacturaNo HAVING (FacturaNo = '" & CodigoCuenta & "')"
  Me.DtaConsulta.Refresh

  If Not Me.DtaConsulta.Recordset.EOF Then
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
    Debito = Me.DtaConsulta.Recordset("MDebito")
    Else
     Debito = 0
   End If
  If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
   Credito = Me.DtaConsulta.Recordset("MCredito")
  Else
   Credito = 0
  End If
   Select Case TipoMoneda
     Case "Dólares"
        If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
         Saldo = (Debito - Credito) / MontoTasa
        Else
         Saldo = (Credito - Debito) / MontoTasa
        End If
     Case "Libras"
         If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
            Saldo = (Debito - Credito)
         Else
            Saldo = (Credito - Debito)
         End If
     Case "Córdobas"
         If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
            Saldo = (Debito - Credito)
         Else
            Saldo = (Credito - Debito)
         End If
   End Select
   Total1 = Total1 + Saldo
  Else
   Saldo = 0#
 End If
    Select Case Periodo
      Case 1: Me.Txt13.Text = Format(Saldo, "##,##0.00")
      Case 2: Me.Txt14.Text = Format(Saldo, "##,##0.00")
      Case 3: Me.Txt15.Text = Format(Saldo, "##,##0.00")
      Case 4: Me.Txt16.Text = Format(Saldo, "##,##0.00")
      Case 5: Me.Txt17.Text = Format(Saldo, "##,##0.00")
      Case 6: Me.Txt18.Text = Format(Saldo, "##,##0.00")
      Case 7: Me.Txt19.Text = Format(Saldo, "##,##0.00")
      Case 8: Me.Txt20.Text = Format(Saldo, "##,##0.00")
      Case 9: Me.Txt21.Text = Format(Saldo, "##,##0.00")
      Case 10: Me.Txt22.Text = Format(Saldo, "##,##0.00")
      Case 11: Me.Txt23.Text = Format(Saldo, "##,##0.00")
      Case 12: Me.Txt24.Text = Format(Saldo, "##,##0.00")
 
   
    End Select
    
  Me.DtaPeriodos.Recordset.MoveNext
 Loop

 Me.LblTotal2 = Format(Total2, "##,##0.00")
 '//////////////////////////////////////////////////////////////
 '//////////////Datos del tercer periodo////////////////////////
 '//////////////////////////////////////////////////////////////
  Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 3))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  FechaIni = "01/" & Month(Me.DtaPeriodos.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaPeriodos.Recordset("FechaPeriodo"))
  FechaFin = Me.DtaPeriodos.Recordset("FechaPeriodo")
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
'  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito, Transacciones.NPeriodo From Transacciones GROUP BY Transacciones.CodCuentas, Transacciones.NPeriodo HAVING (((Transacciones.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.NPeriodo)=" & NumeroPeriodo & "))"
  Me.DtaConsulta.RecordSource = "SELECT SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "YYYY-MM-DD") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "YYYY-MM-DD") & "', 102)) GROUP BY FacturaNo HAVING (FacturaNo = '" & CodigoCuenta & "') "
  Me.DtaConsulta.Refresh

  If Not Me.DtaConsulta.Recordset.EOF Then
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
    Debito = Me.DtaConsulta.Recordset("MDebito")
    Else
     Debito = 0
   End If
  If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
   Credito = Me.DtaConsulta.Recordset("MCredito")
  Else
   Credito = 0
  End If
   Select Case TipoMoneda
     Case "Dólares"
         Saldo = (Debito - Credito) / MontoTasa
     Case "Libras"
         Saldo = (Debito - Credito)
   End Select
   Total1 = Total1 + Saldo
  Else
   Saldo = 0#
 End If
    Select Case Periodo
      Case 1: Me.Txt25.Text = Format(Saldo, "##,##0.00")
      Case 2: Me.Txt26.Text = Format(Saldo, "##,##0.00")
      Case 3: Me.Txt27.Text = Format(Saldo, "##,##0.00")
      Case 4: Me.Txt28.Text = Format(Saldo, "##,##0.00")
      Case 5: Me.Txt29.Text = Format(Saldo, "##,##0.00")
      Case 6: Me.Txt30.Text = Format(Saldo, "##,##0.00")
      Case 7: Me.Txt31.Text = Format(Saldo, "##,##0.00")
      Case 8: Me.Txt32.Text = Format(Saldo, "##,##0.00")
      Case 9: Me.Txt33.Text = Format(Saldo, "##,##0.00")
      Case 10: Me.Txt34.Text = Format(Saldo, "##,##0.00")
      Case 11: Me.Txt35.Text = Format(Saldo, "##,##0.00")
      Case 12: Me.Txt36.Text = Format(Saldo, "##,##0.00")
 
   
    End Select
    
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 Me.LblTotal3.Caption = Format(Total3, "##,##0.00")
 '//////////////////////////////////////////////////////////////
 '//////////////Datos del PRIMER periodo PRESUPUESTO////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 1))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Not Me.DtaConsulta.Recordset.EOF Then
   Saldo = Me.DtaConsulta.Recordset!MontoPresupuestado
   Total4 = Total4 + Saldo
  Else
   Saldo = 0#
 End If
    Select Case Periodo
      Case 1: Me.Text1 = Format(Saldo, "##,##0.00")
      Case 2: Me.Text2 = Format(Saldo, "##,##0.00")
      Case 3: Me.Text3 = Format(Saldo, "##,##0.00")
      Case 4: Me.Text4 = Format(Saldo, "##,##0.00")
      Case 5: Me.Text5 = Format(Saldo, "##,##0.00")
      Case 6: Me.Text6 = Format(Saldo, "##,##0.00")
      Case 7: Me.Text7 = Format(Saldo, "##,##0.00")
      Case 8: Me.Text8 = Format(Saldo, "##,##0.00")
      Case 9: Me.Text9 = Format(Saldo, "##,##0.00")
      Case 10: Me.Text10 = Format(Saldo, "##,##0.00")
      Case 11: Me.Text11 = Format(Saldo, "##,##0.00")
      Case 12: Me.Text12 = Format(Saldo, "##,##0.00")
 
   
    End Select
    
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
' Me.LblTotal4.Caption = Format(Total4, "##,##0.00")
 
  '//////////////////////////////////////////////////////////////
 '//////////////Datos del SEGUNDO periodo PRESUPUESTO////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 2))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Not Me.DtaConsulta.Recordset.EOF Then
   Saldo = Me.DtaConsulta.Recordset!MontoPresupuestado
   Total5 = Total5 + Saldo
  Else
   Saldo = 0#
 End If
    Select Case Periodo
      Case 1: Me.Text13 = Format(Saldo, "##,##0.00")
      Case 2: Me.Text14 = Format(Saldo, "##,##0.00")
      Case 3: Me.Text15 = Format(Saldo, "##,##0.00")
      Case 4: Me.Text16 = Format(Saldo, "##,##0.00")
      Case 5: Me.Text17 = Format(Saldo, "##,##0.00")
      Case 6: Me.Text18 = Format(Saldo, "##,##0.00")
      Case 7: Me.Text19 = Format(Saldo, "##,##0.00")
      Case 8: Me.Text20 = Format(Saldo, "##,##0.00")
      Case 9: Me.Text21 = Format(Saldo, "##,##0.00")
      Case 10: Me.Text22 = Format(Saldo, "##,##0.00")
      Case 11: Me.Text23 = Format(Saldo, "##,##0.00")
      Case 12: Me.Text24 = Format(Saldo, "##,##0.00")
 
   
    End Select
    
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
' Me.LblTotal5.Caption = Format(Total5, "##,##0.00")
 
   '//////////////////////////////////////////////////////////////
 '//////////////Datos del TERCER periodo PRESUPUESTO////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 3))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = KeyPrincipal
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Not Me.DtaConsulta.Recordset.EOF Then
   Saldo = Me.DtaConsulta.Recordset!MontoPresupuestado
   Total6 = Total6 + Saldo
  Else
   Saldo = 0#
 End If
    Select Case Periodo
      Case 1: Me.Text25 = Format(Saldo, "##,##0.00")
      Case 2: Me.Text26 = Format(Saldo, "##,##0.00")
      Case 3: Me.Text27 = Format(Saldo, "##,##0.00")
      Case 4: Me.Text28 = Format(Saldo, "##,##0.00")
      Case 5: Me.Text29 = Format(Saldo, "##,##0.00")
      Case 6: Me.Text30 = Format(Saldo, "##,##0.00")
      Case 7: Me.Text31 = Format(Saldo, "##,##0.00")
      Case 8: Me.Text32 = Format(Saldo, "##,##0.00")
      Case 9: Me.Text33 = Format(Saldo, "##,##0.00")
      Case 10: Me.Text34 = Format(Saldo, "##,##0.00")
      Case 11: Me.Text35 = Format(Saldo, "##,##0.00")
      Case 12: Me.Text36 = Format(Saldo, "##,##0.00")
 
   
    End Select
    
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
' Me.LblTotal6.Caption = Format(Total6, "##,##0.00")
 
 Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 1) And ((PresupuestoAnual.CodigoCuenta) = '" & KeyPrincipal & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If Not DtaPresupuestoAnual.Recordset.EOF Then
   Me.TxtTotal1.Text = Format(DtaPresupuestoAnual.Recordset!MontoAnual, "##,##0.00")

 End If
 
  Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 2) And ((PresupuestoAnual.CodigoCuenta) = '" & KeyPrincipal & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If Not DtaPresupuestoAnual.Recordset.EOF Then
   Me.TxtTotal2.Text = Format(DtaPresupuestoAnual.Recordset!MontoAnual, "##,##0.00")

 End If
 
  Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 3) And ((PresupuestoAnual.CodigoCuenta) = '" & KeyPrincipal & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If Not DtaPresupuestoAnual.Recordset.EOF Then
   Me.TxtTotal3.Text = Format(DtaPresupuestoAnual.Recordset!MontoAnual, "##,##0.00")

 End If
 
 
End If 'IF FINAL


Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd

Enfoque = 1
QUIEN = "Load"
TotalAño1 = 0
TotalAño2 = 0
TotalAño3 = 0


Dim NodX As Node
Dim Relatives As String, RelationsShips As String
Dim LLave As String, Texto As String, Imagen1 As Integer
Dim Imagen2 As Integer
Dim Año As String

With Me.DtaGrupos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaCuentas
   .ConnectionString = Conexion
End With

With Me.DtaSaldoCuenta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


 Me.TDBGridCuentas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridCuentas.OddRowStyle.BackColor = &H80000005
 Me.TDBGridCuentas.AlternatingRowStyle = True

i = 1
 ReDim MatrizCuentas(100)
 Me.DtaGrupos.RecordSource = "SELECT EstructuraPresupuesto.KeyGrupo, EstructuraPresupuesto.KeyGrupoSuperior, EstructuraPresupuesto.Child, EstructuraPresupuesto.DescripcionGrupo, EstructuraPresupuesto.Imagen1, EstructuraPresupuesto.Imagen2 From EstructuraPresupuesto ORDER BY EstructuraPresupuesto.KeyGrupo"
 Me.DtaGrupos.Refresh
 Do While Not Me.DtaGrupos.Recordset.EOF
   If Not IsNull(Me.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
    Relatives = Me.DtaGrupos.Recordset("KeyGrupoSuperior")
   Else
     Relatives = ""
   End If
   If Not IsNull(Me.DtaGrupos.Recordset("Child")) Then
     RelationsShips = Me.DtaGrupos.Recordset("Child")
   Else
     RelationsShips = ""
   End If
   LLave = Me.DtaGrupos.Recordset("KeyGrupo")
   Texto = Me.DtaGrupos.Recordset("DescripcionGrupo")
   Imagen1 = Me.DtaGrupos.Recordset("Imagen1")
   Imagen2 = Me.DtaGrupos.Recordset("Imagen2")
   
   If Relatives = "" And RelationsShips = "" Then
     Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Texto, Imagen1, Imagen2)
   Else
     Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, LLave, Texto, Imagen1, Imagen2)
   End If
   
  Me.DtaGrupos.Recordset.MoveNext
 Loop



KeyPrincipal = "A"
Me.TreeView1.Nodes(Me.TreeView1.Nodes.Count).EnsureVisible
NodoBase = True
'Me.DtaCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & KeyPrincipal & "')) ORDER BY Cuentas.CodCuentas"
'Me.DtaCuentas.Refresh
Me.TDBGridCuentas.Columns(0).Caption = "Còdigo Cuentas"
Me.TDBGridCuentas.Columns(0).Width = 2000
Me.TDBGridCuentas.Columns(1).Caption = "Descripcion Cuentas"
Me.TDBGridCuentas.Columns(1).Width = 4350

Me.CmdProcesar.Visible = False




On Error GoTo TipoErrs
'pasa algunos controles con el skin
Dim Control As Control
For Each Control In Me.Controls
    If TypeOf Control Is Frame Then
        MDIPrimero.Skin1.ApplySkin Control.hWnd
    ElseIf TypeOf Control Is TextBox Then
        MDIPrimero.Skin1.ApplySkin Control.hWnd
    ElseIf TypeOf Control Is CommandButton Then
        MDIPrimero.Skin1.ApplySkin Control.hWnd
    End If
Next Control


With Me.DtaPresupuestoAnual
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Accesos"
   .Refresh
End With


With Me.DtaCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaPeriodos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
    .ConnectionString = Conexion
    .RecordSource = "Select * from Periodos"
    .Refresh
End With

With Me.DtaPresupuesto
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from EstructuraPresupuesto"
   .Refresh
End With

With Me.AdoPresupuesto
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT  Presupuesto.* From Presupuesto"
   .Refresh
End With

With Me.AdoPresupuestoAnual
   .ConnectionString = Conexion
End With

'Me.DtaCuentas.RecordSource = "SELECT Cuentas.TipoMoneda,Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas ORDER BY Cuentas.CodCuentas"
Me.DtaCuentas.RecordSource = "SELECT KeyGrupo, CodGrupo, KeyGrupoSuperior, Child, DescripcionGrupo, Imagen1, Imagen2 From EstructuraPresupuesto"
Me.DtaCuentas.Refresh
'LlenarDataCombos DtaCuentas, DBCliente, "CodCuentas", "CodCuentas"
'Me.DBCliente.ListField = "CodCuentas"
i = 1
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.Periodo, Periodos.FechaPeriodo From Periodos Where (((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3) And ((Periodos.Periodo) = 1)) ORDER BY Periodos.NumeroTabla"
Me.DtaConsulta.Refresh
Do While Not Me.DtaConsulta.Recordset.EOF
 Año = Year(Me.DtaConsulta.Recordset!FechaPeriodo)
 If i = 1 Then
  Me.Lbl1.Caption = "Año " & Año
  Me.Lbl4.Caption = "Año " & Año
 ElseIf i = 2 Then
  Me.Lbl2.Caption = "Año " & Año
  Me.Lbl5.Caption = "Año " & Año
 ElseIf i = 3 Then
  Me.Lbl3.Caption = "Año " & Año
  Me.Lbl6.Caption = "Año " & Año
 End If
 
 Me.DtaConsulta.Recordset.MoveNext
 i = i + 1
Loop

QUIEN = "NoLoad"

Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub SmartButton1_Click()
DescripcionNodo = Me.TreeView1.SelectedItem
FrmEditaPresupuesto.Show 1
Me.TreeView1.SetFocus
End Sub

Private Sub SmartButton2_Click()
Me.TreeView1.SetFocus
 If NodoBase = True Then
  FrmCreaNodosPresupuesto.Option1.Enabled = False
  FrmCreaNodosPresupuesto.Option2.Value = True
 End If
 FrmCreaNodosPresupuesto.Show 1
 Me.TreeView1.Sorted = True
End Sub

Private Sub Text1_Change()
Enfoque = 1

If Me.Text1.Text = "" Then
  Me.Text1.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text1_Click()
 Enfoque = 1
End Sub

Private Sub Text10_Change()
Enfoque = 10

If Me.Text10.Text = "" Then
  Me.Text10.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text10_Click()
Enfoque = 10
End Sub

Private Sub Text11_Change()
Enfoque = 11

If Me.Text11.Text = "" Then
  Me.Text11.Text = "0.00"
End If
Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text11_Click()
Enfoque = 11
End Sub

Private Sub Text12_Change()
Enfoque = 12

If Me.Text12.Text = "" Then
  Me.Text12.Text = "0.00"
End If
Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text12_Click()
Enfoque = 12
End Sub

Private Sub Text13_Change()
Enfoque = 13

If Me.Text13.Text = "" Then
  Me.Text13.Text = "0.00"
End If

Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text13_Click()
Enfoque = 13
End Sub

Private Sub Text14_Change()
Enfoque = 14

If Me.Text14.Text = "" Then
  Me.Text14.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text14_Click()
Enfoque = 14
End Sub

Private Sub Text15_Change()
Enfoque = 15

If Me.Text15.Text = "" Then
  Me.Text15.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text15_Click()
Enfoque = 15
End Sub

Private Sub Text16_Change()
Enfoque = 16

If Me.Text16.Text = "" Then
  Me.Text16.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text16_Click()
Enfoque = 16
End Sub

Private Sub Text17_Change()
Enfoque = 17

If Me.Text17.Text = "" Then
  Me.Text17.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text17_Click()
Enfoque = 17
End Sub

Private Sub Text18_Change()
Enfoque = 18

If Me.Text18.Text = "" Then
  Me.Text18.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text18_Click()
Enfoque = 18
End Sub

Private Sub Text19_Change()
Enfoque = 19

If Me.Text19.Text = "" Then
  Me.Text19.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text19_Click()
Enfoque = 19
End Sub

Private Sub Text2_Change()
Enfoque = 2
If Me.Text2.Text = "" Then
  Me.Text2.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text2_Click()
Enfoque = 2
End Sub

Private Sub Text20_Change()
Enfoque = 20

If Me.Text20.Text = "" Then
  Me.Text20.Text = "0.00"
End If

Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text20_Click()
Enfoque = 20
End Sub

Private Sub Text21_Change()
Enfoque = 21

If Me.Text21.Text = "" Then
  Me.Text21.Text = "0.00"
End If

Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text21_Click()
Enfoque = 21
End Sub

Private Sub Text22_Change()
Enfoque = 22

If Me.Text22.Text = "" Then
  Me.Text22.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text22_Click()
Enfoque = 22
End Sub

Private Sub Text23_Change()
Enfoque = 23

If Me.Text23.Text = "" Then
  Me.Text23.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text23_Click()
Enfoque = 23
End Sub

Private Sub Text24_Change()
Enfoque = 24

If Me.Text24.Text = "" Then
  Me.Text24.Text = "0.00"
End If
Me.TxtTotal2.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text24_Click()
Enfoque = 24
End Sub

Private Sub Text25_Change()
Enfoque = 25

If Me.Text25.Text = "" Then
  Me.Text25.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text25_Click()
Enfoque = 25
End Sub

Private Sub Text26_Change()
Enfoque = 26

If Me.Text26.Text = "" Then
  Me.Text26.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text26_Click()
Enfoque = 26
End Sub

Private Sub Text27_Change()
Enfoque = 27

If Me.Text27.Text = "" Then
  Me.Text27.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text27_Click()
Enfoque = 27
End Sub

Private Sub Text28_Change()
Enfoque = 28

If Me.Text28.Text = "" Then
  Me.Text28.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text28_Click()
Enfoque = 28
End Sub

Private Sub Text29_Change()
Enfoque = 29

If Me.Text29.Text = "" Then
  Me.Text29.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text29_Click()
Enfoque = 29
End Sub

Private Sub Text3_Change()
Enfoque = 3
If Me.Text4.Text = "" Then
  Me.Text4.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text3_Click()
Enfoque = 3
End Sub

Private Sub Text30_Change()
Enfoque = 30

If Me.Text30.Text = "" Then
  Me.Text30.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text30_Click()
Enfoque = 30
End Sub

Private Sub Text31_Change()
Enfoque = 31

If Me.Text31.Text = "" Then
  Me.Text31.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text31_Click()
Enfoque = 31
End Sub

Private Sub Text32_Change()
Enfoque = 32

If Me.Text32.Text = "" Then
  Me.Text32.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text32_Click()
Enfoque = 32
End Sub

Private Sub Text33_Change()
Enfoque = 33

If Me.Text33.Text = "" Then
  Me.Text33.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text33_Click()
Enfoque = 33
End Sub

Private Sub Text34_Change()
Enfoque = 34

If Me.Text34.Text = "" Then
  Me.Text34.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text34_Click()
Enfoque = 34
End Sub

Private Sub Text35_Change()
Enfoque = 35

If Me.Text35.Text = "" Then
  Me.Text35.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text35_Click()
Enfoque = 35
End Sub

Private Sub Text36_Change()
Enfoque = 36

If Me.Text36.Text = "" Then
  Me.Text36.Text = "0.00"
End If
Me.TxtTotal3.Text = Format(TotalAño(Enfoque), "##,##0.00")
End Sub

Private Sub Text36_Click()
Enfoque = 36
End Sub

Private Sub Text4_Change()
Enfoque = 4
If Me.Text4.Text = "" Then
  Me.Text4.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text4_Click()
Enfoque = 4
End Sub

Private Sub Text5_Change()
Enfoque = 5
If Me.Text5.Text = "" Then
  Me.Text5.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text5_Click()
Enfoque = 5
End Sub

Private Sub Text6_Change()
Enfoque = 6
If Me.Text6.Text = "" Then
  Me.Text6.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text6_Click()
Enfoque = 6
End Sub

Private Sub Text7_Change()
Enfoque = 7
If Me.Text7.Text = "" Then
  Me.Text7.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text7_Click()
Enfoque = 7
End Sub

Private Sub Text8_Change()
Enfoque = 8

If Me.Text8.Text = "" Then
  Me.Text8.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text8_Click()
Enfoque = 8
End Sub

Private Sub Text9_Change()
Enfoque = 9

If Me.Text9.Text = "" Then
  Me.Text9.Text = "0.00"
End If

Me.TxtTotal1.Text = Format(TotalAño(Enfoque), "##,##0.00")

End Sub

Private Sub Text9_Click()
Enfoque = 9
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
  Dim numero As Integer
  Dim Cadena1 As String, Cadena2 As String
  KeyPadre = ""
  KeyHijo = ""
  KeyNodoUltimo = ""
  KeyPrincipal = Node.Key

' Me.DtaCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & KeyPrincipal & "')) ORDER BY Cuentas.CodCuentas"
' Me.DtaCuentas.Refresh
'Me.TDBGridCuentas.Columns(0).Caption = "Còdigo Cuentas"
'Me.TDBGridCuentas.Columns(0).Width = 2000
'Me.TDBGridCuentas.Columns(1).Caption = "Descripcion Cuentas"
'Me.TDBGridCuentas.Columns(1).Width = 4350



If Len(KeyPrincipal) = 1 Then
    NodoBase = True
Else
    NodoBase = False
    KeyPadre = Node.Parent.Key
End If

If QUIEN <> "Load" Then
  Me.Cargar_Presupuesto
End If


'SaldosPeriodos (KeyPrincipal)

End Sub
Private Sub SaldosPeriodos(KeyPrincipal As String)

     Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 1))"
     Me.DtaPeriodos.Refresh
     Do While Not Me.DtaPeriodos.Recordset.EOF
      '////////////////////////////////SELECCIONO EL PERIODO ///////////////////////////////////////
      FechaIni = "01/" & Month(Me.DtaPeriodos.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaPeriodos.Recordset("FechaPeriodo"))
      FechaFin = Me.DtaPeriodos.Recordset("FechaPeriodo")
      NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
      Periodo = Me.DtaPeriodos.Recordset!Periodo
      
      CodigoCuenta = KeyPrincipal
      Me.DtaConsulta.RecordSource = "SELECT CodCuentas, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "YYYY-MM-DD") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "YYYY-MM-DD") & "', 102)) GROUP BY CodCuentas HAVING (CodCuentas = '" & CodigoCuenta & "') "
      Me.DtaConsulta.Refresh

      If Not Me.DtaConsulta.Recordset.EOF Then
        If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
        Debito = Me.DtaConsulta.Recordset("MDebito")
        Else
         Debito = 0
       End If
      If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
       Credito = Me.DtaConsulta.Recordset("MCredito")
      Else
       Credito = 0
      End If
       Select Case TipoMoneda
         Case "Dólares"
            If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
             Saldo = (Debito - Credito) / MontoTasa
            Else
             Saldo = (Credito - Debito) / MontoTasa
            End If
         Case "Libras"
             If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                Saldo = (Debito - Credito)
             Else
                Saldo = (Credito - Debito)
             End If
         Case "Córdobas"
             If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                Saldo = (Debito - Credito)
             Else
                Saldo = (Credito - Debito)
             End If
       End Select
       Total1 = Total1 + Saldo
      Else
       Saldo = 0#
     End If
        Select Case Periodo
          Case 1: Me.Txt1.Text = Format(Saldo, "##,##0.00")
          Case 2: Me.Txt2.Text = Format(Saldo, "##,##0.00")
          Case 3: Me.Txt3.Text = Format(Saldo, "##,##0.00")
          Case 4: Me.Txt4.Text = Format(Saldo, "##,##0.00")
          Case 5: Me.Txt5.Text = Format(Saldo, "##,##0.00")
          Case 6: Me.Txt6.Text = Format(Saldo, "##,##0.00")
          Case 7: Me.Txt7.Text = Format(Saldo, "##,##0.00")
          Case 8: Me.Txt8.Text = Format(Saldo, "##,##0.00")
          Case 9: Me.Txt9.Text = Format(Saldo, "##,##0.00")
          Case 10: Me.Txt10.Text = Format(Saldo, "##,##0.00")
          Case 11: Me.Txt11.Text = Format(Saldo, "##,##0.00")
          Case 12: Me.Txt12.Text = Format(Saldo, "##,##0.00")
     
       
        End Select
    
      Me.DtaPeriodos.Recordset.MoveNext
     Loop
    
     Me.LblTotal1.Caption = Format(Total1, "##,##0.00")




End Sub



Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim numero As Integer
  Dim Cadena1 As String, Cadena2 As String
  KeyPadre = ""
  KeyHijo = ""
  KeyNodoUltimo = ""
  KeyPrincipal = Node.Key

' Me.DtaCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.KeyGrupo) = '" & KeyPrincipal & "')) ORDER BY Cuentas.CodCuentas"
' Me.DtaCuentas.Refresh
'Me.TDBGridCuentas.Columns(0).Caption = "Còdigo Cuentas"
'Me.TDBGridCuentas.Columns(0).Width = 2000
'Me.TDBGridCuentas.Columns(1).Caption = "Descripcion Cuentas"
'Me.TDBGridCuentas.Columns(1).Width = 4350


If Len(KeyPrincipal) = 1 Then
    NodoBase = True
Else
    NodoBase = False
    KeyPadre = Node.Parent.Key
End If

  Me.Cargar_Presupuesto


End Sub
