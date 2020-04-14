VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmConfiguraCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion de la Impresion de Cheques"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmConfiguraCheque.frx":0000
   ScaleHeight     =   7050
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Y22 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      MaxLength       =   3
      TabIndex        =   101
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox X22 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11520
      MaxLength       =   3
      TabIndex        =   100
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox Y21 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10440
      MaxLength       =   3
      TabIndex        =   97
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox X21 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10080
      MaxLength       =   3
      TabIndex        =   96
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox Y20 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12240
      MaxLength       =   3
      TabIndex        =   93
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox X20 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      MaxLength       =   3
      TabIndex        =   92
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox Y19 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   89
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox X19 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10680
      MaxLength       =   3
      TabIndex        =   88
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox TxtNCaracteresConcepto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   85
      Text            =   "5"
      Top             =   2280
      Width           =   480
   End
   Begin VB.TextBox TxtNCaracteres 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   83
      Text            =   "5"
      Top             =   2160
      Width           =   480
   End
   Begin VB.TextBox X17 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      MaxLength       =   3
      TabIndex        =   79
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Y17 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12240
      MaxLength       =   3
      TabIndex        =   78
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox X18 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      MaxLength       =   3
      TabIndex        =   75
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Y18 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      MaxLength       =   3
      TabIndex        =   74
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox X16 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12000
      MaxLength       =   3
      TabIndex        =   71
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Y16 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12360
      MaxLength       =   3
      TabIndex        =   70
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox X15 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      MaxLength       =   3
      TabIndex        =   67
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Y15 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      MaxLength       =   3
      TabIndex        =   66
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox X14 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   63
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Y14 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   62
      Top             =   840
      Width           =   375
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   1680
      TabIndex        =   61
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "TxtNLineas"
      BuddyDispid     =   196609
      OrigLeft        =   480
      OrigTop         =   2280
      OrigRight       =   735
      OrigBottom      =   2535
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox TxtNLineas 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   59
      Text            =   "5"
      Top             =   1800
      Width           =   480
   End
   Begin VB.TextBox X13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      MaxLength       =   3
      TabIndex        =   56
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Y13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12240
      MaxLength       =   3
      TabIndex        =   55
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox X12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10680
      MaxLength       =   3
      TabIndex        =   52
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Y12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11040
      MaxLength       =   3
      TabIndex        =   51
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox X11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   48
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Y11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   47
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox X10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   44
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Y10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   43
      Top             =   4320
      Width           =   375
   End
   Begin MSAdodcLib.Adodc AdoCordenadas 
      Height          =   375
      Left            =   240
      Top             =   7920
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "AdoCordenadas"
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
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   240
      TabIndex        =   42
      Top             =   6480
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc AdoBancos 
      Height          =   375
      Left            =   4080
      Top             =   8040
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
      Caption         =   "AdoBancos"
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
   Begin VB.TextBox TxtNombreBanco 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5760
      Width           =   6135
   End
   Begin VB.TextBox X9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      MaxLength       =   3
      TabIndex        =   16
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Y9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   3
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Y8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      MaxLength       =   3
      TabIndex        =   15
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox X8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11520
      MaxLength       =   3
      TabIndex        =   14
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox Y7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10440
      MaxLength       =   3
      TabIndex        =   13
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox X7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10080
      MaxLength       =   3
      TabIndex        =   12
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox Y6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   11
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox X6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   10
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Y5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11520
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox X5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11160
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Y2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9240
      MaxLength       =   3
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox X2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox X4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11520
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Y4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11880
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox X3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11160
      MaxLength       =   3
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Y3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11520
      MaxLength       =   3
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Y1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox X1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin MSDataListLib.DataCombo DBCodigo 
      Bindings        =   "FrmConfiguraCheque.frx":FB086
      Height          =   315
      Left            =   120
      TabIndex        =   40
      Top             =   5760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   285
      Left            =   1680
      TabIndex        =   82
      Top             =   2160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "TxtNLineas"
      BuddyDispid     =   196609
      OrigLeft        =   480
      OrigTop         =   2280
      OrigRight       =   735
      OrigBottom      =   2535
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown3 
      Height          =   285
      Left            =   5520
      TabIndex        =   86
      Top             =   2280
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "TxtNLineas"
      BuddyDispid     =   196609
      OrigLeft        =   480
      OrigTop         =   2280
      OrigRight       =   735
      OrigBottom      =   2535
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y22"
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
      Left            =   11880
      TabIndex        =   103
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X22"
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
      Left            =   11520
      TabIndex        =   102
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y21"
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
      Left            =   10440
      TabIndex        =   99
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X21"
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
      Left            =   10080
      TabIndex        =   98
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y20"
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
      Left            =   12240
      TabIndex        =   95
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X20"
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
      Left            =   11880
      TabIndex        =   94
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y19"
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
      Left            =   11040
      TabIndex        =   91
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X19"
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
      Left            =   10680
      TabIndex        =   90
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caracteres Concepto"
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
      Left            =   3960
      TabIndex        =   87
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caracteres Lineas"
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
      Left            =   120
      TabIndex        =   84
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X17"
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
      Left            =   11880
      TabIndex        =   81
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y17"
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
      Left            =   12240
      TabIndex        =   80
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "X18"
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
      Left            =   6480
      TabIndex        =   77
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y18"
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
      Left            =   6840
      TabIndex        =   76
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X16"
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
      Left            =   12000
      TabIndex        =   73
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y16"
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
      Left            =   12360
      TabIndex        =   72
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X15"
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
      Left            =   8760
      TabIndex        =   69
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y15"
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
      Left            =   9120
      TabIndex        =   68
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X14"
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
      Left            =   5040
      TabIndex        =   65
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y14"
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
      Left            =   5400
      TabIndex        =   64
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lineas"
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
      Left            =   480
      TabIndex        =   60
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X13"
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
      Left            =   11880
      TabIndex        =   58
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y13"
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
      Left            =   12240
      TabIndex        =   57
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X12"
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
      Left            =   10680
      TabIndex        =   54
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y12"
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
      Left            =   11040
      TabIndex        =   53
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X11"
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
      Left            =   4440
      TabIndex        =   50
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y11"
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
      Left            =   4800
      TabIndex        =   49
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X10"
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
      Left            =   2280
      TabIndex        =   46
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y10"
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
      Left            =   2640
      TabIndex        =   45
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Banco"
      Height          =   255
      Left            =   3360
      TabIndex        =   39
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo del Banco"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X9"
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
      Left            =   600
      TabIndex        =   37
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y9"
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
      Left            =   960
      TabIndex        =   36
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Left            =   1560
      TabIndex        =   35
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaccion No"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y8"
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
      Left            =   11880
      TabIndex        =   33
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X8"
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
      Left            =   11520
      TabIndex        =   32
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y7"
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
      Left            =   10440
      TabIndex        =   31
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X7"
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
      Left            =   10080
      TabIndex        =   30
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y6"
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
      Left            =   840
      TabIndex        =   29
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X6"
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
      Left            =   480
      TabIndex        =   28
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y5"
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
      Left            =   11520
      TabIndex        =   27
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X5"
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
      Left            =   11160
      TabIndex        =   26
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y2"
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
      Left            =   9240
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X2"
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
      Left            =   8880
      TabIndex        =   24
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X4"
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
      Left            =   11520
      TabIndex        =   23
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y4"
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
      Left            =   11880
      TabIndex        =   22
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X3"
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
      Left            =   11160
      TabIndex        =   21
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y3"
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
      Left            =   11520
      TabIndex        =   20
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y1"
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
      Left            =   7920
      TabIndex        =   19
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X1"
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
      Left            =   7560
      TabIndex        =   18
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "FrmConfiguraCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGrabar_Click()
   Me.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8, X9, Y9,X10, Y10,X11, Y11 ,X12, Y12,X13, Y13,X14, Y14,X15, Y15,X16, Y16,X17, Y17, X18, Y18 ,X19, Y19, X20, Y20,X21, Y21, X22, Y22, NLineas,CaracteresLineas, CaracteresConcepto From CordenadasCheque WHERE (CodCuenta = '" & Me.DBCodigo.Text & "')"
   Me.AdoCordenadas.Refresh
   If Not Me.AdoCordenadas.Recordset.EOF Then
        Me.AdoCordenadas.Recordset("X1") = Me.X1.Text
        Me.AdoCordenadas.Recordset("Y1") = Me.Y1.Text
        Me.AdoCordenadas.Recordset("X2") = Me.X2.Text
        Me.AdoCordenadas.Recordset("Y2") = Me.Y2.Text
        Me.AdoCordenadas.Recordset("X3") = Me.X3.Text
        Me.AdoCordenadas.Recordset("Y3") = Me.Y3.Text
        Me.AdoCordenadas.Recordset("X4") = Me.X4.Text
        Me.AdoCordenadas.Recordset("Y4") = Me.Y4.Text
        Me.AdoCordenadas.Recordset("X5") = Me.X5.Text
        Me.AdoCordenadas.Recordset("Y5") = Me.Y5.Text
        Me.AdoCordenadas.Recordset("X6") = Me.X6.Text
        Me.AdoCordenadas.Recordset("Y6") = Me.Y6.Text
        Me.AdoCordenadas.Recordset("X7") = Me.X7.Text
        Me.AdoCordenadas.Recordset("Y7") = Me.Y7.Text
        Me.AdoCordenadas.Recordset("X8") = Me.X8.Text
        Me.AdoCordenadas.Recordset("Y8") = Me.Y8.Text
        Me.AdoCordenadas.Recordset("X9") = Me.X9.Text
        Me.AdoCordenadas.Recordset("Y9") = Me.Y9.Text
        Me.AdoCordenadas.Recordset("X10") = Me.X10.Text
        Me.AdoCordenadas.Recordset("Y10") = Me.Y10.Text
        Me.AdoCordenadas.Recordset("X11") = Me.X11.Text
        Me.AdoCordenadas.Recordset("Y11") = Me.Y11.Text
        Me.AdoCordenadas.Recordset("X12") = Me.X12.Text
        Me.AdoCordenadas.Recordset("Y12") = Me.Y12.Text
        Me.AdoCordenadas.Recordset("X13") = Me.X13.Text
        Me.AdoCordenadas.Recordset("Y13") = Me.Y13.Text
        Me.AdoCordenadas.Recordset("X14") = Me.X14.Text
        Me.AdoCordenadas.Recordset("Y14") = Me.Y14.Text
        Me.AdoCordenadas.Recordset("X15") = Me.X15.Text
        Me.AdoCordenadas.Recordset("Y15") = Me.Y15.Text
        Me.AdoCordenadas.Recordset("X16") = Me.X16.Text
        Me.AdoCordenadas.Recordset("Y16") = Me.Y16.Text
        Me.AdoCordenadas.Recordset("X17") = Me.X17.Text
        Me.AdoCordenadas.Recordset("Y17") = Me.Y17.Text
        Me.AdoCordenadas.Recordset("X18") = Me.X18.Text
        Me.AdoCordenadas.Recordset("Y18") = Me.Y18.Text
        Me.AdoCordenadas.Recordset("X19") = Me.X19.Text
        Me.AdoCordenadas.Recordset("Y19") = Me.Y19.Text
        Me.AdoCordenadas.Recordset("X20") = Me.X20.Text
        Me.AdoCordenadas.Recordset("Y20") = Me.Y20.Text
        Me.AdoCordenadas.Recordset("X21") = Me.X21.Text
        Me.AdoCordenadas.Recordset("Y21") = Me.Y21.Text
        Me.AdoCordenadas.Recordset("X22") = Me.X22.Text
        Me.AdoCordenadas.Recordset("Y22") = Me.Y22.Text
        Me.AdoCordenadas.Recordset("NLineas") = Me.TxtNLineas.Text
        Me.AdoCordenadas.Recordset("CaracteresLineas") = Me.TxtNCaracteres.Text
        Me.AdoCordenadas.Recordset("CaracteresConcepto") = Me.TxtNCaracteresConcepto.Text
        
       Me.AdoCordenadas.Recordset.Update
    Else
       Me.AdoCordenadas.Recordset.AddNew
        Me.AdoCordenadas.Recordset("CodCuenta") = Me.DBCodigo.Text
        Me.AdoCordenadas.Recordset("X1") = Me.X1.Text
        Me.AdoCordenadas.Recordset("Y1") = Me.Y1.Text
        Me.AdoCordenadas.Recordset("X2") = Me.X2.Text
        Me.AdoCordenadas.Recordset("Y2") = Me.Y2.Text
        Me.AdoCordenadas.Recordset("X3") = Me.X3.Text
        Me.AdoCordenadas.Recordset("Y3") = Me.Y3.Text
        Me.AdoCordenadas.Recordset("X4") = Me.X4.Text
        Me.AdoCordenadas.Recordset("Y4") = Me.Y4.Text
        Me.AdoCordenadas.Recordset("X5") = Me.X5.Text
        Me.AdoCordenadas.Recordset("Y5") = Me.Y5.Text
        Me.AdoCordenadas.Recordset("X6") = Me.X6.Text
        Me.AdoCordenadas.Recordset("Y6") = Me.Y6.Text
        Me.AdoCordenadas.Recordset("X7") = Me.X7.Text
        Me.AdoCordenadas.Recordset("Y7") = Me.Y7.Text
        Me.AdoCordenadas.Recordset("X8") = Me.X8.Text
        Me.AdoCordenadas.Recordset("Y8") = Me.Y8.Text
        Me.AdoCordenadas.Recordset("X9") = Me.X9.Text
        Me.AdoCordenadas.Recordset("Y9") = Me.Y9.Text
        Me.AdoCordenadas.Recordset("X10") = Me.X10.Text
        Me.AdoCordenadas.Recordset("Y10") = Me.Y10.Text
        Me.AdoCordenadas.Recordset("X11") = Me.X11.Text
        Me.AdoCordenadas.Recordset("Y11") = Me.Y11.Text
        Me.AdoCordenadas.Recordset("X12") = Me.X12.Text
        Me.AdoCordenadas.Recordset("Y12") = Me.Y12.Text
        Me.AdoCordenadas.Recordset("X13") = Me.X13.Text
        Me.AdoCordenadas.Recordset("Y13") = Me.Y13.Text
        Me.AdoCordenadas.Recordset("X14") = Me.X14.Text
        Me.AdoCordenadas.Recordset("Y14") = Me.Y14.Text
        Me.AdoCordenadas.Recordset("X15") = Me.X15.Text
        Me.AdoCordenadas.Recordset("Y15") = Me.Y15.Text
        Me.AdoCordenadas.Recordset("X16") = Me.X16.Text
        Me.AdoCordenadas.Recordset("Y16") = Me.Y16.Text
        Me.AdoCordenadas.Recordset("X17") = Me.X17.Text
        Me.AdoCordenadas.Recordset("Y17") = Me.Y17.Text
        Me.AdoCordenadas.Recordset("X18") = Me.X18.Text
        Me.AdoCordenadas.Recordset("Y18") = Me.Y18.Text
        Me.AdoCordenadas.Recordset("X19") = Me.X19.Text
        Me.AdoCordenadas.Recordset("Y19") = Me.Y19.Text
        Me.AdoCordenadas.Recordset("X20") = Me.X20.Text
        Me.AdoCordenadas.Recordset("Y20") = Me.Y20.Text
        Me.AdoCordenadas.Recordset("X21") = Me.X21.Text
        Me.AdoCordenadas.Recordset("Y21") = Me.Y21.Text
        Me.AdoCordenadas.Recordset("X22") = Me.X22.Text
        Me.AdoCordenadas.Recordset("Y22") = Me.Y22.Text
        Me.AdoCordenadas.Recordset("NLineas") = Me.TxtNLineas.Text
        Me.AdoCordenadas.Recordset("CaracteresLineas") = Me.TxtNCaracteres.Text
        Me.AdoCordenadas.Recordset("CaracteresConcepto") = Me.TxtNCaracteresConcepto.Text
       Me.AdoCordenadas.Recordset.Update
    
    End If
    
        Me.X1.Text = "0"
        Me.Y1.Text = "0"
        Me.X2.Text = "0"
        Me.Y2.Text = "0"
        Me.X3.Text = "0"
        Me.Y3.Text = "0"
        Me.X4.Text = "0"
        Me.Y4.Text = "0"
        Me.X5.Text = "0"
        Me.Y5.Text = "0"
        Me.X6.Text = "0"
        Me.Y6.Text = "0"
        Me.X7.Text = "0"
        Me.Y7.Text = "0"
        Me.X8.Text = "0"
        Me.Y8.Text = "0"
        Me.X9.Text = "0"
        Me.Y9.Text = "0"
        Me.X10.Text = "0"
        Me.Y10.Text = "0"
        Me.X11.Text = "0"
        Me.Y11.Text = "0"
        Me.X12.Text = "0"
        Me.Y12.Text = "0"
        Me.X13.Text = "0"
        Me.Y13.Text = "0"
        Me.X14.Text = "0"
        Me.Y14.Text = "0"
        Me.X15.Text = "0"
        Me.Y15.Text = "0"
        Me.X16.Text = "0"
        Me.Y16.Text = "0"
        Me.X17.Text = "0"
        Me.Y17.Text = "0"
        Me.X18.Text = "0"
        Me.Y18.Text = "0"
        Me.X19.Text = "0"
        Me.Y19.Text = "0"
        Me.X20.Text = "0"
        Me.Y20.Text = "0"
        Me.X21.Text = "0"
        Me.Y21.Text = "0"
        Me.X22.Text = "0"
        Me.Y22.Text = "0"
        Me.TxtNLineas.Text = "5"
        Me.TxtNCaracteres.Text = "10"
        Me.TxtNCaracteresConcepto.Text = "20"
End Sub

Private Sub DBCodigo_Change()
Criterio = "CodCuentas='" & Me.DBCodigo.Text & "'"
If Me.AdoBancos.Recordset.RecordCount > 0 Then Me.AdoBancos.Recordset.MoveFirst
Me.AdoBancos.Recordset.Find (Criterio)
 If Not Me.AdoBancos.Recordset.EOF Then
  Me.TxtNombreBanco.Text = Me.AdoBancos.Recordset("DescripcionCuentas")
  
   
   
   Me.AdoCordenadas.RecordSource = "SELECT CodCuenta, X1, Y1, X2, Y2, X3, Y3, X4, Y4, X5, Y5, X6, Y6, X7, Y7, X8, Y8,X9, Y9 ,X10, Y10,X11, Y11 ,X12, Y12,X13, Y13,X14, Y14,X15, Y15,X16, Y16,X17, Y17, X18, Y18,X19, Y19, X20, Y20,X21, Y21, X22, Y22, NLineas, CaracteresLineas, CaracteresConcepto From CordenadasCheque WHERE (CodCuenta = '" & Me.DBCodigo.Text & "')"
   Me.AdoCordenadas.Refresh
   If Not Me.AdoCordenadas.Recordset.EOF Then
        Me.X1.Text = Me.AdoCordenadas.Recordset("X1")
        Me.Y1.Text = Me.AdoCordenadas.Recordset("Y1")
        Me.X2.Text = Me.AdoCordenadas.Recordset("X2")
        Me.Y2.Text = Me.AdoCordenadas.Recordset("Y2")
        Me.X3.Text = Me.AdoCordenadas.Recordset("X3")
        Me.Y3.Text = Me.AdoCordenadas.Recordset("Y3")
        Me.X4.Text = Me.AdoCordenadas.Recordset("X4")
        Me.Y4.Text = Me.AdoCordenadas.Recordset("Y4")
        Me.X5.Text = Me.AdoCordenadas.Recordset("X5")
        Me.Y5.Text = Me.AdoCordenadas.Recordset("Y5")
        Me.X6.Text = Me.AdoCordenadas.Recordset("X6")
        Me.Y6.Text = Me.AdoCordenadas.Recordset("Y6")
        Me.X7.Text = Me.AdoCordenadas.Recordset("X7")
        Me.Y7.Text = Me.AdoCordenadas.Recordset("Y7")
        Me.X8.Text = Me.AdoCordenadas.Recordset("X8")
        Me.Y8.Text = Me.AdoCordenadas.Recordset("Y8")
        Me.X9.Text = Me.AdoCordenadas.Recordset("X9")
        Me.Y9.Text = Me.AdoCordenadas.Recordset("Y9")
        Me.X10.Text = Me.AdoCordenadas.Recordset("X10")
        Me.Y10.Text = Me.AdoCordenadas.Recordset("Y10")
        Me.X11.Text = Me.AdoCordenadas.Recordset("X11")
        Me.Y11.Text = Me.AdoCordenadas.Recordset("Y11")
        Me.X12.Text = Me.AdoCordenadas.Recordset("X12")
        Me.Y12.Text = Me.AdoCordenadas.Recordset("Y12")
        Me.X13.Text = Me.AdoCordenadas.Recordset("X13")
        Me.Y13.Text = Me.AdoCordenadas.Recordset("Y13")
        Me.X14.Text = Me.AdoCordenadas.Recordset("X14")
        Me.Y14.Text = Me.AdoCordenadas.Recordset("Y14")
        Me.X15.Text = Me.AdoCordenadas.Recordset("X15")
        Me.Y15.Text = Me.AdoCordenadas.Recordset("Y15")
        Me.X16.Text = Me.AdoCordenadas.Recordset("X16")
        Me.Y16.Text = Me.AdoCordenadas.Recordset("Y16")
        Me.X17.Text = Me.AdoCordenadas.Recordset("X17")
        Me.Y17.Text = Me.AdoCordenadas.Recordset("Y17")
        Me.X18.Text = Me.AdoCordenadas.Recordset("X18")
        Me.Y18.Text = Me.AdoCordenadas.Recordset("Y18")
        Me.X19.Text = Me.AdoCordenadas.Recordset("X19")
        Me.Y19.Text = Me.AdoCordenadas.Recordset("Y19")
        Me.X20.Text = Me.AdoCordenadas.Recordset("X20")
        Me.Y20.Text = Me.AdoCordenadas.Recordset("Y20")
        Me.X21.Text = Me.AdoCordenadas.Recordset("X21")
        Me.Y21.Text = Me.AdoCordenadas.Recordset("Y21")
        Me.X22.Text = Me.AdoCordenadas.Recordset("X22")
        Me.Y22.Text = Me.AdoCordenadas.Recordset("Y22")
        
        Me.TxtNLineas.Text = Me.AdoCordenadas.Recordset("NLineas")
        
        If Not IsNull(Me.AdoCordenadas.Recordset("CaracteresLineas")) Then
          Me.TxtNCaracteres.Text = Me.AdoCordenadas.Recordset("CaracteresLineas")
        End If
        
        If Not IsNull(Me.AdoCordenadas.Recordset("CaracteresConcepto")) Then
          Me.TxtNCaracteresConcepto.Text = Me.AdoCordenadas.Recordset("CaracteresConcepto")
        End If
        
        
        
    End If
  
  
  
  
  
  
 End If
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(219, 226, 242)
MDIPrimero.Skin1.ApplySkin Me.CmdGrabar.hWnd

With Me.AdoBancos
  .ConnectionString = Conexion
End With

With Me.AdoCordenadas
  .ConnectionString = Conexion
End With

Me.AdoBancos.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas Where (((Cuentas.TipoCuenta) = 'Bancos' Or (Cuentas.TipoCuenta) = 'Bancos')) ORDER BY Cuentas.CodCuentas"
Me.AdoBancos.Refresh
Me.DBCodigo.ListField = "CodCuentas"


End Sub

