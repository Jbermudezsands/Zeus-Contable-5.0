VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPresupuesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Presupuestos"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "FrmPresupuesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SmartButton5 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8880
      TabIndex        =   195
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   4800
      TabIndex        =   194
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   3240
      TabIndex        =   193
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1560
      TabIndex        =   192
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   191
      Top             =   5400
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DBCliente 
      Height          =   315
      Left            =   1200
      TabIndex        =   190
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
      Left            =   8760
      TabIndex        =   187
      Top             =   4560
      Width           =   1215
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
      Left            =   7200
      TabIndex        =   186
      Top             =   4560
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
      Left            =   5640
      TabIndex        =   185
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox TxtDescripcion 
      Height          =   285
      Left            =   4680
      MaxLength       =   255
      TabIndex        =   176
      Top             =   240
      Width           =   5295
   End
   Begin VB.TextBox Text36 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   123
      Text            =   "0.00"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text35 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   122
      Text            =   "0.00"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text34 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   121
      Text            =   "0.00"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text33 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   120
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text32 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   119
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text31 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   118
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text30 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   117
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text29 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   116
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text28 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   115
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text27 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   114
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text26 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   113
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text25 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   112
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text24 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   111
      Text            =   "0.00"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   110
      Text            =   "0.00"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   109
      Text            =   "0.00"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text21 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   108
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   107
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   106
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   105
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   104
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   103
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   102
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   101
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7200
      TabIndex        =   100
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   99
      Text            =   "0.00"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   98
      Text            =   "0.00"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   97
      Text            =   "0.00"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   96
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   95
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   94
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   93
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   92
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   91
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   90
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   89
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   88
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt13 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt14 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt15 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt16 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt17 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt18 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt19 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt20 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt21 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt22 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0.00"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txt23 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Txt24 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Txt25 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt26 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt27 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt28 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt29 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt30 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt31 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt32 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt33 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt34 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txt35 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Txt36 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Txt2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt6 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt7 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt9 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt10 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txt11 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Txt12 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Txt1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc DtaPeriodos 
      Height          =   330
      Left            =   120
      Top             =   6840
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
   Begin MSAdodcLib.Adodc DtaCuentas 
      Height          =   330
      Left            =   120
      Top             =   7200
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   330
      Left            =   2640
      Top             =   6840
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
      Left            =   2640
      Top             =   7200
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
      Left            =   5280
      Top             =   7200
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
      Left            =   5280
      Top             =   6840
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
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Totales Anuales  Reales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   480
      TabIndex        =   189
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Totales Anuales  Presupuestados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   5640
      TabIndex        =   188
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label LblTotal3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   3600
      TabIndex        =   181
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label LblTotal2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   2040
      TabIndex        =   180
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label LblTotal1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      TabIndex        =   179
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label94 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   3480
      TabIndex        =   178
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label93 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      Height          =   255
      Left            =   240
      TabIndex        =   177
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00800000&
      X1              =   5160
      X2              =   5160
      Y1              =   840
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00800000&
      X1              =   4920
      X2              =   4920
      Y1              =   840
      Y2              =   4560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   5280
      X2              =   9960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   5280
      X2              =   9960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label92 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Presupuesto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5400
      TabIndex        =   175
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Lbl6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A�o 2003"
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
      Height          =   255
      Left            =   8400
      TabIndex        =   174
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Lbl5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A�o 2003"
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
      Height          =   255
      Left            =   6840
      TabIndex        =   173
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Lbl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A�o 2003"
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
      Height          =   255
      Left            =   5280
      TabIndex        =   172
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label88 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   6840
      TabIndex        =   171
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label87 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   6840
      TabIndex        =   170
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label86 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   6840
      TabIndex        =   169
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label85 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   6840
      TabIndex        =   168
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label84 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   6840
      TabIndex        =   167
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label83 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   6840
      TabIndex        =   166
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label82 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   6840
      TabIndex        =   165
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label81 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   6840
      TabIndex        =   164
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label80 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   6840
      TabIndex        =   163
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label79 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   6840
      TabIndex        =   162
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label78 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   6840
      TabIndex        =   161
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   6840
      TabIndex        =   160
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label76 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   159
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label75 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   158
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   157
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   156
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   155
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   154
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   153
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   152
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   151
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   150
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   149
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   7200
      TabIndex        =   148
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   8400
      TabIndex        =   147
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   8400
      TabIndex        =   146
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   8400
      TabIndex        =   145
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   8400
      TabIndex        =   144
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   8400
      TabIndex        =   143
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   8400
      TabIndex        =   142
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   8400
      TabIndex        =   141
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   8400
      TabIndex        =   140
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   8400
      TabIndex        =   139
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   8400
      TabIndex        =   138
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   8400
      TabIndex        =   137
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   8400
      TabIndex        =   136
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   5280
      TabIndex        =   135
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   5280
      TabIndex        =   134
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   5280
      TabIndex        =   133
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   5280
      TabIndex        =   132
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   5280
      TabIndex        =   131
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   5280
      TabIndex        =   130
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   5280
      TabIndex        =   129
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   5280
      TabIndex        =   128
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   5280
      TabIndex        =   127
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   5280
      TabIndex        =   126
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   5280
      TabIndex        =   125
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   5280
      TabIndex        =   124
      Top             =   1560
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   120
      X2              =   4800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   120
      X2              =   4800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      TabIndex        =   87
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A�o 2003"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   86
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A�o 2003"
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
      Height          =   255
      Left            =   1680
      TabIndex        =   85
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A�o 2003"
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
      Height          =   255
      Left            =   120
      TabIndex        =   84
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   3240
      TabIndex        =   83
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   3240
      TabIndex        =   82
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   3240
      TabIndex        =   81
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   3240
      TabIndex        =   80
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   3240
      TabIndex        =   79
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   3240
      TabIndex        =   78
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   3240
      TabIndex        =   77
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   3240
      TabIndex        =   76
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   3240
      TabIndex        =   75
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   3240
      TabIndex        =   74
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   3240
      TabIndex        =   73
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   3240
      TabIndex        =   72
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   71
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   70
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   69
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   68
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   67
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   66
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   65
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   64
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   63
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   62
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   61
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   3600
      TabIndex        =   60
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   1680
      TabIndex        =   59
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label63 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   1680
      TabIndex        =   58
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   1680
      TabIndex        =   57
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   1680
      TabIndex        =   56
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   1680
      TabIndex        =   55
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label67 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   1680
      TabIndex        =   54
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   1680
      TabIndex        =   53
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   1680
      TabIndex        =   52
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   1680
      TabIndex        =   51
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   1680
      TabIndex        =   50
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label72 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   1680
      TabIndex        =   49
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   1680
      TabIndex        =   48
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label98 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label99 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label100 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label101 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label102 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label103 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label104 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label105 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label106 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label107 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label108 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label109 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   11250
      Left            =   0
      Picture         =   "FrmPresupuesto.frx":030A
      Top             =   0
      Width           =   15000
   End
   Begin VB.Label LblTotal6 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   8520
      TabIndex        =   184
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label LblTotal5 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      TabIndex        =   183
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label LblTotal4 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   5520
      TabIndex        =   182
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaCuentas.Recordset.MovePrevious
If DtaCuentas.Recordset.BOF Then
   DtaCuentas.Recordset.MoveNext
   MsgBox "Este es el Primer Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCliente.Text = DtaCuentas.Recordset("CodCuentas")
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub





Private Sub CmdGrabar_Click()
'On Error GoTo TipoErrs
Dim NumeroPeriodo As Integer, Periodo As Integer
Dim Saldo As Double
 
 '//////////////////////////////////////////////////////////////
 '//////////////Datos del primer  periodo////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 1))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = Me.DBCliente.Text
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.DtaPresupuesto.Recordset.AddNew
      Me.DtaPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.DtaPresupuesto.Recordset!CodCuenta = CodigoCuenta
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
  CodigoCuenta = Me.DBCliente.Text
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.DtaPresupuesto.Recordset.AddNew
      Me.DtaPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.DtaPresupuesto.Recordset!CodCuenta = CodigoCuenta
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
  CodigoCuenta = Me.DBCliente.Text
  Me.DtaConsulta.RecordSource = "SELECT Presupuesto.NPeriodo, Presupuesto.CodCuenta, Presupuesto.MontoPresupuestado, Presupuesto.SaldoReal From Presupuesto Where (((Presupuesto.NPeriodo) = " & NumeroPeriodo & " ) And ((Presupuesto.CodCuenta) = '" & CodigoCuenta & "'))"
  Me.DtaConsulta.Refresh
  If Me.DtaConsulta.Recordset.EOF Then
     Me.DtaPresupuesto.Recordset.AddNew
      Me.DtaPresupuesto.Recordset!NPeriodo = NumeroPeriodo
      Me.DtaPresupuesto.Recordset!CodCuenta = CodigoCuenta
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
  Me.DtaPeriodos.Recordset.MoveNext
 Loop
 
 
 
 
 
 
 '//////si existen datos grabo el monto anual////////////
 If Not Me.TxtTotal1.Text = "" Then
 Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 1) And ((PresupuestoAnual.CodigoCuenta) = '" & Me.DBCliente.Text & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If DtaPresupuestoAnual.Recordset.EOF Then
  DtaPresupuestoAnual.Recordset.AddNew
    DtaPresupuestoAnual.Recordset!NumeroTabla = 1
    DtaPresupuestoAnual.Recordset!CodigoCuenta = Me.DBCliente.Text
    DtaPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal1.Text
  DtaPresupuestoAnual.Recordset.Update
 Else
  'DtaPresupuestoAnual.Recordset.Edit
    DtaPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal1.Text
  DtaPresupuestoAnual.Recordset.Update
 End If
 End If
 
  If Not Me.TxtTotal2.Text = "" Then
 Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 2) And ((PresupuestoAnual.CodigoCuenta) = '" & Me.DBCliente.Text & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If DtaPresupuestoAnual.Recordset.EOF Then
  DtaPresupuestoAnual.Recordset.AddNew
    DtaPresupuestoAnual.Recordset!NumeroTabla = 2
    DtaPresupuestoAnual.Recordset!CodigoCuenta = Me.DBCliente.Text
    DtaPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal2.Text
  DtaPresupuestoAnual.Recordset.Update
 Else
  'DtaPresupuestoAnual.Recordset.Edit
    DtaPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal2.Text
  DtaPresupuestoAnual.Recordset.Update
 End If
 End If
 
 If Not Me.TxtTotal3.Text = "" Then
 Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 3) And ((PresupuestoAnual.CodigoCuenta) = '" & Me.DBCliente.Text & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If DtaPresupuestoAnual.Recordset.EOF Then
  DtaPresupuestoAnual.Recordset.AddNew
    DtaPresupuestoAnual.Recordset!NumeroTabla = 3
    DtaPresupuestoAnual.Recordset!CodigoCuenta = Me.DBCliente.Text
    DtaPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal3.Text
  DtaPresupuestoAnual.Recordset.Update
 Else
  'DtaPresupuestoAnual.Recordset.Edit
    DtaPresupuestoAnual.Recordset!MontoAnual = Me.TxtTotal3.Text
  DtaPresupuestoAnual.Recordset.Update
 End If
 End If
 
 
 Exit Sub
TipoErrs:
 MsgBox err.Description
 Exit Sub
End Sub

Private Sub CmdNuevo_Click()
Me.DBCliente.Text = ""
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
Dim Respuesta As Integer
DtaCuentas.Recordset.MoveNext
If DtaCuentas.Recordset.EOF Then
   DtaCuentas.Recordset.MovePrevious
   MsgBox "Este es el Ultimo Registro", vbInfoContabilidadtion, "Control de Cuentas Contabilidad"
Else
  Me.DBCliente.Text = DtaCuentas.Recordset("CodCuentas")
End If
   Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub DBCliente_Change()

On Error GoTo TipoErrs
Dim NumeroPeriodo As Integer, Periodo As Integer, TipoCuenta As String
Dim Saldo As Double, Total1 As Double, Total2 As Double, Total3 As Double, Total4 As Double, Total5 As Double, Total6 As Double
Dim Monto1 As Double, Monto2 As Double, Monto3 As Double, Monto4 As Double, Monto5 As Double, Monto6 As Double
Dim FechaInicio As Date, FechaFin As Date
 Criterio = "CodCuentas='" & Me.DBCliente.Text & "'"
If DtaCuentas.Recordset.RecordCount <> 0 Then DtaCuentas.Recordset.MoveFirst
Me.DtaCuentas.Recordset.Find (Criterio)
If DtaCuentas.Recordset.EOF Then

       

  Me.TxtDescripcion.Text = ""
  Me.LblTotal1.Caption = "0.00"
  Me.LblTotal2.Caption = "0.00"
  Me.LblTotal3.Caption = "0.00"
  Me.LblTotal4.Caption = "0.00"
  Me.LblTotal5.Caption = "0.00"
  Me.LblTotal6.Caption = "0.00"
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
    TipoMoneda = DtaCuentas.Recordset("TipoMoneda")
    TipoCuenta = DtaCuentas.Recordset("TipoCuenta")
   Select Case TipoMoneda
      Case "D�lares"
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

 If Not IsNull(Me.DtaCuentas.Recordset("DescripcionCuentas")) Then
  Me.TxtDescripcion.Text = Me.DtaCuentas.Recordset("DescripcionCuentas")
 End If
 
 
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 1))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  FechaIni = "01/" & Month(Me.DtaPeriodos.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaPeriodos.Recordset("FechaPeriodo"))
  FechaFin = Me.DtaPeriodos.Recordset("FechaPeriodo")
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = Me.DBCliente.Text
'  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito, Transacciones.NPeriodo From Transacciones GROUP BY Transacciones.CodCuentas, Transacciones.NPeriodo HAVING (((Transacciones.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.NPeriodo)=" & NumeroPeriodo & "))"
  Me.DtaConsulta.RecordSource = "SELECT CodCuentas, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito From Transacciones WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "YYYY-MM-DD") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "YYYY-MM-DD") & "', 102)) GROUP BY CodCuentas HAVING (CodCuentas = '" & CodigoCuenta & "') "
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
     Case "D�lares"
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
     Case "C�rdobas"
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
  CodigoCuenta = Me.DBCliente.Text
'  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito, Transacciones.NPeriodo From Transacciones GROUP BY Transacciones.CodCuentas, Transacciones.NPeriodo HAVING (((Transacciones.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.NPeriodo)=" & NumeroPeriodo & "))"
'  Me.DtaConsulta.Refresh
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
     Case "D�lares"
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
     Case "C�rdobas"
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
  CodigoCuenta = Me.DBCliente.Text
'  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Sum(Transacciones.Debito*Transacciones.TCambio) AS MDebito, Sum(Transacciones.TCambio*Transacciones.Credito) AS MCredito, Transacciones.NPeriodo From Transacciones GROUP BY Transacciones.CodCuentas, Transacciones.NPeriodo HAVING (((Transacciones.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.NPeriodo)=" & NumeroPeriodo & "))"
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
     Case "D�lares"
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
  CodigoCuenta = Me.DBCliente.Text
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
 
 Me.LblTotal4.Caption = Format(Total4, "##,##0.00")
 
  '//////////////////////////////////////////////////////////////
 '//////////////Datos del SEGUNDO periodo PRESUPUESTO////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 2))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = Me.DBCliente.Text
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
 Me.LblTotal5.Caption = Format(Total5, "##,##0.00")
 
   '//////////////////////////////////////////////////////////////
 '//////////////Datos del TERCER periodo PRESUPUESTO////////////////////////
 '//////////////////////////////////////////////////////////////
 Me.DtaPeriodos.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.Periodo From Periodos Where (((Periodos.NumeroTabla) = 3))"
 Me.DtaPeriodos.Refresh
 Do While Not Me.DtaPeriodos.Recordset.EOF
  NumeroPeriodo = Me.DtaPeriodos.Recordset!NPeriodo
  Periodo = Me.DtaPeriodos.Recordset!Periodo
  CodigoCuenta = Me.DBCliente.Text
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
 
 Me.LblTotal6.Caption = Format(Total6, "##,##0.00")
 
 Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 1) And ((PresupuestoAnual.CodigoCuenta) = '" & Me.DBCliente.Text & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If Not DtaPresupuestoAnual.Recordset.EOF Then
   Me.TxtTotal1.Text = Format(DtaPresupuestoAnual.Recordset!MontoAnual, "##,##0.00")

 End If
 
  Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 2) And ((PresupuestoAnual.CodigoCuenta) = '" & Me.DBCliente.Text & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If Not DtaPresupuestoAnual.Recordset.EOF Then
   Me.TxtTotal2.Text = Format(DtaPresupuestoAnual.Recordset!MontoAnual, "##,##0.00")

 End If
 
  Me.DtaPresupuestoAnual.RecordSource = "SELECT PresupuestoAnual.NumeroTabla, PresupuestoAnual.CodigoCuenta, PresupuestoAnual.MontoAnual From PresupuestoAnual Where (((PresupuestoAnual.NumeroTabla) = 3) And ((PresupuestoAnual.CodigoCuenta) = '" & Me.DBCliente.Text & "'))"
 Me.DtaPresupuestoAnual.Refresh
 If Not DtaPresupuestoAnual.Recordset.EOF Then
   Me.TxtTotal3.Text = Format(DtaPresupuestoAnual.Recordset!MontoAnual, "##,##0.00")

 End If
 
 
End If 'IF FINAL


Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Activate()
 On Error GoTo TipoErrs
 
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Presupuesto'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdGrabar.Enabled = False
 End If
End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
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
Dim A�o As String, i As Integer

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
   .RecordSource = "Select * from Presupuesto"
   .Refresh
End With

'Me.DtaCuentas.RecordSource = "SELECT Cuentas.TipoMoneda,Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas Where (((Cuentas.TipoCuenta) = 'Gastos' Or (Cuentas.TipoCuenta) = 'Ingresos - Ventas' Or (Cuentas.TipoCuenta) = 'Costos')) ORDER BY Cuentas.CodCuentas"
Me.DtaCuentas.RecordSource = "SELECT Cuentas.TipoMoneda,Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta From Cuentas ORDER BY Cuentas.CodCuentas"
Me.DtaCuentas.Refresh
LlenarDataCombos DtaCuentas, DBCliente, "CodCuentas", "CodCuentas"
Me.DBCliente.ListField = "CodCuentas"
i = 1
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.Periodo, Periodos.FechaPeriodo From Periodos Where (((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3) And ((Periodos.Periodo) = 1)) ORDER BY Periodos.NumeroTabla"
Me.DtaConsulta.Refresh
Do While Not Me.DtaConsulta.Recordset.EOF
 A�o = Year(Me.DtaConsulta.Recordset!FechaPeriodo)
 If i = 1 Then
  Me.Lbl1.Caption = "A�o " & A�o
  Me.Lbl4.Caption = "A�o " & A�o
 ElseIf i = 2 Then
  Me.Lbl2.Caption = "A�o " & A�o
  Me.Lbl5.Caption = "A�o " & A�o
 ElseIf i = 3 Then
  Me.Lbl3.Caption = "A�o " & A�o
  Me.Lbl6.Caption = "A�o " & A�o
 End If
 
 Me.DtaConsulta.Recordset.MoveNext
 i = i + 1
Loop

Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub SmartButton5_Click()
Unload Me
End Sub

Private Sub Text1_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text10_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text11_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text12_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text13_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text14_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text15_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text16_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text17_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text18_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text19_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text2_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text20_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text21_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text22_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text23_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text24_Change()
Me.TxtTotal2.Text = Val(Me.Text13) + Val(Me.Text14) + Val(Me.Text15) + Val(Me.Text16) + Val(Me.Text17) + Val(Me.Text18) + Val(Me.Text19) + Val(Me.Text20) + Val(Me.Text21) + Val(Me.Text22) + Val(Me.Text23) + Val(Me.Text24)
End Sub

Private Sub Text25_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text26_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text27_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text28_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text29_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text3_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text30_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text31_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text32_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text33_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text34_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text35_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text36_Change()
Me.TxtTotal3.Text = Val(Me.Text25) + Val(Me.Text26) + Val(Me.Text27) + Val(Me.Text28) + Val(Me.Text29) + Val(Me.Text30) + Val(Me.Text31) + Val(Me.Text32) + Val(Me.Text33) + Val(Me.Text34) + Val(Me.Text35) + Val(Me.Text36)
End Sub

Private Sub Text4_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text5_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text6_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text7_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text8_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub

Private Sub Text9_Change()
Me.TxtTotal1.Text = Val(Me.Text1) + Val(Me.Text2) + Val(Me.Text3) + Val(Me.Text4) + Val(Me.Text5) + Val(Me.Text6) + Val(Me.Text7) + Val(Me.Text8) + Val(Me.Text9) + Val(Me.Text10) + Val(Me.Text11) + Val(Me.Text12)
End Sub
