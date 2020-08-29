VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmPeriodos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Periodos"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbMoneda 
      Height          =   315
      ItemData        =   "FrmPeriodos.frx":0000
      Left            =   1320
      List            =   "FrmPeriodos.frx":000A
      TabIndex        =   137
      Text            =   "Dólares"
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8040
      TabIndex        =   94
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   120
      TabIndex        =   88
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   6720
      TabIndex        =   93
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   5400
      TabIndex        =   92
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerrar Periodo"
      Height          =   375
      Left            =   3960
      TabIndex        =   91
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton CmdDesBloquear 
      Caption         =   "Desbloquear"
      Height          =   375
      Left            =   2760
      TabIndex        =   90
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CmdBloquear 
      Caption         =   "Bloquear "
      Height          =   375
      Left            =   1440
      TabIndex        =   89
      Top             =   4200
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DtaTransacciones 
      Height          =   375
      Left            =   480
      Top             =   8520
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
      Caption         =   "DtaTransacciones"
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
   Begin MSAdodcLib.Adodc DtaIndices 
      Height          =   375
      Left            =   480
      Top             =   11040
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "DtaIndices"
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
   Begin MSAdodcLib.Adodc DtaTasas 
      Height          =   375
      Left            =   480
      Top             =   10680
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "DtaTasas"
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
      Left            =   480
      Top             =   10320
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   375
      Left            =   480
      Top             =   9960
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc DtaAnSi 
      Height          =   375
      Left            =   480
      Top             =   9600
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaAnSi"
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
   Begin MSAdodcLib.Adodc DtaPeriodos 
      Height          =   375
      Left            =   480
      Top             =   9240
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   "Periodos"
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
   Begin MSAdodcLib.Adodc DtaSaldos 
      Height          =   375
      Left            =   480
      Top             =   8760
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaSaldos"
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
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "01/01/2004"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "01/01/2004"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Txt1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "01/01/2004"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Txt12 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "01/01/2004"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txt11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "01/01/2004"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt10 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "01/01/2004"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "01/01/2004"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt8 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "01/01/2004"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "01/01/2004"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "01/01/2004"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "01/01/2004"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "01/01/2004"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt36 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "01/01/2004"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txt35 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "01/01/2004"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt34 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "01/01/2004"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt33 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "01/01/2004"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt32 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "01/01/2004"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt31 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "01/01/2004"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt30 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "01/01/2004"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt29 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "01/01/2004"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt28 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "01/01/2004"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt27 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "01/01/2004"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt26 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "01/01/2004"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Txt25 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "01/01/2004"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Txt24 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "01/01/2004"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Txt23 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "01/01/2004"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt22 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "01/01/2004"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Txt21 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "01/01/2004"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Txt20 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "01/01/2004"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Txt19 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "01/01/2004"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Txt18 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "01/01/2004"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Txt17 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "01/01/2004"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Txt16 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "01/01/2004"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Txt15 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "01/01/2004"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Txt14 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "01/01/2004"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Txt13 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "01/01/2004"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   -50
      Width           =   9015
      Begin VB.TextBox TxtContracuenta 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker TxtFechaCierre 
         Height          =   285
         Left            =   2160
         TabIndex        =   37
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Format          =   78184449
         CurrentDate     =   37994
      End
      Begin VB.CommandButton CmdBuscaCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8520
         Picture         =   "FrmPeriodos.frx":0021
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmPeriodos.frx":016F
         TabIndex        =   134
         Top             =   360
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label50 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmPeriodos.frx":0209
         TabIndex        =   135
         Top             =   360
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label26 
         Height          =   255
         Left            =   6600
         OleObjectBlob   =   "FrmPeriodos.frx":0293
         TabIndex        =   97
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdCrear 
      Caption         =   "Crear Año"
      Height          =   375
      Left            =   120
      TabIndex        =   95
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton SmartButton5 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   96
      Top             =   4200
      Width           =   1095
   End
   Begin XtremeSuiteControls.ProgressBar BarraCierre 
      Height          =   375
      Left            =   1440
      TabIndex        =   136
      Top             =   4320
      Visible         =   0   'False
      Width           =   6375
      _Version        =   786432
      _ExtentX        =   11245
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin VB.Label LblMoneda 
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda Cierre"
      Height          =   255
      Left            =   120
      TabIndex        =   138
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   85
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Lbl2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   84
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Lbl3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   83
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Lbl4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   82
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Lbl5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   81
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Lbl6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   80
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Lbl7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   79
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Lbl8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   78
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Lbl9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   77
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Lbl10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   76
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Lbl11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   75
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Lbl12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   74
      Top             =   3720
      Width           =   855
   End
   Begin VB.Shape Shape36 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape35 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape34 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape Shape33 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape32 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape31 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape Shape30 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape29 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Shape28 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape27 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape26 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape25 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lbl13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   73
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Lbl14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   72
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Lbl15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   71
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Lbl16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   70
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Lbl17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   69
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Lbl18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   68
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Lbl19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   67
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Lbl20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   66
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Lbl21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   65
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Lbl22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   64
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Lbl23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   63
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Lbl24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   62
      Top             =   3720
      Width           =   855
   End
   Begin VB.Shape Shape24 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape23 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape22 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape Shape21 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape20 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape19 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape14 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lbl36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   61
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Lbl35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   60
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Lbl34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   59
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Lbl33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   58
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Lbl32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   57
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Lbl31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   56
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Lbl30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   55
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Lbl29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   54
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Lbl28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   53
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Lbl27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   52
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Lbl26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   51
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Lbl25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7920
      TabIndex        =   50
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   6240
      TabIndex        =   133
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   6240
      TabIndex        =   132
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   6240
      TabIndex        =   131
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   6240
      TabIndex        =   130
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   6240
      TabIndex        =   129
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   6240
      TabIndex        =   128
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   6240
      TabIndex        =   127
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   6240
      TabIndex        =   126
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   6240
      TabIndex        =   125
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   6240
      TabIndex        =   124
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   6240
      TabIndex        =   123
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   6240
      TabIndex        =   122
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   3120
      TabIndex        =   121
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label63 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   3120
      TabIndex        =   120
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   3120
      TabIndex        =   119
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   3120
      TabIndex        =   118
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   3120
      TabIndex        =   117
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label67 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   3120
      TabIndex        =   116
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   3120
      TabIndex        =   115
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   3120
      TabIndex        =   114
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   3120
      TabIndex        =   113
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   3120
      TabIndex        =   112
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label72 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   3120
      TabIndex        =   111
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   3120
      TabIndex        =   110
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label98 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      Height          =   255
      Left            =   120
      TabIndex        =   109
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label99 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      Height          =   255
      Left            =   120
      TabIndex        =   108
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label100 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   120
      TabIndex        =   107
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label101 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Left            =   120
      TabIndex        =   106
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label102 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   120
      TabIndex        =   105
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label103 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Left            =   120
      TabIndex        =   104
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label104 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Left            =   120
      TabIndex        =   103
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label105 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   120
      TabIndex        =   102
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label106 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Left            =   120
      TabIndex        =   101
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label107 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Left            =   120
      TabIndex        =   100
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label108 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   120
      TabIndex        =   99
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label109 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   120
      MouseIcon       =   "FrmPeriodos.frx":030B
      MousePointer    =   4  'Icon
      TabIndex        =   98
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   11250
      Left            =   -120
      Picture         =   "FrmPeriodos.frx":217D
      Top             =   -5880
      Width           =   15000
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   49
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   48
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   47
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   46
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   45
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   44
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   43
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   42
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   41
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   40
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   39
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2004"
      Height          =   255
      Left            =   6720
      TabIndex        =   38
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPeriodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

End Sub

Private Sub Label97_Click()
 Me.Label109.BorderStyle = 0
End Sub

Private Sub Label87_Click()
End Sub

Private Sub Label58_Click()
End Sub

Private Sub Label56_Click()
End Sub

Private Sub SmartButton2_Click()

End Sub

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
'///////////Busco los Datos del Primer Periodo/////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
If Not DtaConsulta.Recordset.EOF Then
 NPeriodo = Val(DtaConsulta.Recordset("NPeriodo")) - 12
 NPeriodoFin = NPeriodo + 11
End If

 Me.DtaAnSi.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NPeriodo) = " & NPeriodo & " )) ORDER BY Periodos.NPeriodo"
 Me.DtaAnSi.Refresh
If Not DtaAnSi.Recordset.EOF Then

 '//////////Edito el Periodo 3 y se convierte en el periodo 0/////////
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
  DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 0
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop

 '//////////Edito el Periodo 2 y se convierte en el tercer año/////////
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 2)) ORDER BY Periodos.NPeriodo"
  DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 3
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop

 '//////Edito los registros del Primero año /////////////////
 '///////se convierte en el segundo año/////////////////////
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
 DtaConsulta.Refresh
    Do While Not DtaConsulta.Recordset.EOF
        'DtaConsulta.Recordset.Edit
           DtaConsulta.Recordset("NumeroTabla") = 2
        DtaConsulta.Recordset.Update
      DtaConsulta.Recordset.MoveNext
    Loop
 

 
  
 '//////////Edito el Periodo Cero en el 1//////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos WHERE (((Periodos.NPeriodo) Between " & NPeriodo & " And " & NPeriodoFin & " )) ORDER BY Periodos.NPeriodo"
   Me.DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 1
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop
  
Else
 MsgBox "Este es el Primer Registro", vbInformation, "sistema Contable"
 Exit Sub
   
End If

'/////////Actulizo los cambios efectuados en los periodos/////////
  '////LLeno los datos del primer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
If Not DtaConsulta.Recordset.EOF Then
 Me.TxtFechaCierre.Value = DtaConsulta.Recordset("FechaPeriodo")
End If
Do While Not DtaConsulta.Recordset.EOF

Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
 

 End Select


Select Case i
     Case 1: Me.Txt1.Text = Fecha
             Me.Lbl1.Caption = NTransacciones
             Me.Shape1.BackColor = Color
     Case 2: Me.Txt2.Text = Fecha
             Me.Lbl2.Caption = NTransacciones
             Me.Shape2.BackColor = Color
     Case 3: Me.Txt3.Text = Fecha
             Me.Lbl3.Caption = NTransacciones
             Me.Shape3.BackColor = Color
     Case 4: Me.Txt4.Text = Fecha
             Me.Lbl4.Caption = NTransacciones
             Me.Shape4.BackColor = Color
     Case 5: Me.Txt5.Text = Fecha
             Me.Lbl5.Caption = NTransacciones
             Me.Shape5.BackColor = Color
     Case 6: Me.Txt6.Text = Fecha
             Me.Lbl6.Caption = NTransacciones
             Me.Shape6.BackColor = Color
     Case 7: Me.Txt7.Text = Fecha
             Me.Lbl7.Caption = NTransacciones
             Me.Shape7.BackColor = Color
     Case 8: Me.Txt8.Text = Fecha
             Me.Lbl8.Caption = NTransacciones
             Me.Shape8.BackColor = Color
     Case 9: Me.Txt9.Text = Fecha
             Me.Lbl9.Caption = NTransacciones
             Me.Shape9.BackColor = Color
     Case 10: Me.Txt10.Text = Fecha
             Me.Lbl10.Caption = NTransacciones
             Me.Shape10.BackColor = Color
     Case 11: Me.Txt11.Text = Fecha
             Me.Lbl11.Caption = NTransacciones
             Me.Shape11.BackColor = Color
     Case 12: Me.Txt12.Text = Fecha
             Me.Lbl12.Caption = NTransacciones
             Me.Shape12.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del segundo año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 2)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
 End Select
Select Case i
     Case 1: Me.Txt13.Text = Fecha
             Me.Lbl13.Caption = NTransacciones
             Me.Shape13.BackColor = Color
     Case 2: Me.Txt14.Text = Fecha
             Me.Lbl14.Caption = NTransacciones
             Me.Shape14.BackColor = Color
     Case 3: Me.Txt15.Text = Fecha
             Me.Lbl15.Caption = NTransacciones
             Me.Shape15.BackColor = Color
     Case 4: Me.Txt16.Text = Fecha
             Me.Lbl16.Caption = NTransacciones
             Me.Shape16.BackColor = Color
     Case 5: Me.Txt17.Text = Fecha
             Me.Lbl17.Caption = NTransacciones
             Me.Shape17.BackColor = Color
     Case 6: Me.Txt18.Text = Fecha
             Me.Lbl18.Caption = NTransacciones
             Me.Shape18.BackColor = Color
     Case 7: Me.Txt19.Text = Fecha
             Me.Lbl19.Caption = NTransacciones
             Me.Shape19.BackColor = Color
     Case 8: Me.Txt20.Text = Fecha
             Me.Lbl20.Caption = NTransacciones
             Me.Shape20.BackColor = Color
     Case 9: Me.Txt21.Text = Fecha
             Me.Lbl21.Caption = NTransacciones
             Me.Shape21.BackColor = Color
     Case 10: Me.Txt22.Text = Fecha
             Me.Lbl22.Caption = NTransacciones
             Me.Shape22.BackColor = Color
     Case 11: Me.Txt23.Text = Fecha
              Me.Lbl23.Caption = NTransacciones
              Me.Shape23.BackColor = Color
     Case 12: Me.Txt24.Text = Fecha
             Me.Lbl24.Caption = NTransacciones
             Me.Shape24.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del tercer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))

estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
  End Select


Select Case i
     Case 1: Me.Txt25.Text = Fecha
             Me.Lbl25.Caption = NTransacciones
             Me.Shape25.BackColor = Color
     Case 2: Me.Txt26.Text = Fecha
             Me.Lbl26.Caption = NTransacciones
             Me.Shape26.BackColor = Color
     Case 3: Me.Txt27.Text = Fecha
             Me.Lbl27.Caption = NTransacciones
             Me.Shape27.BackColor = Color
     Case 4: Me.Txt28.Text = Fecha
             Me.Lbl28.Caption = NTransacciones
             Me.Shape28.BackColor = Color
     Case 5: Me.Txt29.Text = Fecha
             Me.Lbl29.Caption = NTransacciones
             Me.Shape29.BackColor = Color
     Case 6: Me.Txt30.Text = Fecha
             Me.Lbl30.Caption = NTransacciones
             Me.Shape30.BackColor = Color
     Case 7: Me.Txt31.Text = Fecha
             Me.Lbl31.Caption = NTransacciones
             Me.Shape31.BackColor = Color
     Case 8: Me.Txt32.Text = Fecha
             Me.Lbl32.Caption = NTransacciones
             Me.Shape32.BackColor = Color
     Case 9: Me.Txt33.Text = Fecha
             Me.Lbl33.Caption = NTransacciones
             Me.Shape33.BackColor = Color
     Case 10: Me.Txt34.Text = Fecha
             Me.Lbl34.Caption = NTransacciones
             Me.Shape34.BackColor = Color
     Case 11: Me.Txt35.Text = Fecha
             Me.Lbl35.Caption = NTransacciones
             Me.Shape35.BackColor = Color
     Case 12: Me.Txt36.Text = Fecha
             Me.Lbl36.Caption = NTransacciones
             Me.Shape36.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop


Exit Sub
TipoErrs:
MsgBox err.Description

End Sub

Private Sub CmdBloquear_Click()
Dim NumFecha1 As Long, Fechas As String
 Select Case Opt
   Case "1": Me.Shape1.BackColor = &HC0FFFF
             Me.Txt1.SetFocus
             NumFecha1 = CDate(Me.Txt1.Text)
   Case "2": Me.Shape2.BackColor = &HC0FFFF
             Me.Txt2.SetFocus
             NumFecha1 = CDate(Me.Txt2.Text)
   Case "3": Me.Shape3.BackColor = &HC0FFFF
             Me.Txt3.SetFocus
             NumFecha1 = CDate(Me.Txt3.Text)
    Case "4": Me.Shape4.BackColor = &HC0FFFF
             Me.Txt4.SetFocus
             NumFecha1 = CDate(Me.Txt4.Text)
    Case "5": Me.Shape5.BackColor = &HC0FFFF
             Me.Txt5.SetFocus
             NumFecha1 = CDate(Me.Txt5.Text)
    Case "6": Me.Shape6.BackColor = &HC0FFFF
             Me.Txt6.SetFocus
             NumFecha1 = CDate(Me.Txt6.Text)
   Case "7": Me.Shape7.BackColor = &HC0FFFF
             Me.Txt7.SetFocus
             NumFecha1 = CDate(Me.Txt7.Text)
   Case "8": Me.Shape8.BackColor = &HC0FFFF
             Me.Txt8.SetFocus
             NumFecha1 = CDate(Me.Txt8.Text)
   Case "9": Me.Shape9.BackColor = &HC0FFFF
             Me.Txt9.SetFocus
             NumFecha1 = CDate(Me.Txt9.Text)
  Case "10": Me.Shape10.BackColor = &HC0FFFF
             Me.Txt10.SetFocus
             NumFecha1 = CDate(Me.Txt10.Text)
  Case "11": Me.Shape11.BackColor = &HC0FFFF
             Me.Txt11.SetFocus
             NumFecha1 = CDate(Me.Txt11.Text)
  Case "12": Me.Shape12.BackColor = &HC0FFFF
             Me.Txt12.SetFocus
             NumFecha1 = CDate(Me.Txt12.Text)
  Case "13": Me.Shape13.BackColor = &HC0FFFF
             Me.Txt13.SetFocus
             NumFecha1 = CDate(Me.Txt13.Text)
  Case "14": Me.Shape14.BackColor = &HC0FFFF
             Me.Txt14.SetFocus
             NumFecha1 = CDate(Me.Txt14.Text)
  Case "15": Me.Shape15.BackColor = &HC0FFFF
             Me.Txt15.SetFocus
             NumFecha1 = CDate(Me.Txt15.Text)
  Case "16": Me.Shape16.BackColor = &HC0FFFF
             Me.Txt16.SetFocus
             NumFecha1 = CDate(Me.Txt16.Text)
  Case "17": Me.Shape17.BackColor = &HC0FFFF
             Me.Txt17.SetFocus
             NumFecha1 = CDate(Me.Txt17.Text)
  Case "18": Me.Shape18.BackColor = &HC0FFFF
             Me.Txt18.SetFocus
             NumFecha1 = CDate(Me.Txt18.Text)
  Case "19": Me.Shape19.BackColor = &HC0FFFF
             Me.Txt19.SetFocus
             NumFecha1 = CDate(Me.Txt19.Text)
  Case "20": Me.Shape20.BackColor = &HC0FFFF
             Me.Txt20.SetFocus
             NumFecha1 = CDate(Me.Txt20.Text)
  Case "21": Me.Shape21.BackColor = &HC0FFFF
             Me.Txt21.SetFocus
             NumFecha1 = CDate(Me.Txt21.Text)
  Case "22": Me.Shape22.BackColor = &HC0FFFF
             Me.Txt22.SetFocus
             NumFecha1 = CDate(Me.Txt22.Text)
  Case "23": Me.Shape23.BackColor = &HC0FFFF
             Me.Txt23.SetFocus
             NumFecha1 = CDate(Me.Txt23.Text)
    Case "24": Me.Shape24.BackColor = &HC0FFFF
             Me.Txt24.SetFocus
             NumFecha1 = CDate(Me.Txt24.Text)
   Case "25": Me.Shape25.BackColor = &HC0FFFF
             Me.Txt25.SetFocus
             NumFecha1 = CDate(Me.Txt25.Text)
   Case "26": Me.Shape26.BackColor = &HC0FFFF
             Me.Txt26.SetFocus
             NumFecha1 = CDate(Me.Txt26.Text)
    Case "27": Me.Shape27.BackColor = &HC0FFFF
             Me.Txt27.SetFocus
             NumFecha1 = CDate(Me.Txt27.Text)
    Case "28": Me.Shape28.BackColor = &HC0FFFF
             Me.Txt28.SetFocus
             NumFecha1 = CDate(Me.Txt28.Text)
    Case "29": Me.Shape29.BackColor = &HC0FFFF
             Me.Txt29.SetFocus
             NumFecha1 = CDate(Me.Txt29.Text)
    Case "30": Me.Shape30.BackColor = &HC0FFFF
             Me.Txt30.SetFocus
             NumFecha1 = CDate(Me.Txt30.Text)
    Case "31": Me.Shape31.BackColor = &HC0FFFF
             Me.Txt31.SetFocus
             NumFecha1 = CDate(Me.Txt31.Text)
     Case "32": Me.Shape32.BackColor = &HC0FFFF
             Me.Txt32.SetFocus
             NumFecha1 = CDate(Me.Txt32.Text)
   Case "33": Me.Shape33.BackColor = &HC0FFFF
             Me.Txt33.SetFocus
             NumFecha1 = CDate(Me.Txt33.Text)
             
    Case "34": Me.Shape34.BackColor = &HC0FFFF
             Me.Txt34.SetFocus
             NumFecha1 = CDate(Me.Txt34.Text)
    Case "35": Me.Shape35.BackColor = &HC0FFFF
             Me.Txt35.SetFocus
             NumFecha1 = CDate(Me.Txt35.Text)
     Case "36": Me.Shape36.BackColor = &HC0FFFF
             Me.Txt36.SetFocus
             NumFecha1 = CDate(Me.Txt36.Text)
             
             
 End Select
 
Fechas = CDate(NumFecha1)
Fechas = Format(Fechas, "yyyy/mm/dd")
 
 Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones From Periodos WHERE     (FechaPeriodo = CONVERT(DATETIME, '" & Fechas & "', 102)) ORDER BY NPeriodo"
 'Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.FechaPeriodo) = " & NumFecha1 & " )) ORDER BY Periodos.NPeriodo"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not DtaConsulta.Recordset("EstadoPeriodo") = "C" Then
   'Me.'DtaConsulta.Recordset.Edit
    DtaConsulta.Recordset("EstadoPeriodo") = "B"
   Me.DtaConsulta.Recordset.Update
  Else
    MsgBox "No se puede Bloquear un Periodo Cerrado", vbCritical, "sistema Contable"
    Select Case Opt
   Case "1": Me.Shape1.BackColor = &HFF&
   Case "2": Me.Shape2.BackColor = &HFF&
   Case "3": Me.Shape3.BackColor = &HFF&
   Case "4": Me.Shape4.BackColor = &HFF&
   Case "5": Me.Shape5.BackColor = &HFF&
   Case "6": Me.Shape6.BackColor = &HFF&
   Case "7": Me.Shape7.BackColor = &HFF&
   Case "8": Me.Shape8.BackColor = &HFF&
   Case "9": Me.Shape9.BackColor = &HFF&
   Case "10": Me.Shape10.BackColor = &HFF&
   Case "11": Me.Shape11.BackColor = &HFF&
   Case "12": Me.Shape12.BackColor = &HFF&
   Case "13": Me.Shape13.BackColor = &HFF&
   Case "14": Me.Shape14.BackColor = &HFF&
   Case "15": Me.Shape15.BackColor = &HFF&
   Case "16": Me.Shape16.BackColor = &HFF&
   Case "17": Me.Shape17.BackColor = &HFF&
   Case "18": Me.Shape18.BackColor = &HFF&
   Case "19": Me.Shape19.BackColor = &HFF&
   Case "20": Me.Shape20.BackColor = &HFF&
   Case "21": Me.Shape21.BackColor = &HFF&
   Case "22": Me.Shape22.BackColor = &HFF&
   Case "23": Me.Shape23.BackColor = &HFF&
   Case "24": Me.Shape24.BackColor = &HFF&
   Case "25": Me.Shape25.BackColor = &HFF&
   Case "26": Me.Shape26.BackColor = &HFF&
   Case "27": Me.Shape27.BackColor = &HFF&
   Case "28": Me.Shape28.BackColor = &HFF&
   Case "29": Me.Shape29.BackColor = &HFF&
   Case "30": Me.Shape30.BackColor = &HFF&
   Case "31": Me.Shape31.BackColor = &HFF&
   Case "32": Me.Shape32.BackColor = &HFF&
   Case "33": Me.Shape33.BackColor = &HFF&
   Case "34": Me.Shape34.BackColor = &HFF&
   Case "35": Me.Shape35.BackColor = &HFF&
   Case "36": Me.Shape36.BackColor = &HFF&
  End Select
  End If
 
 End If
 
End Sub

Private Sub SmartButton3_Click()

End Sub

Private Sub CmdBuscaCuenta_Click()
QueProducto = "Periodo"
FrmConsulta.Show 1
End Sub

Private Sub CmdCancelar_Click()
     Me.CmdAnterior.Visible = True
     Me.CmdBloquear.Visible = True
     Me.CmdCerrar.Visible = True
     Me.CmdCrear.Visible = True
     Me.CmdDesBloquear.Visible = True
     Me.CmdSiguiente.Visible = True
     Me.Label26.Visible = False
     Me.TxtContracuenta.Visible = False
     Me.CmdBuscaCuenta.Visible = False
     Me.SmartButton5.Visible = True
     Me.CmdProcesar.Visible = False
     Me.CmdCancelar.Visible = False
     Me.BarraCierre.Visible = False
Salir = True
End Sub

Private Sub CmdCerrar_Click()
Dim CadenaSQL As String
  Respuesta = MsgBox("Esta seguro de Cerrar el Periodo?", vbYesNo, "Cerrando el Periodo: " & Opt)
   If Not Respuesta = 6 Then
    Exit Sub
   End If

 Select Case Opt
   Case "1": Me.Shape1.BackColor = &HFF&
             Me.Txt1.SetFocus
             NumFecha1 = CDate(Me.Txt1.Text)
   Case "2": Me.Shape2.BackColor = &HFF&
             Me.Txt2.SetFocus
             NumFecha1 = CDate(Me.Txt2.Text)
   Case "3": Me.Shape3.BackColor = &HFF&
             Me.Txt3.SetFocus
             NumFecha1 = CDate(Me.Txt3.Text)
    Case "4": Me.Shape4.BackColor = &HFF&
             Me.Txt4.SetFocus
             NumFecha1 = CDate(Me.Txt4.Text)
    Case "5": Me.Shape5.BackColor = &HFF&
             Me.Txt5.SetFocus
             NumFecha1 = CDate(Me.Txt5.Text)
    Case "6": Me.Shape6.BackColor = &HFF&
             Me.Txt6.SetFocus
             NumFecha1 = CDate(Me.Txt6.Text)
   Case "7": Me.Shape7.BackColor = &HFF&
             Me.Txt7.SetFocus
             NumFecha1 = CDate(Me.Txt7.Text)
   Case "8": Me.Shape8.BackColor = &HFF&
             Me.Txt8.SetFocus
             NumFecha1 = CDate(Me.Txt8.Text)
   Case "9": Me.Shape9.BackColor = &HFF&
             Me.Txt9.SetFocus
             NumFecha1 = CDate(Me.Txt9.Text)
  Case "10": Me.Shape10.BackColor = &HFF&
             Me.Txt10.SetFocus
             NumFecha1 = CDate(Me.Txt10.Text)
  Case "11": Me.Shape11.BackColor = &HFF&
             Me.Txt11.SetFocus
             NumFecha1 = CDate(Me.Txt11.Text)
  Case "12": 'Me.Shape12.BackColor = &HFF&
             Me.Txt12.SetFocus
             NumFecha1 = CDate(Me.Txt12.Text)
                     
  Case "13": Me.Shape13.BackColor = &HFF&
             Me.Txt13.SetFocus
             NumFecha1 = CDate(Me.Txt13.Text)
  Case "14": Me.Shape14.BackColor = &HFF&
             Me.Txt14.SetFocus
             NumFecha1 = CDate(Me.Txt14.Text)
  Case "15": Me.Shape15.BackColor = &HFF&
             Me.Txt15.SetFocus
             NumFecha1 = CDate(Me.Txt15.Text)
  Case "16": Me.Shape16.BackColor = &HFF&
             Me.Txt16.SetFocus
             NumFecha1 = CDate(Me.Txt16.Text)
  Case "17": Me.Shape17.BackColor = &HFF&
             Me.Txt17.SetFocus
             NumFecha1 = CDate(Me.Txt17.Text)
  Case "18": Me.Shape18.BackColor = &HFF&
             Me.Txt18.SetFocus
             NumFecha1 = CDate(Me.Txt18.Text)
  Case "19": Me.Shape19.BackColor = &HFF&
             Me.Txt19.SetFocus
             NumFecha1 = CDate(Me.Txt19.Text)
  Case "20": Me.Shape20.BackColor = &HFF&
             Me.Txt20.SetFocus
             NumFecha1 = CDate(Me.Txt20.Text)
  Case "21": Me.Shape21.BackColor = &HFF&
             Me.Txt21.SetFocus
             NumFecha1 = CDate(Me.Txt21.Text)
  Case "22": Me.Shape22.BackColor = &HFF&
             Me.Txt22.SetFocus
             NumFecha1 = CDate(Me.Txt22.Text)
  Case "23": Me.Shape23.BackColor = &HFF&
             Me.Txt23.SetFocus
             NumFecha1 = CDate(Me.Txt23.Text)
    Case "24": 'Me.Shape24.BackColor = &HFF&
             Me.Txt24.SetFocus
             NumFecha1 = CDate(Me.Txt24.Text)
   Case "25": Me.Shape25.BackColor = &HFF&
             Me.Txt25.SetFocus
             NumFecha1 = CDate(Me.Txt25.Text)
   Case "26": Me.Shape26.BackColor = &HFF&
             Me.Txt26.SetFocus
             NumFecha1 = CDate(Me.Txt26.Text)
    Case "27": Me.Shape27.BackColor = &HFF&
             Me.Txt27.SetFocus
             NumFecha1 = CDate(Me.Txt27.Text)
    Case "28": Me.Shape28.BackColor = &HFF&
             Me.Txt28.SetFocus
             NumFecha1 = CDate(Me.Txt28.Text)
    Case "29": Me.Shape29.BackColor = &HFF&
             Me.Txt29.SetFocus
             NumFecha1 = CDate(Me.Txt29.Text)
    Case "30": Me.Shape30.BackColor = &HFF&
             Me.Txt30.SetFocus
             NumFecha1 = CDate(Me.Txt30.Text)
    Case "31": Me.Shape31.BackColor = &HFF&
             Me.Txt31.SetFocus
             NumFecha1 = CDate(Me.Txt31.Text)
     Case "32": Me.Shape32.BackColor = &HFF&
             Me.Txt32.SetFocus
             NumFecha1 = CDate(Me.Txt32.Text)
   Case "33": Me.Shape33.BackColor = &HFF&
             Me.Txt33.SetFocus
             NumFecha1 = CDate(Me.Txt33.Text)
             
    Case "34": Me.Shape34.BackColor = &HFF&
             Me.Txt34.SetFocus
             NumFecha1 = CDate(Me.Txt34.Text)
    Case "35": Me.Shape35.BackColor = &HFF&
             Me.Txt35.SetFocus
             NumFecha1 = CDate(Me.Txt35.Text)
     Case "36": 'Me.Shape36.BackColor = &HFF&
             Me.Txt36.SetFocus
             NumFecha1 = CDate(Me.Txt36.Text)
             
             
 End Select
 Fechas = CDate(NumFecha1)
 Fechas = Format(Fechas, "yyyy/mm/dd")
 Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones From Periodos  WHERE     (FechaPeriodo = CONVERT(DATETIME, '" & Fechas & "', 102))ORDER BY NPeriodo"
' Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.FechaPeriodo) = " & NumFecha1 & " )) ORDER BY Periodos.NPeriodo"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  NumeroPeriodo = Me.DtaConsulta.Recordset("NPeriodo")
  NtransaccionPeriodo = Me.DtaConsulta.Recordset("NTransacciones")
  NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
  Fechas2 = Me.DtaConsulta.Recordset("FechaPeriodo")
  Me.DtaTasas.RecordSource = "SELECT FechaTasas, MontoCordobas, MontoLibras From Tasas WHERE     (FechaTasas = CONVERT(DATETIME, '" & Fechas & "', 102))"
  Me.DtaTasas.Refresh
  If Me.DtaTasas.Recordset.EOF Then
    MsgBox "Se necesita la tasa de Cambio del  " & Format(Fechas2, "dd/mm/yyyy"), vbCritical, "Sistema Contable"
    Unload Me
    Exit Sub
  Else
    TasaCambioCordobas = 1
    TasaCambioDolares = Me.DtaTasas.Recordset("MontoCordobas")
    TasaCambioEuro = Me.DtaTasas.Recordset("MontoCordobas")
  End If
  If DtaConsulta.Recordset("EstadoPeriodo") = "C" Then
       MsgBox "No se Puede Cerrar un Periodo Cerrado", vbCritical, "sistema Contable"
  ElseIf Not DtaConsulta.Recordset("EstadoPeriodo") = "B" Then
    
  Select Case Opt
   Case "12"
     
     NumeroPeriodo = NumeroPeriodo - 1
     Criterio = "NPeriodo=" & NumeroPeriodo & " "
     Me.DtaPeriodos.Recordset.Find (Criterio)
     If Not Me.DtaPeriodos.Recordset.EOF Then
     If Me.DtaPeriodos.Recordset("EstadoPeriodo") = "C" Then
     Me.CmdAnterior.Visible = False
     Me.CmdBloquear.Visible = False
     Me.CmdCerrar.Visible = False
     Me.CmdCrear.Visible = False
     Me.CmdDesBloquear.Visible = False
     Me.CmdSiguiente.Visible = False
     Me.Label26.Visible = True
     Me.TxtContracuenta.Visible = True
     Me.CmdBuscaCuenta.Visible = True
     Me.LblMoneda.Visible = True
     Me.CmbMoneda.Visible = True
     Me.SmartButton5.Visible = False
     Me.CmdProcesar.Visible = True
     Me.CmdCancelar.Visible = True
     Me.BarraCierre.Visible = True
     NumeroPeriodo = NumeroPeriodo + 1
     Salir = False
     Else
        MsgBox "Para Cerrar un Periodo, Necesita Cerrar el Periodo Anterior", vbCritical, "sistema Contable"
        Salir = True
     End If

     End If
   Case "24"
     NumeroPeriodo = NumeroPeriodo - 1
     Criterio = "NPeriodo=" & NumeroPeriodo & " "
     Me.DtaPeriodos.Recordset.Find (Criterio)
     If Not Me.DtaPeriodos.Recordset.EOF Then
     If Me.DtaPeriodos.Recordset("EstadoPeriodo") = "C" Then
     Me.CmdAnterior.Visible = False
     Me.CmdBloquear.Visible = False
     Me.CmdCerrar.Visible = False
     Me.CmdCrear.Visible = False
     Me.CmdDesBloquear.Visible = False
     Me.CmdSiguiente.Visible = False
     Me.Label26.Visible = True
     Me.TxtContracuenta.Visible = True
     Me.CmdBuscaCuenta.Visible = True
     Me.LblMoneda.Visible = True
     Me.CmbMoneda.Visible = True
     Me.CmdProcesar.Visible = True
     Me.CmdCancelar.Visible = True
     Me.BarraCierre.Visible = True
     NumeroPeriodo = NumeroPeriodo + 1
     Salir = False
     Else
        MsgBox "Para Cerrar un Periodo, Necesita Cerrar el Periodo Anterior", vbCritical, "sistema Contable"
        Salir = True
     End If

     End If
   Case "36"
     NumeroPeriodo = NumeroPeriodo - 1
     Criterio = "NPeriodo=" & NumeroPeriodo & " "
     Me.DtaPeriodos.Recordset.Find (Criterio)
     If Not Me.DtaPeriodos.Recordset.EOF Then
     If Me.DtaPeriodos.Recordset("EstadoPeriodo") = "C" Then
     Me.CmdAnterior.Visible = False
     Me.CmdBloquear.Visible = False
     Me.CmdCerrar.Visible = False
     Me.CmdCrear.Visible = False
     Me.CmdDesBloquear.Visible = False
     Me.CmdSiguiente.Visible = False
     Me.Label26.Visible = True
     Me.TxtContracuenta.Visible = True
     Me.CmdBuscaCuenta.Visible = True
     Me.LblMoneda.Visible = True
     Me.CmbMoneda.Visible = True
     Me.CmdProcesar.Visible = True
     Me.CmdCancelar.Visible = True
     Me.BarraCierre.Visible = True
     NumeroPeriodo = NumeroPeriodo + 1
     Salir = False
     Else
        MsgBox "Para Cerrar un Periodo, Necesita Cerrar el Periodo Anterior", vbCritical, "sistema Contable"
        Salir = True
     End If
     
     End If
   Case Else
        NumeroPeriodo = NumeroPeriodo - 1
        Criterio = "NPeriodo=" & NumeroPeriodo & " "
        Me.DtaPeriodos.Recordset.Find (Criterio)
    If Not Me.DtaPeriodos.Recordset.EOF Then
       If Me.DtaPeriodos.Recordset("EstadoPeriodo") = "C" Then
        'Me.'DtaConsulta.Recordset.Edit
        DtaConsulta.Recordset("EstadoPeriodo") = "C"
        Me.DtaConsulta.Recordset.Update
        Salir = True
       Else
        Me.DtaPeriodos.Recordset.MoveFirst
        NumeroPeriodo = NumeroPeriodo + 1
        If Me.DtaPeriodos.Recordset("NPeriodo") = NumeroPeriodo Then
        'Me.'DtaConsulta.Recordset.Edit
        DtaConsulta.Recordset("EstadoPeriodo") = "C"
        Me.DtaConsulta.Recordset.Update
        Salir = True
        
        Else
        MsgBox "Para Cerrar un Periodo, Necesita Cerrar el Periodo Anterior", vbCritical, "sistema Contable"
        Select Case Opt
        Case "1": Me.Shape1.BackColor = &HC000&
        Case "2": Me.Shape2.BackColor = &HC000&
        Case "3": Me.Shape3.BackColor = &HC000&
        Case "4": Me.Shape4.BackColor = &HC000&
        Case "5": Me.Shape5.BackColor = &HC000&
        Case "6": Me.Shape6.BackColor = &HC000&
        Case "7": Me.Shape7.BackColor = &HC000&
        Case "8": Me.Shape8.BackColor = &HC000&
        Case "9": Me.Shape9.BackColor = &HC000&
        Case "10": Me.Shape10.BackColor = &HC000&
        Case "11": Me.Shape11.BackColor = &HC000&
        Case "12": Me.Shape12.BackColor = &HC000&
        Case "13": Me.Shape13.BackColor = &HC000&
        Case "14": Me.Shape14.BackColor = &HC000&
        Case "15": Me.Shape15.BackColor = &HC000&
        Case "16": Me.Shape16.BackColor = &HC000&
        Case "17": Me.Shape17.BackColor = &HC000&
        Case "18": Me.Shape18.BackColor = &HC000&
        Case "19": Me.Shape19.BackColor = &HC000&
        Case "20": Me.Shape20.BackColor = &HC000&
        Case "21": Me.Shape21.BackColor = &HC000&
        Case "22": Me.Shape22.BackColor = &HC000&
        Case "23": Me.Shape23.BackColor = &HC000&
        Case "24": Me.Shape24.BackColor = &HC000&
        Case "25": Me.Shape25.BackColor = &HC000&
        Case "26": Me.Shape26.BackColor = &HC000&
        Case "27": Me.Shape27.BackColor = &HC000&
        Case "28": Me.Shape28.BackColor = &HC000&
        Case "29": Me.Shape29.BackColor = &HC000&
        Case "30": Me.Shape30.BackColor = &HC000&
        Case "31": Me.Shape31.BackColor = &HC000&
        Case "32": Me.Shape32.BackColor = &HC000&
        Case "33": Me.Shape33.BackColor = &HC000&
        Case "34": Me.Shape34.BackColor = &HC000&
        Case "35": Me.Shape35.BackColor = &HC000&
        Case "36": Me.Shape36.BackColor = &HC000&
        End Select
        End If
      End If
    Else
       Me.DtaPeriodos.Recordset.MoveFirst
        'Me.'DtaConsulta.Recordset.Edit
        DtaConsulta.Recordset("EstadoPeriodo") = "C"
        Me.DtaConsulta.Recordset.Update
        Salir = True
     
    End If

   End Select
   


 Else
   MsgBox "No se Puede Cerrar un Periodo Bloqueado", vbCritical, "sistema Contable"
    Select Case Opt
   Case "1": Me.Shape1.BackColor = &HC0FFFF
   Case "2": Me.Shape2.BackColor = &HC0FFFF
   Case "3": Me.Shape3.BackColor = &HC0FFFF
   Case "4": Me.Shape4.BackColor = &HC0FFFF
   Case "5": Me.Shape5.BackColor = &HC0FFFF
   Case "6": Me.Shape6.BackColor = &HC0FFFF
   Case "7": Me.Shape7.BackColor = &HC0FFFF
   Case "8": Me.Shape8.BackColor = &HC0FFFF
   Case "9": Me.Shape9.BackColor = &HC0FFFF
   Case "10": Me.Shape10.BackColor = &HC0FFFF
   Case "11": Me.Shape11.BackColor = &HC0FFFF
   Case "12": Me.Shape12.BackColor = &HC0FFFF
   Case "13": Me.Shape13.BackColor = &HC0FFFF
   Case "14": Me.Shape14.BackColor = &HC0FFFF
   Case "15": Me.Shape15.BackColor = &HC0FFFF
   Case "16": Me.Shape16.BackColor = &HC0FFFF
   Case "17": Me.Shape17.BackColor = &HC0FFFF
   Case "18": Me.Shape18.BackColor = &HC0FFFF
   Case "19": Me.Shape19.BackColor = &HC0FFFF
   Case "20": Me.Shape20.BackColor = &HC0FFFF
   Case "21": Me.Shape21.BackColor = &HC0FFFF
   Case "22": Me.Shape22.BackColor = &HC0FFFF
   Case "23": Me.Shape23.BackColor = &HC0FFFF
   Case "24": Me.Shape24.BackColor = &HC0FFFF
   Case "25": Me.Shape25.BackColor = &HC0FFFF
   Case "26": Me.Shape26.BackColor = &HC0FFFF
   Case "27": Me.Shape27.BackColor = &HC0FFFF
   Case "28": Me.Shape28.BackColor = &HC0FFFF
   Case "29": Me.Shape29.BackColor = &HC0FFFF
   Case "30": Me.Shape30.BackColor = &HC0FFFF
   Case "31": Me.Shape31.BackColor = &HC0FFFF
   Case "32": Me.Shape32.BackColor = &HC0FFFF
   Case "33": Me.Shape33.BackColor = &HC0FFFF
   Case "34": Me.Shape34.BackColor = &HC0FFFF
   Case "35": Me.Shape35.BackColor = &HC0FFFF
   Case "36": Me.Shape36.BackColor = &HC0FFFF
  End Select
  
  End If
 
 End If

Me.DtaPeriodos.Refresh
End Sub

Private Sub CmdCrear_Click()
Dim i As Integer, J As Integer, k As Integer
Dim Fecha As Date, Año As Integer, mes As Integer, Dia As Integer
Dim Fecha1 As String
 Me.DtaPeriodos.Refresh
mes = Month(Me.TxtFechaCierre.Value)
Dia = Day(Me.TxtFechaCierre)
Año = Year(Me.TxtFechaCierre)

Me.TxtFechaCierre.Enabled = True

'//////////Borro todos los registros de la tabla////////////
Fecha1 = "  /  /    "
 For i = 1 To 36
   Select Case i
     Case 1: Me.Txt1.Text = Fecha1
     Case 2: Me.Txt2.Text = Fecha1
     Case 3: Me.Txt3.Text = Fecha1
     Case 4: Me.Txt4.Text = Fecha1
     Case 5: Me.Txt5.Text = Fecha1
     Case 6: Me.Txt6.Text = Fecha1
     Case 7: Me.Txt7.Text = Fecha1
     Case 8: Me.Txt8.Text = Fecha1
     Case 9: Me.Txt9.Text = Fecha1
     Case 10: Me.Txt10.Text = Fecha1
     Case 11: Me.Txt11.Text = Fecha1
     Case 12: Me.Txt12.Text = Fecha1
     Case 13: Me.Txt13.Text = Fecha1
     Case 14: Me.Txt14.Text = Fecha1
     Case 15: Me.Txt15.Text = Fecha1
     Case 16: Me.Txt16.Text = Fecha1
     Case 17: Me.Txt17.Text = Fecha1
     Case 18: Me.Txt18.Text = Fecha1
     Case 19: Me.Txt19.Text = Fecha1
     Case 20: Me.Txt20.Text = Fecha1
     Case 21: Me.Txt21.Text = Fecha1
     Case 22: Me.Txt22.Text = Fecha1
     Case 23: Me.Txt23.Text = Fecha1
     Case 24: Me.Txt24.Text = Fecha1
     Case 25: Me.Txt25.Text = Fecha1
     Case 26: Me.Txt26.Text = Fecha1
     Case 27: Me.Txt27.Text = Fecha1
     Case 28: Me.Txt28.Text = Fecha1
     Case 29: Me.Txt29.Text = Fecha1
     Case 30: Me.Txt30.Text = Fecha1
     Case 31: Me.Txt31.Text = Fecha1
     Case 32: Me.Txt32.Text = Fecha1
     Case 33: Me.Txt33.Text = Fecha1
     Case 34: Me.Txt34.Text = Fecha1
     Case 35: Me.Txt35.Text = Fecha1
     Case 36: Me.Txt36.Text = Fecha1
   End Select
 
 
 


 Next

DtaPeriodos.Refresh
If DtaPeriodos.Recordset.EOF Then
 Me.CmdBloquear.Enabled = True
  Me.CmdCerrar.Enabled = True
  Me.CmdDesBloquear.Enabled = True
  Me.CmdAnterior.Enabled = True
  Me.CmdSiguiente.Enabled = True
  For i = 1 To 12
    Fecha = DateSerial(Año, mes + i, 1 - 1)
   Select Case i
     Case 1: Me.Txt1.Text = Fecha
             Me.Lbl1.Caption = 0
     Case 2: Me.Txt2.Text = Fecha
             Me.Lbl2.Caption = 0
     Case 3: Me.Txt3.Text = Fecha
             Me.Lbl3.Caption = 0
     Case 4: Me.Txt4.Text = Fecha
             Me.Lbl4.Caption = 0
     Case 5: Me.Txt5.Text = Fecha
             Me.Lbl5.Caption = 0
     Case 6: Me.Txt6.Text = Fecha
             Me.Lbl6.Caption = 0
     Case 7: Me.Txt7.Text = Fecha
             Me.Lbl7.Caption = 0
     Case 8: Me.Txt8.Text = Fecha
             Me.Lbl8.Caption = 0
     Case 9: Me.Txt9.Text = Fecha
             Me.Lbl9.Caption = 0
     Case 10: Me.Txt10.Text = Fecha
             Me.Lbl10.Caption = 0
     Case 11: Me.Txt11.Text = Fecha
             Me.Lbl11.Caption = 0
     Case 12: Me.Txt12.Text = Fecha
             Me.Lbl12.Caption = 0
   End Select
   Me.DtaPeriodos.Recordset.AddNew
    DtaPeriodos.Recordset("FechaPeriodo") = Fecha
    DtaPeriodos.Recordset("EstadoPeriodo") = "A"
'    DtaPeriodos.Recordset("NPeriodo") = 0
    DtaPeriodos.Recordset("Periodo") = i
    DtaPeriodos.Recordset("NumeroTabla") = 1
  Me.DtaPeriodos.Recordset.Update
  Next
  
  For i = 1 To 12
    Fecha = DateSerial(Año + 1, mes + i, 1 - 1)
   Select Case i
     Case 1: Me.Txt13.Text = Fecha
             Me.Lbl13.Caption = 0
     Case 2: Me.Txt14.Text = Fecha
             Me.Lbl14.Caption = 0
     Case 3: Me.Txt15.Text = Fecha
             Me.Lbl15.Caption = 0
     Case 4: Me.Txt16.Text = Fecha
             Me.Lbl16.Caption = 0
     Case 5: Me.Txt17.Text = Fecha
             Me.Lbl17.Caption = 0
     Case 6: Me.Txt18.Text = Fecha
             Me.Lbl18.Caption = 0
     Case 7: Me.Txt19.Text = Fecha
             Me.Lbl19.Caption = 0
     Case 8: Me.Txt20.Text = Fecha
             Me.Lbl20.Caption = 0
     Case 9: Me.Txt21.Text = Fecha
             Me.Lbl21.Caption = 0
     Case 10: Me.Txt22.Text = Fecha
             Me.Lbl22.Caption = 0
     Case 11: Me.Txt23.Text = Fecha
              Me.Lbl23.Caption = 0
     Case 12: Me.Txt24.Text = Fecha
             Me.Lbl24.Caption = 0
              
   End Select
  Me.DtaPeriodos.Recordset.AddNew
    DtaPeriodos.Recordset("FechaPeriodo") = Fecha
    DtaPeriodos.Recordset("EstadoPeriodo") = "A"
'    DtaPeriodos.Recordset("NPeriodo") = 0
    DtaPeriodos.Recordset("Periodo") = i
    DtaPeriodos.Recordset("NumeroTabla") = 2
  Me.DtaPeriodos.Recordset.Update
  Next
  
  For i = 1 To 12
    Fecha = DateSerial(Año + 2, mes + i, 1 - 1)
   Select Case i
     Case 1: Me.Txt25.Text = Fecha
             Me.Lbl25.Caption = 0
     Case 2: Me.Txt26.Text = Fecha
             Me.Lbl26.Caption = 0
     Case 3: Me.Txt27.Text = Fecha
             Me.Lbl27.Caption = 0
     Case 4: Me.Txt28.Text = Fecha
             Me.Lbl28.Caption = 0
     Case 5: Me.Txt29.Text = Fecha
             Me.Lbl29.Caption = 0
     Case 6: Me.Txt30.Text = Fecha
             Me.Lbl30.Caption = 0
     Case 7: Me.Txt31.Text = Fecha
             Me.Lbl31.Caption = 0
     Case 8: Me.Txt32.Text = Fecha
             Me.Lbl32.Caption = 0
     Case 9: Me.Txt33.Text = Fecha
             Me.Lbl33.Caption = 0
     Case 10: Me.Txt34.Text = Fecha
             Me.Lbl34.Caption = 0
     Case 11: Me.Txt35.Text = Fecha
             Me.Lbl35.Caption = 0
     Case 12: Me.Txt36.Text = Fecha
             Me.Lbl36.Caption = 0
   End Select
   Me.DtaPeriodos.Recordset.AddNew
    DtaPeriodos.Recordset("FechaPeriodo") = Fecha
    DtaPeriodos.Recordset("EstadoPeriodo") = "A"
'    DtaPeriodos.Recordset("NPeriodo") = 0
    DtaPeriodos.Recordset("NumeroTabla") = 3
    DtaPeriodos.Recordset("Periodo") = i
  Me.DtaPeriodos.Recordset.Update
  Next
 
Else
  '///////En caso que exista registro solo agrego un Año///////////
  '//////Muevo la tabla de registro un año///////////
  
  '///////Edito el Periodo Uno y lo saco de la Tabla de 3 años/////////////
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
  DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 0
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop
  '//////////Edito el Periodo 2 y se convierte en el periodo 0/////////
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 2)) ORDER BY Periodos.NPeriodo"
  DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 0
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop
  
  '//////////Edito el Periodo 3 y se convierte en el periodo 0/////////
  '///////Busco los datos del ultimo año/////////////
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
  DtaConsulta.Refresh
  
   Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 0
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop
  
  

  
  
  Me.DtaPeriodos.Recordset.MoveLast
   
  NPeriodoFin = DtaPeriodos.Recordset("NPeriodo")
  NPeriodo = NPeriodoFin - 11
  
  
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos WHERE (((Periodos.NPeriodo) Between " & NPeriodo & " And " & NPeriodoFin & " )) ORDER BY Periodos.NPeriodo"
  Me.DtaConsulta.Refresh
  mes = Month(DtaConsulta.Recordset("FechaPeriodo"))
  Dia = Day(DtaConsulta.Recordset("FechaPeriodo"))
  Año = Year(DtaConsulta.Recordset("FechaPeriodo"))
  
  
  
  For i = 1 To 12
    Fecha = DateSerial(Año + 1, mes + i, 1 - 1)
   Select Case i
     Case 1: Me.Txt25.Text = Fecha
     Case 2: Me.Txt26.Text = Fecha
     Case 3: Me.Txt27.Text = Fecha
     Case 4: Me.Txt28.Text = Fecha
     Case 5: Me.Txt29.Text = Fecha
     Case 6: Me.Txt30.Text = Fecha
     Case 7: Me.Txt31.Text = Fecha
     Case 8: Me.Txt32.Text = Fecha
     Case 9: Me.Txt33.Text = Fecha
     Case 10: Me.Txt34.Text = Fecha
     Case 11: Me.Txt35.Text = Fecha
     Case 12: Me.Txt36.Text = Fecha
   End Select
   Me.DtaPeriodos.Refresh
   Me.DtaPeriodos.Recordset.AddNew
    DtaPeriodos.Recordset("FechaPeriodo") = Fecha
    DtaPeriodos.Recordset("EstadoPeriodo") = "A"
'    DtaPeriodos.Recordset("NPeriodo") = 0
    DtaPeriodos.Recordset("NumeroTabla") = 3
    DtaPeriodos.Recordset("Periodo") = i
  Me.DtaPeriodos.Recordset.Update
  Next

  
  '*******Registro los Datos del Año 2************************
  
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos WHERE (((Periodos.NPeriodo) Between " & NPeriodo & " And " & NPeriodoFin & " )) ORDER BY Periodos.NPeriodo"
  Me.DtaConsulta.Refresh
  
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 2
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop
    
  '******Registro los Datos del Año 1****************
  
  NPeriodo = NPeriodo - 12
  NPeriodoFin = NPeriodo + 11
   
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos WHERE (((Periodos.NPeriodo) Between " & NPeriodo & " And " & NPeriodoFin & " )) ORDER BY Periodos.NPeriodo"
  Me.DtaConsulta.Refresh
  
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 1
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
 Loop
'/////////Actulizo los cambios efectuados en los periodos/////////
  '////LLeno los datos del primer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
If Not DtaConsulta.Recordset.EOF Then
 Me.TxtFechaCierre.Value = DtaConsulta.Recordset("FechaPeriodo")
End If
Do While Not DtaConsulta.Recordset.EOF

Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
 

 End Select


Select Case i
     Case 1: Me.Txt1.Text = Fecha
             Me.Lbl1.Caption = NTransacciones
             Me.Shape1.BackColor = Color
     Case 2: Me.Txt2.Text = Fecha
             Me.Lbl2.Caption = NTransacciones
             Me.Shape2.BackColor = Color
     Case 3: Me.Txt3.Text = Fecha
             Me.Lbl3.Caption = NTransacciones
             Me.Shape3.BackColor = Color
     Case 4: Me.Txt4.Text = Fecha
             Me.Lbl4.Caption = NTransacciones
             Me.Shape4.BackColor = Color
     Case 5: Me.Txt5.Text = Fecha
             Me.Lbl5.Caption = NTransacciones
             Me.Shape5.BackColor = Color
     Case 6: Me.Txt6.Text = Fecha
             Me.Lbl6.Caption = NTransacciones
             Me.Shape6.BackColor = Color
     Case 7: Me.Txt7.Text = Fecha
             Me.Lbl7.Caption = NTransacciones
             Me.Shape7.BackColor = Color
     Case 8: Me.Txt8.Text = Fecha
             Me.Lbl8.Caption = NTransacciones
             Me.Shape8.BackColor = Color
     Case 9: Me.Txt9.Text = Fecha
             Me.Lbl9.Caption = NTransacciones
             Me.Shape9.BackColor = Color
     Case 10: Me.Txt10.Text = Fecha
             Me.Lbl10.Caption = NTransacciones
             Me.Shape10.BackColor = Color
     Case 11: Me.Txt11.Text = Fecha
             Me.Lbl11.Caption = NTransacciones
             Me.Shape11.BackColor = Color
     Case 12: Me.Txt12.Text = Fecha
             Me.Lbl12.Caption = NTransacciones
             Me.Shape12.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del segundo año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 2)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
 End Select
Select Case i
     Case 1: Me.Txt13.Text = Fecha
             Me.Lbl13.Caption = NTransacciones
             Me.Shape13.BackColor = Color
     Case 2: Me.Txt14.Text = Fecha
             Me.Lbl14.Caption = NTransacciones
             Me.Shape14.BackColor = Color
     Case 3: Me.Txt15.Text = Fecha
             Me.Lbl15.Caption = NTransacciones
             Me.Shape15.BackColor = Color
     Case 4: Me.Txt16.Text = Fecha
             Me.Lbl16.Caption = NTransacciones
             Me.Shape16.BackColor = Color
     Case 5: Me.Txt17.Text = Fecha
             Me.Lbl17.Caption = NTransacciones
             Me.Shape17.BackColor = Color
     Case 6: Me.Txt18.Text = Fecha
             Me.Lbl18.Caption = NTransacciones
             Me.Shape18.BackColor = Color
     Case 7: Me.Txt19.Text = Fecha
             Me.Lbl19.Caption = NTransacciones
             Me.Shape19.BackColor = Color
     Case 8: Me.Txt20.Text = Fecha
             Me.Lbl20.Caption = NTransacciones
             Me.Shape20.BackColor = Color
     Case 9: Me.Txt21.Text = Fecha
             Me.Lbl21.Caption = NTransacciones
             Me.Shape21.BackColor = Color
     Case 10: Me.Txt22.Text = Fecha
             Me.Lbl22.Caption = NTransacciones
             Me.Shape22.BackColor = Color
     Case 11: Me.Txt23.Text = Fecha
              Me.Lbl23.Caption = NTransacciones
              Me.Shape23.BackColor = Color
     Case 12: Me.Txt24.Text = Fecha
             Me.Lbl24.Caption = NTransacciones
             Me.Shape24.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del tercer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))

estado = DtaConsulta.Recordset("EstadoPeriodo")
If Not IsNull(DtaConsulta.Recordset("NTransacciones")) Then
    NTransacciones = DtaConsulta.Recordset("NTransacciones")
End If

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
  End Select


Select Case i
     Case 1: Me.Txt25.Text = Fecha
             Me.Lbl25.Caption = NTransacciones
             Me.Shape25.BackColor = Color
     Case 2: Me.Txt26.Text = Fecha
             Me.Lbl26.Caption = NTransacciones
             Me.Shape26.BackColor = Color
     Case 3: Me.Txt27.Text = Fecha
             Me.Lbl27.Caption = NTransacciones
             Me.Shape27.BackColor = Color
     Case 4: Me.Txt28.Text = Fecha
             Me.Lbl28.Caption = NTransacciones
             Me.Shape28.BackColor = Color
     Case 5: Me.Txt29.Text = Fecha
             Me.Lbl29.Caption = NTransacciones
             Me.Shape29.BackColor = Color
     Case 6: Me.Txt30.Text = Fecha
             Me.Lbl30.Caption = NTransacciones
             Me.Shape30.BackColor = Color
     Case 7: Me.Txt31.Text = Fecha
             Me.Lbl31.Caption = NTransacciones
             Me.Shape31.BackColor = Color
     Case 8: Me.Txt32.Text = Fecha
             Me.Lbl32.Caption = NTransacciones
             Me.Shape32.BackColor = Color
     Case 9: Me.Txt33.Text = Fecha
             Me.Lbl33.Caption = NTransacciones
             Me.Shape33.BackColor = Color
     Case 10: Me.Txt34.Text = Fecha
             Me.Lbl34.Caption = NTransacciones
             Me.Shape34.BackColor = Color
     Case 11: Me.Txt35.Text = Fecha
             Me.Lbl35.Caption = NTransacciones
             Me.Shape35.BackColor = Color
     Case 12: Me.Txt36.Text = Fecha
             Me.Lbl36.Caption = NTransacciones
             Me.Shape36.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop


  


End If

End Sub

Private Sub CmdDesBloquear_Click()
Dim NumFecha1 As Long, Fechas As String
 Select Case Opt
   Case "1": Me.Shape1.BackColor = &HC000&
             Me.Txt1.SetFocus
             NumFecha1 = CDate(Me.Txt1.Text)
   Case "2": Me.Shape2.BackColor = &HC000&
             Me.Txt2.SetFocus
             NumFecha1 = CDate(Me.Txt2.Text)
   Case "3": Me.Shape3.BackColor = &HC000&
             Me.Txt3.SetFocus
             NumFecha1 = CDate(Me.Txt3.Text)
    Case "4": Me.Shape4.BackColor = &HC000&
             Me.Txt4.SetFocus
             NumFecha1 = CDate(Me.Txt4.Text)
    Case "5": Me.Shape5.BackColor = &HC000&
             Me.Txt5.SetFocus
             NumFecha1 = CDate(Me.Txt5.Text)
    Case "6": Me.Shape6.BackColor = &HC000&
             Me.Txt6.SetFocus
             NumFecha1 = CDate(Me.Txt6.Text)
   Case "7": Me.Shape7.BackColor = &HC000&
             Me.Txt7.SetFocus
             NumFecha1 = CDate(Me.Txt7.Text)
   Case "8": Me.Shape8.BackColor = &HC000&
             Me.Txt8.SetFocus
             NumFecha1 = CDate(Me.Txt8.Text)
   Case "9": Me.Shape9.BackColor = &HC000&
             Me.Txt9.SetFocus
             NumFecha1 = CDate(Me.Txt9.Text)
  Case "10": Me.Shape10.BackColor = &HC000&
             Me.Txt10.SetFocus
             NumFecha1 = CDate(Me.Txt10.Text)
  Case "11": Me.Shape11.BackColor = &HC000&
             Me.Txt11.SetFocus
             NumFecha1 = CDate(Me.Txt11.Text)
  Case "12": Me.Shape12.BackColor = &HC000&
             Me.Txt12.SetFocus
             NumFecha1 = CDate(Me.Txt12.Text)
  Case "13": Me.Shape13.BackColor = &HC000&
             Me.Txt13.SetFocus
             NumFecha1 = CDate(Me.Txt13.Text)
  Case "14": Me.Shape14.BackColor = &HC000&
             Me.Txt14.SetFocus
             NumFecha1 = CDate(Me.Txt14.Text)
  Case "15": Me.Shape15.BackColor = &HC000&
             Me.Txt15.SetFocus
             NumFecha1 = CDate(Me.Txt15.Text)
  Case "16": Me.Shape16.BackColor = &HC000&
             Me.Txt16.SetFocus
             NumFecha1 = CDate(Me.Txt16.Text)
  Case "17": Me.Shape17.BackColor = &HC000&
             Me.Txt17.SetFocus
             NumFecha1 = CDate(Me.Txt17.Text)
  Case "18": Me.Shape18.BackColor = &HC000&
             Me.Txt18.SetFocus
             NumFecha1 = CDate(Me.Txt18.Text)
  Case "19": Me.Shape19.BackColor = &HC000&
             Me.Txt19.SetFocus
             NumFecha1 = CDate(Me.Txt19.Text)
  Case "20": Me.Shape20.BackColor = &HC000&
             Me.Txt20.SetFocus
             NumFecha1 = CDate(Me.Txt20.Text)
  Case "21": Me.Shape21.BackColor = &HC000&
             Me.Txt21.SetFocus
             NumFecha1 = CDate(Me.Txt21.Text)
  Case "22": Me.Shape22.BackColor = &HC000&
             Me.Txt22.SetFocus
             NumFecha1 = CDate(Me.Txt22.Text)
  Case "23": Me.Shape23.BackColor = &HC000&
             Me.Txt23.SetFocus
             NumFecha1 = CDate(Me.Txt23.Text)
    Case "24": Me.Shape24.BackColor = &HC000&
             Me.Txt24.SetFocus
             NumFecha1 = CDate(Me.Txt24.Text)
   Case "25": Me.Shape25.BackColor = &HC000&
             Me.Txt25.SetFocus
             NumFecha1 = CDate(Me.Txt25.Text)
   Case "26": Me.Shape26.BackColor = &HC000&
             Me.Txt26.SetFocus
             NumFecha1 = CDate(Me.Txt26.Text)
    Case "27": Me.Shape27.BackColor = &HC000&
             Me.Txt27.SetFocus
             NumFecha1 = CDate(Me.Txt27.Text)
    Case "28": Me.Shape28.BackColor = &HC000&
             Me.Txt28.SetFocus
             NumFecha1 = CDate(Me.Txt28.Text)
    Case "29": Me.Shape29.BackColor = &HC000&
             Me.Txt29.SetFocus
             NumFecha1 = CDate(Me.Txt29.Text)
    Case "30": Me.Shape30.BackColor = &HC000&
             Me.Txt30.SetFocus
             NumFecha1 = CDate(Me.Txt30.Text)
    Case "31": Me.Shape31.BackColor = &HC000&
             Me.Txt31.SetFocus
             NumFecha1 = CDate(Me.Txt31.Text)
     Case "32": Me.Shape32.BackColor = &HC000&
             Me.Txt32.SetFocus
             NumFecha1 = CDate(Me.Txt32.Text)
   Case "33": Me.Shape33.BackColor = &HC000&
             Me.Txt33.SetFocus
             NumFecha1 = CDate(Me.Txt33.Text)
             
    Case "34": Me.Shape34.BackColor = &HC000&
             Me.Txt34.SetFocus
             NumFecha1 = CDate(Me.Txt34.Text)
    Case "35": Me.Shape35.BackColor = &HC000&
             Me.Txt35.SetFocus
             NumFecha1 = CDate(Me.Txt35.Text)
     Case "36": Me.Shape36.BackColor = &HC000&
             Me.Txt36.SetFocus
             NumFecha1 = CDate(Me.Txt36.Text)
             
             
 End Select
 
Fechas = CDate(NumFecha1)
Fechas = Format(Fechas, "yyyy/mm/dd")
Me.DtaConsulta.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones From Periodos WHERE     (FechaPeriodo = CONVERT(DATETIME, '" & Fechas & "', 102)) ORDER BY NPeriodo"
 'Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.FechaPeriodo) = " & NumFecha1 & " )) ORDER BY Periodos.NPeriodo"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  If Not DtaConsulta.Recordset("EstadoPeriodo") = "C" Then
   'Me.'DtaConsulta.Recordset.Edit
    DtaConsulta.Recordset("EstadoPeriodo") = "A"
   Me.DtaConsulta.Recordset.Update
  Else
   MsgBox "No se puede Abrir un Periodo Cerrado", vbCritical, "sistema Contable"
   Select Case Opt
   Case "1": Me.Shape1.BackColor = &HFF&
   Case "2": Me.Shape2.BackColor = &HFF&
   Case "3": Me.Shape3.BackColor = &HFF&
   Case "4": Me.Shape4.BackColor = &HFF&
   Case "5": Me.Shape5.BackColor = &HFF&
   Case "6": Me.Shape6.BackColor = &HFF&
   Case "7": Me.Shape7.BackColor = &HFF&
   Case "8": Me.Shape8.BackColor = &HFF&
   Case "9": Me.Shape9.BackColor = &HFF&
   Case "10": Me.Shape10.BackColor = &HFF&
   Case "11": Me.Shape11.BackColor = &HFF&
   Case "12": Me.Shape12.BackColor = &HFF&
   Case "13": Me.Shape13.BackColor = &HFF&
   Case "14": Me.Shape14.BackColor = &HFF&
   Case "15": Me.Shape15.BackColor = &HFF&
   Case "16": Me.Shape16.BackColor = &HFF&
   Case "17": Me.Shape17.BackColor = &HFF&
   Case "18": Me.Shape18.BackColor = &HFF&
   Case "19": Me.Shape19.BackColor = &HFF&
   Case "20": Me.Shape20.BackColor = &HFF&
   Case "21": Me.Shape21.BackColor = &HFF&
   Case "22": Me.Shape22.BackColor = &HFF&
   Case "23": Me.Shape23.BackColor = &HFF&
   Case "24": Me.Shape24.BackColor = &HFF&
   Case "25": Me.Shape25.BackColor = &HFF&
   Case "26": Me.Shape26.BackColor = &HFF&
   Case "27": Me.Shape27.BackColor = &HFF&
   Case "28": Me.Shape28.BackColor = &HFF&
   Case "29": Me.Shape29.BackColor = &HFF&
   Case "30": Me.Shape30.BackColor = &HFF&
   Case "31": Me.Shape31.BackColor = &HFF&
   Case "32": Me.Shape32.BackColor = &HFF&
   Case "33": Me.Shape33.BackColor = &HFF&
   Case "34": Me.Shape34.BackColor = &HFF&
   Case "35": Me.Shape35.BackColor = &HFF&
   Case "36": Me.Shape36.BackColor = &HFF&
  End Select
  
  
  End If
 
 End If
End Sub

Private Sub SmartButton4_Click()

End Sub

Private Sub CmdProcesar_Click()
On Error GoTo TipoErrs
Dim NtransaccionDolar As Integer, NtransaccionCordobas As Integer, NtransaccionEuro As Integer, DescripcionCuenta As String
Dim IndiceDolar As Boolean, IndiceCordobas As Boolean, IndiceEuros As Boolean
Dim TotalDolar As Double, TotalCordobas As Double, TotalEuros As Double, MonedaIndice As String
Dim TotalDebitoCordobas As Double, TotalDebitoDolar As Double
Dim TotalCreditoCordobas As Double, TotalCreditoDolar As Double, Ajuste As String
If Me.TxtContracuenta.Text = "" Then
 MsgBox "Se necesita la Contra Cuenta", vbCritical, "Sistema Contable"
 Exit Sub
End If
Me.TxtContracuenta.Enabled = False
Me.CmdBuscaCuenta.Enabled = False
Me.CmdCancelar.Enabled = False
Me.CmdProcesar.Enabled = False
CreadosIndices = False
'/******Busco la Fecha del primer mes del periodo/////////**************
     NumeroPeriodo = NumeroPeriodo - 11
     Criterio = "NPeriodo=" & NumeroPeriodo & " "
     Me.DtaPeriodos.Recordset.Find (Criterio)
     If Not Me.DtaPeriodos.Recordset.EOF Then
      Fechas1 = "01" & Mid(Me.DtaPeriodos.Recordset("FechaPeriodo"), 3, 10)
      NumFecha1 = CDate(Fechas1)
     
     End If

   CadenaSQL = "SELECT  Transacciones.CodCuentas AS CodCuentas, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Cuentas.TipoCuenta) AS TipoCuenta, " & _
               "MAX(Cuentas.TipoMoneda) AS TipoMoneda, MAX(Transacciones.TCambio) AS TCambio, MAX(Transacciones.Fuente) AS Fuente, " & _
               "MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NPeriodo) AS NPeriodo, MAX(Transacciones.NTransaccion) " & _
               "AS Ntransaccion, SUM(Transacciones.TCambio * Transacciones.Debito - Transacciones.TCambio * Transacciones.Credito) AS Saldo " & _
               "FROM  Cuentas INNER JOIN " & _
               "Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
               "GROUP BY Transacciones.CodCuentas " & _
               "HAVING      (MAX(Cuentas.TipoCuenta) = N'Costos') OR " & _
               "(MAX(Cuentas.TipoCuenta) = N'Gastos') OR " & _
               "(MAX(Cuentas.TipoCuenta) = N'Ingresos - Ventas') " & _
               "ORDER BY Transacciones.CodCuentas, MAX(Cuentas.TipoMoneda)"



'    CadenaSQL = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda, Transacciones.TCambio, Transacciones.Fuente," & vbLf
'    CadenaSQL = CadenaSQL & "Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion," & vbLf
'    CadenaSQL = CadenaSQL & "Transacciones.TCambio * Transacciones.Debito - Transacciones.TCambio * Transacciones.Credito AS Saldo" & vbLf
'    CadenaSQL = CadenaSQL & "FROM         Cuentas INNER JOIN" & vbLf
'    CadenaSQL = CadenaSQL & "Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas" & vbLf
'    CadenaSQL = CadenaSQL & "WHERE     (Cuentas.TipoCuenta = 'Costos') OR" & vbLf
'    CadenaSQL = CadenaSQL & "(Cuentas.TipoCuenta = 'Gastos') OR" & vbLf
'    CadenaSQL = CadenaSQL & "(Cuentas.TipoCuenta = 'Ingresos y Ventas')" & vbLf
'    CadenaSQL = CadenaSQL & "ORDER BY Transacciones.CodCuentas, Cuentas.TipoMoneda"

    Me.DtaSaldos.RecordSource = CadenaSQL
    Me.DtaSaldos.Refresh
    If Not Me.DtaSaldos.Recordset.EOF Then
    Me.DtaSaldos.Recordset.MoveLast
    CantRegistros = Me.DtaSaldos.Recordset.RecordCount
    Me.DtaSaldos.Recordset.MoveFirst
    End If
    Me.BarraCierre.Visible = True
    With Me.BarraCierre
     .Min = 0
     .Max = CantRegistros
     .Value = 0
     i = 1
Total1 = 0
IndiceDolar = False
IndiceCordobas = False
IndiceEuros = False
TotalDolar = 0
TotalCordobas = 0
TotalEuros = 0
 NumeroPeriodo = NumeroPeriodo + 11
    Do While Not Me.DtaSaldos.Recordset.EOF
      .Value = i
      DoEvents
'      MonedaCuenta = Me.DtaSaldos.Recordset("TipoMoneda")
      TipoCuenta = Me.DtaSaldos.Recordset("TipoCuenta")
      
      MonedaCuenta = Me.CmbMoneda.Text
      CodigoCuenta = Me.DtaSaldos.Recordset("CodCuentas")
      DescripcionCuenta = Me.DtaSaldos.Recordset("DescripcionCuentas")
      
         If MonedaCuenta = "Córdobas" Then
            Ajuste = "Dólares"
         ElseIf MonedaCuenta = "Dólares" Then
            Ajuste = "Córdobas"
        
         End If
               
 
      Select Case MonedaCuenta
        Case "Dólares"
              Total1 = 0
                '//////////////////Busco el Saldo que tiene la cuenta////////////////////////////////
                CodigoCuenta = Me.DtaSaldos.Recordset("CodCuentas")
                
  
'                Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.Debito, Transacciones.Credito, Transacciones.Debito*Transacciones.TCambio AS MDebito, Transacciones.TCambio*Transacciones.Credito AS MCredito, Transacciones.TCambio From Transacciones WHERE (((Transacciones.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.FechaTransaccion) Between '" & Format(Fechas1, "yyyymmdd") & "' And '" & Format(Fechas2, "yyyymmdd") & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
                Me.DtaConsulta.RecordSource = "SELECT  Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.Debito, Transacciones.Credito, ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Debito ELSE Transacciones.Debito / Tasas.MontoCordobas END, 4) AS MDebito, ROUND(CASE WHEN IndiceTransaccion.TipoMoneda = 'Dólares' THEN Transacciones.Credito ELSE Transacciones.Credito / Tasas.MontoCordobas END, 4) AS MCredito, Transacciones.TCambio, IndiceTransaccion.TipoMoneda, Tasas.MontoCordobas FROM Transacciones INNER JOIN IndiceTransaccion ON Transacciones.FechaTransaccion = IndiceTransaccion.FechaTransaccion AND Transacciones.NumeroMovimiento = IndiceTransaccion.NumeroMovimiento AND Transacciones.NPeriodo = IndiceTransaccion.Nperiodo INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas " & _
                                              "WHERE   (Transacciones.CodCuentas = '" & CodigoCuenta & "') AND (Transacciones.FechaTransaccion BETWEEN '" & Format(Fechas1, "yyyymmdd") & "' AND '" & Format(Fechas2, "yyyymmdd") & "') AND (IndiceTransaccion.Ajuste <> '" & Ajuste & "') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
                Me.DtaConsulta.Refresh
                Do While Not Me.DtaConsulta.Recordset.EOF
                  'Me.CmbTipoMoneda.Enabled = False
                  If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
                        Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
                    End If
                    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
                      Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
                    End If
                      Total1 = Debito - Credito + Total1
                      Debito = 0
                      Credito = 0
                  Else
                      If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
                         Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
                      End If
                      If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
                           Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
                      End If
                           Total1 = Credito - Debito + Total1
                           Debito = 0
                           Credito = 0
                  End If
                  
                  
                  

                  
                  
                  
  
                  Me.DtaConsulta.Recordset.MoveNext

                Loop
                If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                  TotalDolar = TotalDolar - Total1
                Else
                  TotalDolar = TotalDolar + Total1
                End If
                TotalDolar = Format(TotalDolar, "##,##0.00")
                
     
            If IndiceDolar = False Then
              '///////////////////////////////////////////////////////////////////////////////
              '//////////////Busco el ultimo numero de Tranccion del periodo para agregar una ///////////////////////
              '///////////////////////////////////////////////////////////////////////////////

                Criterio = "NPeriodo=" & NumeroPeriodo & " "
                Me.DtaPeriodos.Recordset.Find (Criterio)
                If Not Me.DtaPeriodos.Recordset.EOF Then
                   NtransaccionDolar = Me.DtaPeriodos.Recordset("NTransacciones") + 1
                End If
                '//////////////////Agrego un Indice para Dolar///////////////////////
                Me.DtaIndices.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion"
                Me.DtaIndices.Refresh
                Me.DtaIndices.Recordset.AddNew
                 Me.DtaIndices.Recordset("FechaTransaccion") = Fechas2
                 Me.DtaIndices.Recordset("NumeroMovimiento") = NtransaccionDolar
                 Me.DtaIndices.Recordset("DescripcionMovimiento") = "Cierre Contable (Dolares)"
                 Me.DtaIndices.Recordset("NPeriodo") = NumeroPeriodo
                 Me.DtaIndices.Recordset("Fuente") = "Cierre"
                 Me.DtaIndices.Recordset("FechaTransaccion") = Fechas2
                 Me.DtaIndices.Recordset("TipoMoneda") = "Dólares"
                Me.DtaIndices.Recordset.Update
                MonedaIndice = "Dólares"
                '/////////////////Agrego el numero de transaccion al Periodo////////////////////////
                'Me.DtaPeriodos.Recordset.Edit
                  Me.DtaPeriodos.Recordset("NTransacciones") = Me.DtaPeriodos.Recordset("NTransacciones") + 1
                Me.DtaPeriodos.Recordset.Update
                IndiceDolar = True
                '////////////////////Agrego una nueva Transaccion/////////////////////////////////
                Me.DtaTransacciones.RecordSource = "SELECT Transacciones.* From Transacciones"
                Me.DtaTransacciones.Refresh
                Me.DtaTransacciones.Recordset.AddNew
                  Me.DtaTransacciones.Recordset("CodCuentas") = CodigoCuenta
                  Me.DtaTransacciones.Recordset("FechaTransaccion") = Fechas2
                  Me.DtaTransacciones.Recordset("NPeriodo") = NumeroPeriodo
                  Me.DtaTransacciones.Recordset("NumeroMovimiento") = NtransaccionDolar
                  Me.DtaTransacciones.Recordset("NombreCuenta") = DescripcionCuenta
                  Me.DtaTransacciones.Recordset("DescripcionMovimiento") = "Cierre Contable(Dolares)"
                  If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   End If
                  
                  Else
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   End If

                  End If
                  Me.DtaTransacciones.Recordset("TCambio") = BuscaTasaCambio(CDate(Fechas2))

                   Me.DtaTransacciones.Recordset("Fuente") = "Cierre"
                  Me.DtaTransacciones.Recordset("FechaTasas") = Fechas2
                Me.DtaTransacciones.Recordset.Update
              Else
                '////////////////////Edito la transaccion/////////////////////////////////
                Me.DtaTransacciones.RecordSource = "SELECT Transacciones.* From Transacciones"
                Me.DtaTransacciones.Refresh
                Me.DtaTransacciones.Recordset.AddNew
                  Me.DtaTransacciones.Recordset("CodCuentas") = CodigoCuenta
                  Me.DtaTransacciones.Recordset("FechaTransaccion") = Fechas2
                  Me.DtaTransacciones.Recordset("NPeriodo") = NumeroPeriodo
                  Me.DtaTransacciones.Recordset("NumeroMovimiento") = NtransaccionDolar
                  Me.DtaTransacciones.Recordset("NombreCuenta") = DescripcionCuenta
                  Me.DtaTransacciones.Recordset("DescripcionMovimiento") = "Cierre Contable (Dolares)"
                  If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   End If
                  
                  Else
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   End If

                  End If
                  Me.DtaTransacciones.Recordset("TCambio") = BuscaTasaCambio(CDate(Fechas2))

                   Me.DtaTransacciones.Recordset("Fuente") = "Cierre"
                  Me.DtaTransacciones.Recordset("FechaTasas") = Fechas2
                Me.DtaTransacciones.Recordset.Update
              
              End If
              
        
              
              

   
          Case "Córdobas"
                Total1 = 0
                '//////////////////Busco el Saldo que tiene la cuenta////////////////////////////////
                CodigoCuenta = Me.DtaSaldos.Recordset("CodCuentas")
                Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.Debito, Transacciones.Credito, Transacciones.Debito*Transacciones.TCambio AS MDebito, Transacciones.TCambio*Transacciones.Credito AS MCredito, Transacciones.TCambio From Transacciones WHERE (((Transacciones.CodCuentas)='" & CodigoCuenta & "') AND ((Transacciones.FechaTransaccion) Between '" & Format(Fechas1, "yyyymmdd") & "' And '" & Format(Fechas2, "yyyymmdd") & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
                Me.DtaConsulta.Refresh
                Do While Not Me.DtaConsulta.Recordset.EOF
                  'Me.CmbTipoMoneda.Enabled = False
                  If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
                        Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
                    End If
                    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
                      Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
                    End If
                      Total1 = Debito - Credito + Total1
                      Debito = 0
                      Credito = 0
                  Else
                      If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
                         Debito = Format(Me.DtaConsulta.Recordset("MDebito"), "##,##0.00")
                      End If
                      If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
                           Credito = Format(Me.DtaConsulta.Recordset("MCredito"), "##,##0.00")
                      End If
                           Total1 = Credito - Debito + Total1
                           Debito = 0
                           Credito = 0
                  End If
                  
                Me.DtaConsulta.Recordset.MoveNext

                Loop
                If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                 TotalCordobas = TotalCordobas - Total1
                Else
                 TotalCordobas = TotalCordobas + Total1
                End If
                TotalCordobas = Format(TotalCordobas, "##,##0.00")
                
     
            If IndiceCordobas = False Then
              '///////////////////////////////////////////////////////////////////////////////
              '//////////////Busco el ultimo numero de Tranccion del periodo para agregar una ///////////////////////
              '///////////////////////////////////////////////////////////////////////////////

                Criterio = "NPeriodo=" & NumeroPeriodo & " "
                Me.DtaPeriodos.Recordset.Find (Criterio)
                If Not Me.DtaPeriodos.Recordset.EOF Then
                   NtransaccionCordobas = Me.DtaPeriodos.Recordset("NTransacciones") + 1
                End If
                '//////////////////Agrego un Indice para Cordobas///////////////////////
                Me.DtaIndices.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion"
                Me.DtaIndices.Refresh
                Me.DtaIndices.Recordset.AddNew
                 Me.DtaIndices.Recordset("FechaTransaccion") = Fechas2
                 Me.DtaIndices.Recordset("NumeroMovimiento") = NtransaccionCordobas
                 Me.DtaIndices.Recordset("DescripcionMovimiento") = "Cierre Contable"
                 Me.DtaIndices.Recordset("NPeriodo") = NumeroPeriodo
                 Me.DtaIndices.Recordset("Fuente") = "Cierre"
                 Me.DtaIndices.Recordset("FechaTransaccion") = Fechas2
                 Me.DtaIndices.Recordset("TipoMoneda") = "Córdobas"
                Me.DtaIndices.Recordset.Update
                MonedaIndice = "Córdobas"
                '/////////////////Agrego el numero de transaccion al Periodo////////////////////////
                'Me.DtaPeriodos.Recordset.Edit
                  Me.DtaPeriodos.Recordset("NTransacciones") = NtransaccionCordobas
                Me.DtaPeriodos.Recordset.Update
                IndiceCordobas = True
                '////////////////////Agrego una nueva Transaccion/////////////////////////////////
                Me.DtaTransacciones.RecordSource = "SELECT Transacciones.* From Transacciones"
                Me.DtaTransacciones.Refresh
                Me.DtaTransacciones.Recordset.AddNew
                  Me.DtaTransacciones.Recordset("CodCuentas") = CodigoCuenta
                  Me.DtaTransacciones.Recordset("FechaTransaccion") = Fechas2
                  Me.DtaTransacciones.Recordset("NPeriodo") = NumeroPeriodo
                  Me.DtaTransacciones.Recordset("NumeroMovimiento") = NtransaccionCordobas
                  Me.DtaTransacciones.Recordset("NombreCuenta") = Me.DtaSaldos.Recordset("DescripcionCuentas")
                  Me.DtaTransacciones.Recordset("DescripcionMovimiento") = "Cierre Contable(Cordobas)"
                  If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   End If
                  
                  Else
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   End If

                  End If
                  
                  Me.DtaTransacciones.Recordset("TCambio") = 1

                   Me.DtaTransacciones.Recordset("Fuente") = "Cierre"
                  Me.DtaTransacciones.Recordset("FechaTasas") = Fechas2
                Me.DtaTransacciones.Recordset.Update
                Total1 = 0
              Else
               If Total1 <> 0 Then
                '////////////////////Edito la transaccion/////////////////////////////////
                Me.DtaTransacciones.RecordSource = "SELECT Transacciones.* From Transacciones"
                Me.DtaTransacciones.Refresh
                Me.DtaTransacciones.Recordset.AddNew
                  Me.DtaTransacciones.Recordset("CodCuentas") = CodigoCuenta
                  Me.DtaTransacciones.Recordset("FechaTransaccion") = Fechas2
                  Me.DtaTransacciones.Recordset("NPeriodo") = NumeroPeriodo
                  Me.DtaTransacciones.Recordset("NumeroMovimiento") = NtransaccionCordobas
                  Me.DtaTransacciones.Recordset("NombreCuenta") = Me.DtaSaldos.Recordset("DescripcionCuentas")
                  Me.DtaTransacciones.Recordset("DescripcionMovimiento") = "Cierre contable (Cordobas)"
                  If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   End If
                  
                  Else
                   If Total1 > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(Total1)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(Total1)
                   End If

                  End If
                  Me.DtaTransacciones.Recordset("TCambio") = 1

                   Me.DtaTransacciones.Recordset("Fuente") = "Cierre"
                  Me.DtaTransacciones.Recordset("FechaTasas") = Fechas2
                Me.DtaTransacciones.Recordset.Update
                
               End If
              End If
          
          
          
   
        End Select
        DoEvents
      i = i + 1
      Me.DtaSaldos.Recordset.MoveNext
    Loop
 
If TotalDolar <> 0 Then
 
    '//////////////////////Despues de agregar las cuentas con sus saldos agrego capital///////////////
    '/////////////////////////////////////////en Dolares////////////////////////////////////////////
                     '////////////////////Edito la transaccion/////////////////////////////////
                Me.DtaTransacciones.RecordSource = "SELECT Transacciones.* From Transacciones"
                Me.DtaTransacciones.Refresh
                Me.DtaTransacciones.Recordset.AddNew
                  Me.DtaTransacciones.Recordset("CodCuentas") = Me.TxtContracuenta.Text
                  Me.DtaTransacciones.Recordset("FechaTransaccion") = Fechas2
                  Me.DtaTransacciones.Recordset("NPeriodo") = NumeroPeriodo
                  Me.DtaTransacciones.Recordset("NumeroMovimiento") = NtransaccionDolar
                  Me.DtaTransacciones.Recordset("NombreCuenta") = DescripcionContracuenta
                  Me.DtaTransacciones.Recordset("DescripcionMovimiento") = "Cierre Contable (Dolares)"
              ' If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                   'If TotalDolar > 0 Then
                    'Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    'Me.DtaTransacciones.Recordset("Debito") = Abs(TotalDolar)
                   'Else
                    'Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    'Me.DtaTransacciones.Recordset("Credito") = Abs(TotalDolar)
                   'End If
                'Else
                   If TotalDolar > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(TotalDolar)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(TotalDolar)
                   End If
                 'end if
                  If MonedaContracuenta = "Dólares" Then
                    Me.DtaTransacciones.Recordset("TCambio") = 1
                  Else
                     Me.DtaTransacciones.Recordset("TCambio") = TasaCambioDolares
                  End If
                   Me.DtaTransacciones.Recordset("Fuente") = "Cierre"
                  Me.DtaTransacciones.Recordset("FechaTasas") = Fechas2
                Me.DtaTransacciones.Recordset.Update
    
End If
If TotalCordobas <> 0 Then
    '//////////////////////Despues de agregar las cuentas con sus saldos agrego capital///////////////
    '/////////////////////////////////////////en Cordobas////////////////////////////////////////////
                     '////////////////////Edito la transaccion/////////////////////////////////
                Me.DtaTransacciones.RecordSource = "SELECT Transacciones.* From Transacciones"
                Me.DtaTransacciones.Refresh
                Me.DtaTransacciones.Recordset.AddNew
                  Me.DtaTransacciones.Recordset("CodCuentas") = Me.TxtContracuenta.Text
                  Me.DtaTransacciones.Recordset("FechaTransaccion") = Fechas2
                  Me.DtaTransacciones.Recordset("NPeriodo") = NumeroPeriodo
                  Me.DtaTransacciones.Recordset("NumeroMovimiento") = NtransaccionCordobas
                  Me.DtaTransacciones.Recordset("NombreCuenta") = DescripcionContracuenta
                  Me.DtaTransacciones.Recordset("DescripcionMovimiento") = "Cierre Contable(Cordobas)"
              ' If TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Then
                   'If TotalCordobas > 0 Then
                    'Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    'Me.DtaTransacciones.Recordset("Debito") = Abs(TotalCordobas)
                   'Else
                    'Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    'Me.DtaTransacciones.Recordset("Credito") = Abs(TotalCordobas)
                   'End If
              'Else
                   If TotalCordobas > 0 Then
                    Me.DtaTransacciones.Recordset("Clave") = "Credito"
                    Me.DtaTransacciones.Recordset("Credito") = Abs(TotalCordobas)
                   Else
                    Me.DtaTransacciones.Recordset("Clave") = "Debito"
                    Me.DtaTransacciones.Recordset("Debito") = Abs(TotalCordobas)
                   End If
              'End If
                  If MonedaContracuenta = "Córdobas" Then
                    Me.DtaTransacciones.Recordset("TCambio") = 1
                  Else
                     Me.DtaTransacciones.Recordset("TCambio") = 1 / TasaCambioDolares
                  End If

                   Me.DtaTransacciones.Recordset("Fuente") = "Cierre"
                  Me.DtaTransacciones.Recordset("FechaTasas") = Fechas2
                Me.DtaTransacciones.Recordset.Update
End If
   Criterio = "NPeriodo=" & NumeroPeriodo & " "
   Me.DtaPeriodos.Recordset.Find (Criterio)
   If Not Me.DtaPeriodos.Recordset.EOF Then
     'Me.DtaPeriodos.Recordset.Edit
       Me.DtaPeriodos.Recordset("EstadoPeriodo") = "C"
     Me.DtaPeriodos.Recordset.Update
   End If
    Select Case Opt
       Case "12": Me.Shape12.BackColor = &HFF&
       Case "24": Me.Shape24.BackColor = &HFF&
       Case "36": Me.Shape36.BackColor = &HFF&
    End Select
  End With

   
  MsgBox "El Periodo se ha Cerrado con Exito", vbInformation, "Sistema Contable"
     Me.CmdAnterior.Visible = True
     Me.CmdBloquear.Visible = True
     Me.CmdCerrar.Visible = True
     Me.CmdCrear.Visible = True
     Me.CmdDesBloquear.Visible = True
     Me.CmdSiguiente.Visible = True
     Me.Label26.Visible = False
     Me.TxtContracuenta.Visible = False
     Me.CmdBuscaCuenta.Visible = False
     Me.SmartButton5.Visible = True
     Me.CmdProcesar.Visible = False
     Me.CmdCancelar.Visible = False
     Me.BarraCierre.Visible = False
Salir = True
Exit Sub
TipoErrs:
      
  If err.Number = 3021 Then
      NumeroPeriodo = NumeroPeriodo + 11
     Criterio = "NPeriodo=" & NumeroPeriodo & " "
     Me.DtaPeriodos.Recordset.Find (Criterio)
     If Not Me.DtaPeriodos.Recordset.EOF Then
       'Me.DtaPeriodos.Recordset.Edit
       Me.DtaPeriodos.Recordset("EstadoPeriodo") = "C"
       Me.DtaPeriodos.Recordset.Update
     End If
     Select Case Opt
       Case "12": Me.Shape12.BackColor = &HFF&
       Case "24": Me.Shape24.BackColor = &HFF&
       Case "36": Me.Shape36.BackColor = &HFF&
     End Select
    
     Me.CmdAnterior.Visible = True
     Me.CmdBloquear.Visible = True
     Me.CmdCerrar.Visible = True
     Me.CmdCrear.Visible = True
     Me.CmdDesBloquear.Visible = True
     Me.CmdSiguiente.Visible = True
     Me.Label26.Visible = False
     Me.TxtContracuenta.Visible = False
     Me.CmdBuscaCuenta.Visible = False
     Me.SmartButton5.Visible = True
     Me.CmdProcesar.Visible = False
     Me.CmdCancelar.Visible = False
     Me.BarraCierre.Visible = False
     Salir = True
     Else
       MsgBox err.Description
     End If
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
'///////////Busco los Datos delUltimo Periodo/////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
If Not DtaConsulta.Recordset.EOF Then
 NPeriodo = Val(DtaConsulta.Recordset("NPeriodo")) + 12
 NPeriodoFin = NPeriodo + 11
End If

 Me.DtaAnSi.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NPeriodo) = " & NPeriodo & " )) ORDER BY Periodos.NPeriodo"
 Me.DtaAnSi.Refresh
If Not DtaAnSi.Recordset.EOF Then

 '//////////Edito el Periodo 1 y se convierte en el periodo 0/////////
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
  DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 0
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop

 '//////////Edito el Periodo 2 y se convierte en el tercer año/////////
  Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 2)) ORDER BY Periodos.NPeriodo"
  DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 1
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop

 '//////Edito los registros del tercer año /////////////////
 '///////se convierte en el segundo año/////////////////////
 
 Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
 DtaConsulta.Refresh
    Do While Not DtaConsulta.Recordset.EOF
        'DtaConsulta.Recordset.Edit
           DtaConsulta.Recordset("NumeroTabla") = 2
        DtaConsulta.Recordset.Update
      DtaConsulta.Recordset.MoveNext
    Loop
 

 
  
 '//////////Edito el Periodo Cero en el 3//////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos WHERE (((Periodos.NPeriodo) Between " & NPeriodo & " And " & NPeriodoFin & " )) ORDER BY Periodos.NPeriodo"
   Me.DtaConsulta.Refresh
  Do While Not DtaConsulta.Recordset.EOF
    'DtaConsulta.Recordset.Edit
      DtaConsulta.Recordset("NumeroTabla") = 3
    DtaConsulta.Recordset.Update
   DtaConsulta.Recordset.MoveNext
  Loop
Else
 '////////////////En caso que este al final de los periodos///////
 MsgBox "Este es el ultimo Registro", vbInformation, "Sistema Contable"
 Exit Sub
End If

'/////////Actulizo los cambios efectuados en los periodos/////////
  '////LLeno los datos del primer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
If Not DtaConsulta.Recordset.EOF Then
 Me.TxtFechaCierre.Value = DtaConsulta.Recordset("FechaPeriodo")
End If
Do While Not DtaConsulta.Recordset.EOF

Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
 

 End Select


Select Case i
     Case 1: Me.Txt1.Text = Fecha
             Me.Lbl1.Caption = NTransacciones
             Me.Shape1.BackColor = Color
     Case 2: Me.Txt2.Text = Fecha
             Me.Lbl2.Caption = NTransacciones
             Me.Shape2.BackColor = Color
     Case 3: Me.Txt3.Text = Fecha
             Me.Lbl3.Caption = NTransacciones
             Me.Shape3.BackColor = Color
     Case 4: Me.Txt4.Text = Fecha
             Me.Lbl4.Caption = NTransacciones
             Me.Shape4.BackColor = Color
     Case 5: Me.Txt5.Text = Fecha
             Me.Lbl5.Caption = NTransacciones
             Me.Shape5.BackColor = Color
     Case 6: Me.Txt6.Text = Fecha
             Me.Lbl6.Caption = NTransacciones
             Me.Shape6.BackColor = Color
     Case 7: Me.Txt7.Text = Fecha
             Me.Lbl7.Caption = NTransacciones
             Me.Shape7.BackColor = Color
     Case 8: Me.Txt8.Text = Fecha
             Me.Lbl8.Caption = NTransacciones
             Me.Shape8.BackColor = Color
     Case 9: Me.Txt9.Text = Fecha
             Me.Lbl9.Caption = NTransacciones
             Me.Shape9.BackColor = Color
     Case 10: Me.Txt10.Text = Fecha
             Me.Lbl10.Caption = NTransacciones
             Me.Shape10.BackColor = Color
     Case 11: Me.Txt11.Text = Fecha
             Me.Lbl11.Caption = NTransacciones
             Me.Shape11.BackColor = Color
     Case 12: Me.Txt12.Text = Fecha
             Me.Lbl12.Caption = NTransacciones
             Me.Shape12.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del segundo año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 2)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
 End Select
Select Case i
     Case 1: Me.Txt13.Text = Fecha
             Me.Lbl13.Caption = NTransacciones
             Me.Shape13.BackColor = Color
     Case 2: Me.Txt14.Text = Fecha
             Me.Lbl14.Caption = NTransacciones
             Me.Shape14.BackColor = Color
     Case 3: Me.Txt15.Text = Fecha
             Me.Lbl15.Caption = NTransacciones
             Me.Shape15.BackColor = Color
     Case 4: Me.Txt16.Text = Fecha
             Me.Lbl16.Caption = NTransacciones
             Me.Shape16.BackColor = Color
     Case 5: Me.Txt17.Text = Fecha
             Me.Lbl17.Caption = NTransacciones
             Me.Shape17.BackColor = Color
     Case 6: Me.Txt18.Text = Fecha
             Me.Lbl18.Caption = NTransacciones
             Me.Shape18.BackColor = Color
     Case 7: Me.Txt19.Text = Fecha
             Me.Lbl19.Caption = NTransacciones
             Me.Shape19.BackColor = Color
     Case 8: Me.Txt20.Text = Fecha
             Me.Lbl20.Caption = NTransacciones
             Me.Shape20.BackColor = Color
     Case 9: Me.Txt21.Text = Fecha
             Me.Lbl21.Caption = NTransacciones
             Me.Shape21.BackColor = Color
     Case 10: Me.Txt22.Text = Fecha
             Me.Lbl22.Caption = NTransacciones
             Me.Shape22.BackColor = Color
     Case 11: Me.Txt23.Text = Fecha
              Me.Lbl23.Caption = NTransacciones
              Me.Shape23.BackColor = Color
     Case 12: Me.Txt24.Text = Fecha
             Me.Lbl24.Caption = NTransacciones
             Me.Shape24.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del tercer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))

estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
  End Select


Select Case i
     Case 1: Me.Txt25.Text = Fecha
             Me.Lbl25.Caption = NTransacciones
             Me.Shape25.BackColor = Color
     Case 2: Me.Txt26.Text = Fecha
             Me.Lbl26.Caption = NTransacciones
             Me.Shape26.BackColor = Color
     Case 3: Me.Txt27.Text = Fecha
             Me.Lbl27.Caption = NTransacciones
             Me.Shape27.BackColor = Color
     Case 4: Me.Txt28.Text = Fecha
             Me.Lbl28.Caption = NTransacciones
             Me.Shape28.BackColor = Color
     Case 5: Me.Txt29.Text = Fecha
             Me.Lbl29.Caption = NTransacciones
             Me.Shape29.BackColor = Color
     Case 6: Me.Txt30.Text = Fecha
             Me.Lbl30.Caption = NTransacciones
             Me.Shape30.BackColor = Color
     Case 7: Me.Txt31.Text = Fecha
             Me.Lbl31.Caption = NTransacciones
             Me.Shape31.BackColor = Color
     Case 8: Me.Txt32.Text = Fecha
             Me.Lbl32.Caption = NTransacciones
             Me.Shape32.BackColor = Color
     Case 9: Me.Txt33.Text = Fecha
             Me.Lbl33.Caption = NTransacciones
             Me.Shape33.BackColor = Color
     Case 10: Me.Txt34.Text = Fecha
             Me.Lbl34.Caption = NTransacciones
             Me.Shape34.BackColor = Color
     Case 11: Me.Txt35.Text = Fecha
             Me.Lbl35.Caption = NTransacciones
             Me.Shape35.BackColor = Color
     Case 12: Me.Txt36.Text = Fecha
             Me.Lbl36.Caption = NTransacciones
             Me.Shape36.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop


Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErrs
Me.Txt1.Height = 255
If Not CodigoUsuario = 0 Then
 Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Periodos'))"
 Me.DtaNacceso.Refresh
 If Me.DtaNacceso.Recordset.EOF Then
   Me.CmdCrear.Enabled = False
   Me.CmdCerrar.Enabled = False
   Me.CmdBloquear.Enabled = False
   Me.CmdDesBloquear.Enabled = False
 End If
End If
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub Form_Load()
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

Salir = True
With Me.DtaTasas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With
With Me.DtaIndices
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With
With Me.DtaTransacciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaSaldos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaPeriodos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Periodos"
   .Refresh
End With

With Me.DtaAnSi
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


Me.DtaPeriodos.Refresh
If DtaPeriodos.Recordset.EOF Then
  Me.TxtFechaCierre.Enabled = True
  Me.CmdBloquear.Enabled = False
  Me.CmdCerrar.Enabled = False
  Me.CmdDesBloquear.Enabled = False
  Me.CmdAnterior.Enabled = False
  Me.CmdSiguiente.Enabled = False
Else
  Me.CmdBloquear.Enabled = True
  Me.CmdCerrar.Enabled = True
  Me.CmdDesBloquear.Enabled = True
  Me.CmdAnterior.Enabled = True
  Me.CmdSiguiente.Enabled = True
  
  Me.TxtFechaCierre.Enabled = False
End If

'////LLeno los datos del primer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 1)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
If Not DtaConsulta.Recordset.EOF Then
 Me.TxtFechaCierre.Value = DtaConsulta.Recordset("FechaPeriodo")
End If
Do While Not DtaConsulta.Recordset.EOF

Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
 

 End Select


Select Case i
     Case 1: Me.Txt1.Text = Fecha
             Me.Lbl1.Caption = NTransacciones
             Me.Shape1.BackColor = Color
     Case 2: Me.Txt2.Text = Fecha
             Me.Lbl2.Caption = NTransacciones
             Me.Shape2.BackColor = Color
     Case 3: Me.Txt3.Text = Fecha
             Me.Lbl3.Caption = NTransacciones
             Me.Shape3.BackColor = Color
     Case 4: Me.Txt4.Text = Fecha
             Me.Lbl4.Caption = NTransacciones
             Me.Shape4.BackColor = Color
     Case 5: Me.Txt5.Text = Fecha
             Me.Lbl5.Caption = NTransacciones
             Me.Shape5.BackColor = Color
     Case 6: Me.Txt6.Text = Fecha
             Me.Lbl6.Caption = NTransacciones
             Me.Shape6.BackColor = Color
     Case 7: Me.Txt7.Text = Fecha
             Me.Lbl7.Caption = NTransacciones
             Me.Shape7.BackColor = Color
     Case 8: Me.Txt8.Text = Fecha
             Me.Lbl8.Caption = NTransacciones
             Me.Shape8.BackColor = Color
     Case 9: Me.Txt9.Text = Fecha
             Me.Lbl9.Caption = NTransacciones
             Me.Shape9.BackColor = Color
     Case 10: Me.Txt10.Text = Fecha
             Me.Lbl10.Caption = NTransacciones
             Me.Shape10.BackColor = Color
     Case 11: Me.Txt11.Text = Fecha
             Me.Lbl11.Caption = NTransacciones
             Me.Shape11.BackColor = Color
     Case 12: Me.Txt12.Text = Fecha
             Me.Lbl12.Caption = NTransacciones
             Me.Shape12.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del segundo año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 2)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))
estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
 End Select
Select Case i
     Case 1: Me.Txt13.Text = Fecha
             Me.Lbl13.Caption = NTransacciones
             Me.Shape13.BackColor = Color
     Case 2: Me.Txt14.Text = Fecha
             Me.Lbl14.Caption = NTransacciones
             Me.Shape14.BackColor = Color
     Case 3: Me.Txt15.Text = Fecha
             Me.Lbl15.Caption = NTransacciones
             Me.Shape15.BackColor = Color
     Case 4: Me.Txt16.Text = Fecha
             Me.Lbl16.Caption = NTransacciones
             Me.Shape16.BackColor = Color
     Case 5: Me.Txt17.Text = Fecha
             Me.Lbl17.Caption = NTransacciones
             Me.Shape17.BackColor = Color
     Case 6: Me.Txt18.Text = Fecha
             Me.Lbl18.Caption = NTransacciones
             Me.Shape18.BackColor = Color
     Case 7: Me.Txt19.Text = Fecha
             Me.Lbl19.Caption = NTransacciones
             Me.Shape19.BackColor = Color
     Case 8: Me.Txt20.Text = Fecha
             Me.Lbl20.Caption = NTransacciones
             Me.Shape20.BackColor = Color
     Case 9: Me.Txt21.Text = Fecha
             Me.Lbl21.Caption = NTransacciones
             Me.Shape21.BackColor = Color
     Case 10: Me.Txt22.Text = Fecha
             Me.Lbl22.Caption = NTransacciones
             Me.Shape22.BackColor = Color
     Case 11: Me.Txt23.Text = Fecha
              Me.Lbl23.Caption = NTransacciones
              Me.Shape23.BackColor = Color
     Case 12: Me.Txt24.Text = Fecha
             Me.Lbl24.Caption = NTransacciones
             Me.Shape24.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop

'////LLeno los datos del tercer año//////////////////
Me.DtaConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones From Periodos Where (((Periodos.NumeroTabla) = 3)) ORDER BY Periodos.NPeriodo"
DtaConsulta.Refresh
i = 1
Do While Not DtaConsulta.Recordset.EOF
Fecha = Str(DtaConsulta.Recordset("FechaPeriodo"))

estado = DtaConsulta.Recordset("EstadoPeriodo")
NTransacciones = DtaConsulta.Recordset("NTransacciones")

 Select Case estado
 '//Color Verde////
   Case "A": Color = &HC000&
  '///color Amarillo///
   Case "B": Color = &HC0FFFF
   '////Color Rojo/////
   Case "C": Color = &HFF&
        
  End Select


Select Case i
     Case 1: Me.Txt25.Text = Fecha
             Me.Lbl25.Caption = NTransacciones
             Me.Shape25.BackColor = Color
     Case 2: Me.Txt26.Text = Fecha
             Me.Lbl26.Caption = NTransacciones
             Me.Shape26.BackColor = Color
     Case 3: Me.Txt27.Text = Fecha
             Me.Lbl27.Caption = NTransacciones
             Me.Shape27.BackColor = Color
     Case 4: Me.Txt28.Text = Fecha
             Me.Lbl28.Caption = NTransacciones
             Me.Shape28.BackColor = Color
     Case 5: Me.Txt29.Text = Fecha
             Me.Lbl29.Caption = NTransacciones
             Me.Shape29.BackColor = Color
     Case 6: Me.Txt30.Text = Fecha
             Me.Lbl30.Caption = NTransacciones
             Me.Shape30.BackColor = Color
     Case 7: Me.Txt31.Text = Fecha
             Me.Lbl31.Caption = NTransacciones
             Me.Shape31.BackColor = Color
     Case 8: Me.Txt32.Text = Fecha
             Me.Lbl32.Caption = NTransacciones
             Me.Shape32.BackColor = Color
     Case 9: Me.Txt33.Text = Fecha
             Me.Lbl33.Caption = NTransacciones
             Me.Shape33.BackColor = Color
     Case 10: Me.Txt34.Text = Fecha
             Me.Lbl34.Caption = NTransacciones
             Me.Shape34.BackColor = Color
     Case 11: Me.Txt35.Text = Fecha
             Me.Lbl35.Caption = NTransacciones
             Me.Shape35.BackColor = Color
     Case 12: Me.Txt36.Text = Fecha
             Me.Lbl36.Caption = NTransacciones
             Me.Shape36.BackColor = Color
   End Select
   i = i + 1
 DtaConsulta.Recordset.MoveNext
Loop


End Sub

Private Sub SmartButton1_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Salir = False Then
  Cancel = 1
Else
  Cancel = 0
End If
End Sub

Private Sub SmartButton5_Click()
Unload Me
End Sub
Private Sub Txt1_GotFocus()
 Opt = "1"
  Me.Label109.BorderStyle = 0
End Sub

Private Sub Txt1_LostFocus()
 Me.Label109.BorderStyle = 1
End Sub
Private Sub Txt2_LostFocus()
 Me.Label108.BorderStyle = 1
End Sub
Private Sub Txt3_LostFocus()
 Me.Label107.BorderStyle = 1
End Sub
Private Sub Txt4_LostFocus()
 Me.Label106.BorderStyle = 1
End Sub
Private Sub Txt5_LostFocus()
 Me.Label105.BorderStyle = 1
End Sub
Private Sub Txt6_LostFocus()
 Me.Label104.BorderStyle = 1
End Sub
Private Sub Txt7_LostFocus()
 Me.Label103.BorderStyle = 1
End Sub
Private Sub Txt8_LostFocus()
 Me.Label102.BorderStyle = 1
End Sub
Private Sub Txt9_LostFocus()
 Me.Label101.BorderStyle = 1
End Sub
Private Sub Txt10_LostFocus()
 Me.Label100.BorderStyle = 1
End Sub
Private Sub Txt11_LostFocus()
 Me.Label99.BorderStyle = 1
End Sub
Private Sub Txt12_LostFocus()
 Me.Label98.BorderStyle = 1
End Sub
Private Sub Txt13_LostFocus()
 Me.Label73.BorderStyle = 1
End Sub
Private Sub Txt14_LostFocus()
 Me.Label72.BorderStyle = 1
End Sub
Private Sub Txt15_LostFocus()
 Me.Label71.BorderStyle = 1
End Sub
Private Sub Txt16_LostFocus()
 Me.Label70.BorderStyle = 1
End Sub
Private Sub Txt17_LostFocus()
 Me.Label69.BorderStyle = 1
End Sub
Private Sub Txt18_LostFocus()
 Me.Label68.BorderStyle = 1
End Sub
Private Sub Txt19_LostFocus()
 Me.Label67.BorderStyle = 1
End Sub
Private Sub Txt20_LostFocus()
 Me.Label66.BorderStyle = 1
End Sub
Private Sub Txt21_LostFocus()
 Me.Label65.BorderStyle = 1
End Sub
Private Sub Txt22_LostFocus()
 Me.Label64.BorderStyle = 1
End Sub
Private Sub Txt23_LostFocus()
 Me.Label63.BorderStyle = 1
End Sub
Private Sub Txt24_LostFocus()
 Me.Label62.BorderStyle = 1
End Sub
Private Sub Txt25_LostFocus()
 Me.Label2.BorderStyle = 1
End Sub
Private Sub Txt26_LostFocus()
 Me.Label3.BorderStyle = 1
End Sub
Private Sub Txt27_LostFocus()
 Me.Label4.BorderStyle = 1
End Sub
Private Sub Txt28_LostFocus()
 Me.Label5.BorderStyle = 1
End Sub
Private Sub Txt29_LostFocus()
 Me.Label6.BorderStyle = 1
End Sub
Private Sub Txt30_LostFocus()
 Me.Label7.BorderStyle = 1
End Sub
Private Sub Txt31_LostFocus()
 Me.Label8.BorderStyle = 1
End Sub
Private Sub Txt32_LostFocus()
 Me.Label9.BorderStyle = 1
End Sub
Private Sub Txt33_LostFocus()
 Me.Label10.BorderStyle = 1
End Sub
Private Sub Txt34_LostFocus()
 Me.Label11.BorderStyle = 1
End Sub
Private Sub Txt35_LostFocus()
 Me.Label12.BorderStyle = 1
End Sub
Private Sub Txt36_LostFocus()
 Me.Label13.BorderStyle = 1
End Sub

Private Sub Txt2_GotFocus()
 Opt = "2"
   Me.Label108.BorderStyle = 0
End Sub
Private Sub Txt3_GotFocus()
 Opt = "3"
   Me.Label107.BorderStyle = 0
End Sub
Private Sub Txt4_GotFocus()
 Opt = "4"
   Me.Label106.BorderStyle = 0
End Sub
Private Sub Txt5_GotFocus()
 Opt = "5"
   Me.Label105.BorderStyle = 0
End Sub
Private Sub Txt6_GotFocus()
 Opt = "6"
   Me.Label104.BorderStyle = 0
End Sub
Private Sub Txt7_GotFocus()
 Opt = "7"
   Me.Label103.BorderStyle = 0
End Sub
Private Sub Txt8_GotFocus()
 Opt = "8"
   Me.Label102.BorderStyle = 0
End Sub
Private Sub Txt9_GotFocus()
 Opt = "9"
   Me.Label101.BorderStyle = 0
End Sub
Private Sub Txt10_GotFocus()
 Opt = "10"
   Me.Label100.BorderStyle = 0
End Sub
Private Sub Txt11_GotFocus()
 Opt = "11"
   Me.Label99.BorderStyle = 0
End Sub
Private Sub Txt12_GotFocus()
 Opt = "12"
   Me.Label98.BorderStyle = 0
End Sub
Private Sub Txt13_GotFocus()
 Opt = "13"
    Me.Label73.BorderStyle = 0
End Sub
Private Sub Txt14_GotFocus()
 Opt = "14"
     Me.Label72.BorderStyle = 0
End Sub
Private Sub Txt15_GotFocus()
 Opt = "15"
     Me.Label71.BorderStyle = 0
End Sub
Private Sub Txt16_GotFocus()
 Opt = "16"
    Me.Label70.BorderStyle = 0
End Sub
Private Sub Txt17_GotFocus()
 Opt = "17"
     Me.Label69.BorderStyle = 0
End Sub
Private Sub Txt18_GotFocus()
 Opt = "18"
     Me.Label68.BorderStyle = 0
End Sub
Private Sub Txt19_GotFocus()
 Opt = "19"
     Me.Label67.BorderStyle = 0
End Sub
Private Sub Txt20_GotFocus()
 Opt = "20"
     Me.Label66.BorderStyle = 0
End Sub
Private Sub Txt21_GotFocus()
 Opt = "21"
     Me.Label65.BorderStyle = 0
End Sub
Private Sub Txt22_GotFocus()
 Opt = "22"
     Me.Label64.BorderStyle = 0
End Sub
Private Sub Txt23_GotFocus()
 Opt = "23"
     Me.Label63.BorderStyle = 0
End Sub
Private Sub Txt24_GotFocus()
 Opt = "24"
     Me.Label62.BorderStyle = 0
End Sub
Private Sub Txt25_GotFocus()
 Opt = "25"
      Me.Label2.BorderStyle = 0
End Sub
Private Sub Txt26_GotFocus()
 Opt = "26"
      Me.Label3.BorderStyle = 0
End Sub
Private Sub Txt27_GotFocus()
 Opt = "27"
      Me.Label4.BorderStyle = 0
End Sub
Private Sub Txt28_GotFocus()
 Opt = "28"
      Me.Label5.BorderStyle = 0
End Sub
Private Sub Txt29_GotFocus()
 Opt = "29"
      Me.Label6.BorderStyle = 0
End Sub
Private Sub Txt30_GotFocus()
 Opt = "30"
      Me.Label7.BorderStyle = 0
End Sub

Private Sub Txt31_GotFocus()
 Opt = "31"
      Me.Label8.BorderStyle = 0
End Sub
Private Sub Txt32_GotFocus()
 Opt = "32"
      Me.Label9.BorderStyle = 0
End Sub
Private Sub Txt33_GotFocus()
 Opt = "33"
      Me.Label10.BorderStyle = 0
End Sub
Private Sub Txt34_GotFocus()
 Opt = "34"
      Me.Label11.BorderStyle = 0
End Sub
Private Sub Txt35_GotFocus()
 Opt = "35"
      Me.Label12.BorderStyle = 0
End Sub
Private Sub Txt36_GotFocus()
 Opt = "36"
      Me.Label13.BorderStyle = 0
End Sub
