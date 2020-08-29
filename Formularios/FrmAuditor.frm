VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmAuditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auditoria de Movimientos"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoIndice 
      Height          =   495
      Left            =   960
      Top             =   3600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
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
      Caption         =   "AdoIndice"
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
   Begin MSAdodcLib.Adodc AdoTasas 
      Height          =   375
      Left            =   5040
      Top             =   6960
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Caption         =   "AdoTasas"
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
   Begin XtremeSuiteControls.ProgressBar Barra2 
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   720
      Width           =   4335
      _Version        =   786432
      _ExtentX        =   7646
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   495
      Left            =   240
      TabIndex        =   28
      Top             =   120
      Width           =   9375
      _Version        =   786432
      _ExtentX        =   16536
      _ExtentY        =   873
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin VB.CommandButton CmdReparar 
      Caption         =   "Reparar"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   7920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoPeriodos 
      Height          =   375
      Left            =   4920
      Top             =   5640
      Visible         =   0   'False
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
      Caption         =   "AdoPeriodos"
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   4920
      Top             =   6360
      Visible         =   0   'False
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
      Caption         =   "AdoConsuta"
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
   Begin MSAdodcLib.Adodc AdoGrupos 
      Height          =   375
      Left            =   720
      Top             =   6840
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      Caption         =   "AdoGrupos"
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   495
      Left            =   720
      Top             =   6120
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   873
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
      Caption         =   "AdoCuentas"
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
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   7920
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3413
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "Auditoria General"
      TabPicture(0)   =   "FrmAuditor.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Convertir Transacciones"
      TabPicture(1)   =   "FrmAuditor.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Auditoria Cuentas"
      TabPicture(2)   =   "FrmAuditor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdTasas"
      Tab(2).Control(1)=   "CmdAuditarCuentas"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton CmdTasas 
         Caption         =   "Tasa Cambio"
         Height          =   375
         Left            =   -72960
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CmdAuditarCuentas 
         Caption         =   "Auditar "
         Height          =   375
         Left            =   -74520
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   9255
         Begin VB.CommandButton CmdCorregir 
            Caption         =   "Corregir"
            Height          =   375
            Left            =   1440
            TabIndex        =   32
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton CmdConvertirMonto 
            Caption         =   "Convertir"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox CmbTipoMoneda 
            Height          =   315
            ItemData        =   "FrmAuditor.frx":0054
            Left            =   840
            List            =   "FrmAuditor.frx":005E
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton CmdConvertir 
            Caption         =   "Cambiar Tipo"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   160
            Width           =   1215
         End
         Begin VB.Frame Frame4 
            Caption         =   "Rango para Convertir"
            Height          =   1215
            Left            =   2640
            TabIndex        =   13
            Top             =   120
            Width           =   6495
            Begin VB.TextBox TxtTransaccion 
               Height          =   285
               Left            =   4920
               TabIndex        =   24
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.OptionButton OptConvertirTransaccion 
               Caption         =   "Transaccion"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   840
               Width           =   1575
            End
            Begin VB.OptionButton OptConvertirTodo 
               Caption         =   "Todo"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptConvertirRango 
               Caption         =   "Rango de Fechas"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   540
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel LblHasta 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "FrmAuditor.frx":0075
               TabIndex        =   16
               Top             =   600
               Visible         =   0   'False
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel LblDesde 
               Height          =   255
               Left            =   2040
               OleObjectBlob   =   "FrmAuditor.frx":00DF
               TabIndex        =   17
               Top             =   600
               Visible         =   0   'False
               Width           =   615
            End
            Begin MSComCtl2.DTPicker DTFechaFin 
               Height          =   285
               Left            =   4920
               TabIndex        =   18
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Format          =   79560705
               CurrentDate     =   39117
            End
            Begin MSComCtl2.DTPicker DTFechaIni 
               Height          =   285
               Left            =   2640
               TabIndex        =   19
               Top             =   600
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Format          =   79560705
               CurrentDate     =   39117
            End
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   600
            OleObjectBlob   =   "FrmAuditor.frx":0147
            TabIndex        =   22
            Top             =   1080
            Width           =   135
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   9255
         Begin VB.Frame Frame1 
            Caption         =   "Rango de Auditoria"
            Height          =   975
            Left            =   1920
            TabIndex        =   5
            Top             =   360
            Width           =   6735
            Begin VB.OptionButton Option2 
               Caption         =   "Rango de Fechas"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   600
               Width           =   1575
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Todo"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Value           =   -1  'True
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Left            =   4320
               OleObjectBlob   =   "FrmAuditor.frx":01A9
               TabIndex        =   8
               Top             =   600
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Left            =   2040
               OleObjectBlob   =   "FrmAuditor.frx":0213
               TabIndex        =   9
               Top             =   600
               Width           =   615
            End
            Begin MSComCtl2.DTPicker DTPFechaFin 
               Height          =   285
               Left            =   4920
               TabIndex        =   10
               Top             =   600
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   79560705
               CurrentDate     =   39117
            End
            Begin MSComCtl2.DTPicker DTPFechaIni 
               Height          =   285
               Left            =   2640
               TabIndex        =   11
               Top             =   600
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   79560705
               CurrentDate     =   39117
            End
         End
         Begin VB.CommandButton CmdAuditoria 
            Caption         =   "Auditar "
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin MSAdodcLib.Adodc DtaTransacciones 
      Height          =   375
      Left            =   840
      Top             =   4320
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
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
   Begin MSAdodcLib.Adodc DtaTasas 
      Height          =   375
      Left            =   4800
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
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
   Begin MSAdodcLib.Adodc DtaIndiceTransaccion 
      Height          =   495
      Left            =   720
      Top             =   4920
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
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
      Caption         =   "DtaIndiceTransaccion"
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
   Begin MSAdodcLib.Adodc DtaMovimientos 
      Height          =   495
      Left            =   720
      Top             =   5520
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
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
      Caption         =   "DtaMovimientos"
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
   Begin VB.ListBox List1 
      Height          =   4545
      ItemData        =   "FrmAuditor.frx":027B
      Left            =   120
      List            =   "FrmAuditor.frx":027D
      TabIndex        =   0
      Top             =   3240
      Width           =   9495
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblProceso 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "FrmAuditor.frx":027F
      TabIndex        =   25
      Top             =   750
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "FrmAuditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAuditarCuentas_Click()
On Error GoTo TipoErrs
 Dim TipoCuenta As String, UbicacionReporte As String, CodigoCuenta As String
 Dim KeyGrupo As String, DescripcionGrupo As String, Registros As Double
 Dim i As Double, KeyGrupoF As String, DescripcionGrupoF As String
 
     Me.SSTab1.Enabled = False
     Me.List1.Clear
    
     
     Me.AdoCuentas.RecordSource = "SELECT * From Cuentas"
     Me.AdoCuentas.Refresh
     If Not Me.AdoCuentas.Recordset.EOF Then
       Me.AdoCuentas.Recordset.MoveLast
       Registros = Me.AdoCuentas.Recordset.RecordCount
       Me.AdoCuentas.Recordset.MoveFirst
     End If
     With Barra
       .Min = 0
       .Max = Registros
       .Value = 0
       i = 1
       Me.LblProceso.Visible = True
   
        Do While Not Me.AdoCuentas.Recordset.EOF
         .Value = i
         DoEvents
         
         Me.LblProceso.Caption = "Procesando " & i & " de " & Registros
         TipoCuenta = Me.AdoCuentas.Recordset("TipoCuenta")
         KeyGrupoF = Me.AdoCuentas.Recordset("KeyGrupo")
         DescripcionGrupoF = Me.AdoCuentas.Recordset("DescripcionGrupo")
         CodigoCuenta = Me.AdoCuentas.Recordset("CodCuentas")
         
         If CodigoCuenta = "6000-01-04-" Then
          Cod = 1
         End If
         
         '//////////////////////////BUSCO SI LA CUENTA TIENE BIEN EL GRUPO GRABADO/////////////////////
         Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (DescripcionGrupo = '" & DescripcionGrupoF & "') ORDER BY KeyGrupo"
         Me.AdoGrupos.Refresh
         If Me.AdoGrupos.Recordset.EOF Then
                Me.List1.AddItem ("---------O-----------")
                Me.List1.AddItem ("La Cuenta: " & CodigoCuenta & "No tiene Asignado Grupo")
                
                '////////////BUSCO EL GRUPO PARA ASIGNARLO A LA CUENTA////////////////////////////////
                Select Case TipoCuenta
                    Case "Caja"
                          KeyGrupo = "A"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupoF = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Bancos"
                          KeyGrupo = "A"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Cuentas x Cobrar"
                          KeyGrupo = "A"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Inventario"
                          KeyGrupo = "A"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Activo Fijo"
                          KeyGrupo = "A"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Papeleria - Utiles"
                          KeyGrupo = "A"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Otros Activos"
                          KeyGrupo = "A"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Cuentas x Pagar"
                          KeyGrupo = "B"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Pasivo"
                          KeyGrupo = "B"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Otros Pasivos"
                          KeyGrupo = "B"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Capital"
                          KeyGrupo = "C"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Ingresos - Ventas"
                          KeyGrupo = "D"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Costos"
                          KeyGrupo = "G"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                    Case "Gastos"
                          KeyGrupo = "O"
                          Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                          Me.AdoGrupos.Refresh
                          If Not Me.AdoGrupos.Recordset.EOF Then
                          DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                          End If
                End Select
                
                '//////////////////////BUSCO LA CUENTA /////////////////////////////////////////////////
                Me.AdoConsulta.RecordSource = "SELECT  * From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
                Me.AdoConsulta.Refresh
                If Not Me.AdoConsulta.Recordset.EOF Then
                 Me.AdoConsulta.Recordset("KeyGrupo") = KeyGrupo
                 Me.AdoConsulta.Recordset("DescripcionGrupo") = DescripcionGrupo
                 Me.AdoConsulta.Recordset.Update
                End If
                
                Me.List1.AddItem ("Se Ubico en el grupo: " & DescripcionGrupo)
         Else
                 '//////////////////////////BUSCO EL GRUPO DE LA CUENTA EN LA TABLA GRUPOS/////////////////////
                Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupoF & "') ORDER BY KeyGrupo"
                Me.AdoGrupos.Refresh
                If Not Me.AdoGrupos.Recordset.EOF Then
                   DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                   If DescripcionGrupoF <> DescripcionGrupo Then
                      Me.List1.AddItem ("---------O-----------")
                      Me.List1.AddItem ("La Cuenta: " & CodigoCuenta & "Tiene mal Asignado el Grupo")
                   
                        Select Case TipoCuenta
                            Case "Caja"
                                  KeyGrupo = "A"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupoF = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Bancos"
                                  KeyGrupo = "A"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Cuentas x Cobrar"
                                  KeyGrupo = "A"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Inventario"
                                  KeyGrupo = "A"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Activo Fijo"
                                  KeyGrupo = "A"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Papeleria - Utiles"
                                  KeyGrupo = "A"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Otros Activos"
                                  KeyGrupo = "A"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Cuentas x Pagar"
                                  KeyGrupo = "B"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Pasivo"
                                  KeyGrupo = "B"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Otros Pasivos"
                                  KeyGrupo = "B"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Capital"
                                  KeyGrupo = "C"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Ingresos - Ventas"
                                  KeyGrupo = "D"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Costos"
                                  KeyGrupo = "G"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                            Case "Gastos"
                                  KeyGrupo = "O"
                                  Me.AdoGrupos.RecordSource = "SELECT KeyGrupo, KeyGrupoSuperior, Child, DescripcionGrupo From Grupos WHERE  (KeyGrupo = '" & KeyGrupo & "') ORDER BY KeyGrupo"
                                  Me.AdoGrupos.Refresh
                                  If Not Me.AdoGrupos.Recordset.EOF Then
                                  DescripcionGrupo = Me.AdoGrupos.Recordset("DescripcionGrupo")
                                  End If
                        End Select
                        
                        '//////////////////////BUSCO LA CUENTA /////////////////////////////////////////////////
                        Me.AdoConsulta.RecordSource = "SELECT  * From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                         Me.AdoConsulta.Recordset("KeyGrupo") = KeyGrupo
                         Me.AdoConsulta.Recordset("DescripcionGrupo") = DescripcionGrupo
                         Me.AdoConsulta.Recordset.Update
                        End If
                        
                        Me.List1.AddItem ("Se Ubico en el grupo: " & DescripcionGrupo)

                   
                   
                   
                   
                   End If
                End If
                
                

         
         
         End If
         i = i + 1
         Me.AdoCuentas.Recordset.MoveNext
        Loop
     End With
     
    If Me.List1.Text = "" Then
      Me.List1.AddItem ("---------O-----------")
    End If
    
    Me.SSTab1.Enabled = True

Exit Sub
TipoErrs:
  MsgBox err.Description
End Sub

Private Sub CmdAuditoria_Click()
'On Error GoTo TipoErrs
Dim CantRegistros As Integer, i As Double, J As Double
Dim MonedaCuenta As String, Fechas1 As Date, CodigoCuenta As String
Dim TasaCambio1 As Double, TasaCambio2 As Double
Dim Debito As Double, Credito As Double, SQL As String, CantRegistros2 As Double, i2 As Double

Me.SSTab1.Enabled = False
Me.List1.Clear

'//////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////Cargo todos los movimientos Contable///////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////

If Me.Option1.Value = True Then
 SQL = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.Credito, Transacciones.Debito, Transacciones.FechaTasas, Cuentas.TipoMoneda FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
Else
 SQL = "SELECT     Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NPeriodo, Transacciones.NTransaccion, " & _
       "Transacciones.NumeroMovimiento, Transacciones.NombreCuenta, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, " & _
      "Transacciones.Clave , Transacciones.TCambio, Transacciones.Credito, Transacciones.Debito, Transacciones.FechaTasas, Cuentas.TipoMoneda " & _
      "FROM         Cuentas INNER JOIN " & _
      "Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas " & _
      "WHERE     (Transacciones.FechaTransaccion BETWEEN '" & Format(Me.DTPFechaIni.Value, "yyyymmdd") & "' And '" & Format(Me.DTPFechaFin.Value, "yyyymmdd") & "') " & _
      "ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento "
 End If



Me.DtaMovimientos.RecordSource = SQL
Me.DtaMovimientos.Refresh
If Not Me.DtaMovimientos.Recordset.EOF Then
 Me.DtaMovimientos.Recordset.MoveLast
 CanRegistros = Me.DtaMovimientos.Recordset.RecordCount
 Me.DtaMovimientos.Recordset.MoveFirst
Else
  MsgBox "No Existen Registros", vbInformation, "Sistema Contable"
  Exit Sub
End If

With Barra
 .Min = 0
 .Max = CanRegistros
 .Value = 0
 i = 1
 Me.LblProceso.Visible = True
 Do While Not Me.DtaMovimientos.Recordset.EOF
  
  If Not IsNull(DtaMovimientos.Recordset("TCambio")) Then
  TasaCambio1 = Format(DtaMovimientos.Recordset("TCambio"), "##,##0.0000")
  Else
  TasaCambio1 = 0
  End If
  MonedaCuenta = Me.DtaMovimientos.Recordset("TipoMoneda")
  .Value = i
  
  DoEvents
  Me.LblProceso.Caption = "Procesando " & i & " de " & CanRegistros
  
'  NumFecha1 = DtaMovimientos.Recordset("FechaTransaccion")
  Fechas1 = DtaMovimientos.Recordset("FechaTransaccion")
  NumeroMovimiento = DtaMovimientos.Recordset("NumeroMovimiento")
   
   Debito = 0
   Credito = 0
'   Me.DtaTransacciones.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.Credito, Transacciones.Debito FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas Where (((Transacciones.FechaTransaccion) = '" & Format(Fechas1, "yyyymmdd") & "') And ((Transacciones.NumeroMovimiento) = " & NumeroMovimiento & ")) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"

    Me.DtaTransacciones.RecordSource = "SELECT CodCuentas, FechaTransaccion, NTransaccion, NumeroMovimiento, Credito, Debito From Transacciones WHERE (FechaTransaccion = '" & Format(Fechas1, "yyyymmdd") & "') AND (NumeroMovimiento = " & NumeroMovimiento & ") ORDER BY FechaTransaccion, NumeroMovimiento "
    Me.DtaTransacciones.Refresh

   
   If Not Me.DtaTransacciones.Recordset.EOF Then
   CantRegistros2 = Me.DtaTransacciones.Recordset.RecordCount
   
   Me.Barra2.Value = 0
   Me.Barra2.Min = 0
   Me.Barra2.Max = CantRegistros2
   Me.Barra2.Value = 0
   Me.Barra2.Visible = True
   i2 = 0
   End If
   
   Do While Not Me.DtaTransacciones.Recordset.EOF
     Debito = Me.DtaTransacciones.Recordset("Debito") + Debito
     Credito = Me.DtaTransacciones.Recordset("Credito") + Credito
     CodigoCuenta = Me.DtaTransacciones.Recordset("CodCuentas")
     
     If BuscaCuentas(CodigoCuenta) = False Then
      
        Me.List1.AddItem ("---------O-----------")
        Me.List1.AddItem ("La Cuenta No Existe, Cuenta No: " & CodigoCuenta & " Transaccion No : " & NumeroMovimiento & "   Fecha Transaccion:   " & CDate(Fechas1))
     
     End If
     
     i2 = i2 + 1
     Me.Barra2.Value = i2
     Me.DtaTransacciones.Recordset.MoveNext
   Loop
  
   If Debito <> Credito Then
   Me.List1.AddItem ("---------O-----------")
   
   Me.List1.AddItem ("Existe Descuadre en el Sistema, Transaccion No : " & NumeroMovimiento & "   Fecha Transaccion:   " & CDate(Fechas1))
   Me.List1.AddItem ("por la Cantidad:" & Debito - Credito)
 
   End If
  
  Me.DtaTasas.RecordSource = "SELECT Tasas.FechaTasas, Tasas.MontoCordobas, Tasas.MontoLibras From Tasas Where (((Tasas.FechaTasas) = '" & Format(Fechas1, "yyyymmdd") & "'))"
  Me.DtaTasas.Refresh
  If Me.DtaTasas.Recordset.EOF Then
   Me.List1.AddItem ("---------O-----------")
   Me.List1.AddItem ("No Existen Tasas para la Transaccion No: " & NumeroMovimiento & " Fecha Tasa: " & CDate(Fechas1))
   TasaCambio2 = 0
  Else
   TasaCambio2 = Me.DtaTasas.Recordset("MontoCordobas")
  End If
  
  Me.DtaIndiceTransaccion.RecordSource = "SELECT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento, IndiceTransaccion.DescripcionMovimiento, IndiceTransaccion.Nperiodo, IndiceTransaccion.Fuente, IndiceTransaccion.TipoMoneda From IndiceTransaccion WHERE (((IndiceTransaccion.FechaTransaccion)='" & Format(Fechas1, "yyyymmdd") & "') AND ((IndiceTransaccion.NumeroMovimiento)=" & NumeroMovimiento & "))"
  Me.DtaIndiceTransaccion.Refresh
  
  If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
    If IsNull(Me.DtaIndiceTransaccion.Recordset("TipoMoneda")) Then
        Me.List1.AddItem ("Valor nulo en Tipo de Moneda, Transaccion No: " & NumeroMovimiento & "  Fecha " & CDate(Fechas1))
    Else
    TipoMoneda = Me.DtaIndiceTransaccion.Recordset("TipoMoneda")
       If TipoMoneda = MonedaCuenta Then
            If Not TasaCambio1 = 1 Then
              Me.List1.AddItem ("---------O-----------")
                  Me.List1.AddItem ("Las Tasas de Cambios en Transacciones no Coinciden,Transaccion No: " & NumeroMovimiento & "  Fecha " & CDate(Fechas1))
              'Me.DtaMovimientos.Recordset.Edit
                 DtaMovimientos.Recordset("TCambio") = 1
              Me.DtaMovimientos.Recordset.Update
            End If
       Else
         Select Case TipoMoneda
          

              Case "Córdobas"
                 If Not TasaCambio1 = Format((1 / TasaCambio2), "##,##0.0000") Then
                   Me.List1.AddItem ("---------O-----------")
                   Me.List1.AddItem ("Las Tasas de Cambios en Transacciones no Coinciden,Transaccion No: " & NumeroMovimiento & " Fecha " & CDate(Fechas1))
                   If Not TasaCambio2 = 0 Then
                    'Me.DtaMovimientos.Recordset.Edit
                     DtaMovimientos.Recordset("TCambio") = 1 / TasaCambio2
                    Me.DtaMovimientos.Recordset.Update
                   End If
                 End If
              Case "Dólares"
                 If Not TasaCambio1 = TasaCambio2 Then
                   Me.List1.AddItem ("---------O-----------")
                   Me.List1.AddItem ("Las Tasas de Cambios en Transacciones no Coinciden,Transaccion No: " & NumeroMovimiento & " Fecha " & CDate(Fechas1))
                   If Not TasaCambio2 = 0 Then
                    'Me.DtaMovimientos.Recordset.Edit
                     DtaMovimientos.Recordset("TCambio") = TasaCambio2
                    Me.DtaMovimientos.Recordset.Update
                   End If
                 End If
         End Select
       End If
  End If
  
  
  
  Else
    Me.List1.AddItem ("---------O-----------")
    Me.List1.AddItem ("No Se encuntra la Transaccion en la Tabla Indices, MovimientoNo" & NumeroMovimiento & " Fecha: " & CDate(Fechas1))
  End If
 

  
  Me.DtaMovimientos.Recordset.MoveNext
  i = i + 1

 Loop
End With

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////VERIFICO LOS PERIODOS////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
For J = 1 To 3
  SQL = "SELECT * From Periodos Where (NumeroTabla = " & J & ") ORDER BY NPeriodo"
  Me.AdoPeriodos.RecordSource = SQL
  Me.AdoPeriodos.Refresh
  
  If Not Me.AdoPeriodos.Recordset.EOF Then
     Me.AdoPeriodos.Recordset.MoveLast
     CantRegistros = Me.AdoPeriodos.Recordset.RecordCount
     Me.AdoPeriodos.Recordset.MoveFirst
     If CantRegistros < 12 Then
      Me.List1.AddItem ("---------O-----------")
      Me.List1.AddItem ("ERROR con la tabla Periodos: NO ESTAN COMPLETOS LOS 12 MESES EN EL AÑO: " & J)
     End If
     
    
     With Barra
          .Value = 0
          .Min = 0
          .Max = CantRegistros
          
          i = 1
          Me.LblProceso.Visible = True
          Do While Not Me.AdoPeriodos.Recordset.EOF
            If Me.AdoPeriodos.Recordset("Periodo") <> i Then
                Me.List1.AddItem ("---------O-----------")
                Me.List1.AddItem ("ERROR con la tabla Periodos: LOS PERIODOS NO SON CONSECUTIVOS PERIODO " & i)
                Me.AdoPeriodos.Recordset("Periodo") = i
                Me.AdoPeriodos.Recordset.Update
                Me.List1.AddItem ("SE HA REPARADO EL PROBLEMA DEL PERIODO: " & i)
            End If
            
            .Value = i
            
            DoEvents
            Me.LblProceso.Caption = "Procesando " & i & " de " & CantRegistros
            i = i + 1
            Me.AdoPeriodos.Recordset.MoveNext
          Loop
     End With
  Else
    Me.List1.AddItem ("---------O-----------")
    Me.List1.AddItem ("ERROR con la tabla Periodos: NO SE ENCONTRO EL AÑO: " & J)
  End If
 
Next



If Me.List1.Text = "" Then

     Me.List1.AddItem ("---------O-----------")
End If

Me.SSTab1.Enabled = True

Exit Sub
TipoErrs:
  MsgBox err.Description
End Sub

Private Sub CmdConvertir_Click()
Dim Registros As Double, i As Double, Contador As Double
Dim NTransaccion As Double, Fecha As Date, FechaIni As String, FechaFin As String, NumeroMovimiento As Double
Dim Respuesta As Double

        Me.SSTab1.Enabled = False
        Me.CmdConvertir.Enabled = False
        Me.CmbTipoMoneda.Enabled = False
            If Me.CmbTipoMoneda.Text = "" Then
              MsgBox "Tiene que Seleccionar una Moneda", vbCritical, "Sistema Contable"
              Me.SSTab1.Enabled = True
              Me.CmdConvertir.Enabled = True
              Me.CmbTipoMoneda.Enabled = True
              Exit Sub
            End If
            

            Me.List1.Clear
            
            If Me.OptConvertirTodo.Value = True Then
              '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE TODAS LAS TRANSACCIONES EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              
              Me.DtaIndiceTransaccion.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion ORDER BY Nperiodo, NumeroMovimiento"
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                 If Not Me.DtaIndiceTransaccion.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text Then
                    Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                    NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                                    Me.DtaIndiceTransaccion.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text
                                    Me.DtaIndiceTransaccion.Recordset.Update
                   Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 End If
              

              
                
                
              
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop
            ElseIf Me.OptConvertirRango.Value = True Then
              '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE CON UN RAGO DE FECHAS LAS TRANSACCIONES EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              
              FechaIni = Format(Me.DTFechaIni.Value, "yyyy-mm-dd")
              FechaFin = Format(Me.DTFechaFin.Value, "yyyy-mm-dd")
              
              Me.DtaIndiceTransaccion.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Nperiodo, NumeroMovimiento"
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                 If Not Me.DtaIndiceTransaccion.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text Then
                    Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                    NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                                    Me.DtaIndiceTransaccion.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text
                                    Me.DtaIndiceTransaccion.Recordset.Update
                   Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 End If
              

              
                
                
              
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop
            
            
            
            
            
            ElseIf Me.OptConvertirTransaccion.Value = True Then
            
              If Me.TxtTransaccion.Text = "" Then
                MsgBox "Se necesita un Numero de Transaccion", vbCritical, "Sistema Contable"
                Me.SSTab1.Enabled = True
                 Me.CmdConvertir.Enabled = True
                 Me.CmbTipoMoneda.Enabled = True
                Exit Sub
              End If
              
              If Not IsNumeric(Me.TxtTransaccion.Text) Then
                 MsgBox "El Numero Transaccion es de Tipo: Numerico", vbCritical, "Sistema Contable"
                 Me.SSTab1.Enabled = True
                 Me.CmdConvertir.Enabled = True
                 Me.CmbTipoMoneda.Enabled = True
                 Exit Sub
              End If
              
              
              
             '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE CON UNA TRANSACCION ESPECIFICA EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              FechaIni = Format(Me.DTFechaIni.Value, "yyyy-mm-dd")
              NumeroMovimiento = Me.TxtTransaccion.Text
              
              Me.DtaIndiceTransaccion.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion WHERE (NumeroMovimiento = " & NumeroMovimiento & ") AND (FechaTransaccion = CONVERT(DATETIME, '" & FechaIni & "', 102)) ORDER BY Nperiodo, NumeroMovimiento"
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                 If Not Me.DtaIndiceTransaccion.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text Then
                    Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                    NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                                    Me.DtaIndiceTransaccion.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text
                                    Me.DtaIndiceTransaccion.Recordset.Update
                   Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 End If
              

              
                
                
              
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop

              
            
            End If
            
            Me.SSTab1.Enabled = True
  
  MsgBox "Se Modificaron: " & Contador, vbInformation, "Sistema Contable"
  Me.CmbTipoMoneda.Enabled = True
  Me.CmdConvertir.Enabled = True
  
 If Not Contador = 0 Then
    Respuesta = MsgBox("Es Necesario Ejecutar el Auditor, ¿Desea Ejecutarlo?", vbYesNo, "Sistema Contable")
    If Respuesta = 6 Then
       Me.SSTab1.Tab = 0
       If Me.OptConvertirTodo.Value = True Then
         Me.Option1.Value = True
         CmdAuditoria_Click
       ElseIf Me.OptConvertirRango.Value = True Then
         Me.Option2.Value = True
         Me.DTPFechaIni.Value = Me.DTFechaIni.Value
         Me.DTPFechaFin.Value = Me.DTFechaFin.Value
         CmdAuditoria_Click
       ElseIf Me.OptConvertirTransaccion.Value = True Then
         Me.Option2.Value = True
         Me.DTPFechaIni.Value = Me.DTFechaIni.Value
         Me.DTPFechaFin.Value = Me.DTFechaIni.Value
         CmdAuditoria_Click
       End If
    End If
 End If
End Sub

Private Sub CmdConvertirMonto_Click()
Dim Registros As Double, i As Double, Contador As Double
Dim NTransaccion As Double, Fecha As Date, FechaIni As String, FechaFin As String, NumeroMovimiento As Double
Dim Respuesta As Double, Cod As Variant

        Me.SSTab1.Enabled = False
        Me.CmdConvertir.Enabled = False
        Me.CmbTipoMoneda.Enabled = False
            If Me.CmbTipoMoneda.Text = "" Then
              MsgBox "Tiene que Seleccionar una Moneda", vbCritical, "Sistema Contable"
              Me.SSTab1.Enabled = True
              Me.CmdConvertir.Enabled = True
              Me.CmbTipoMoneda.Enabled = True
              Exit Sub
            End If
            

            Me.List1.Clear
            
            If Me.OptConvertirTodo.Value = True Then
              '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE TODAS LAS TRANSACCIONES EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              
              Me.DtaIndiceTransaccion.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion ORDER BY Nperiodo, NumeroMovimiento"
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                
                    Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                    NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                    Cod = ConvertirMovimiento(NTransaccion, Fecha, Me.CmbTipoMoneda.Text)
                    Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                    
                    Contador = Contador + 1
                
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop
            ElseIf Me.OptConvertirRango.Value = True Then
              '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE CON UN RAGO DE FECHAS LAS TRANSACCIONES EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              
              FechaIni = Format(Me.DTFechaIni.Value, "yyyy-mm-dd")
              FechaFin = Format(Me.DTFechaFin.Value, "yyyy-mm-dd")
              
              Me.DtaIndiceTransaccion.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Nperiodo, NumeroMovimiento"
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                    Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                    NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                    Cod = ConvertirMovimiento(NTransaccion, Fecha, Me.CmbTipoMoneda.Text)
                    Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop
            
            
            
            
            
            ElseIf Me.OptConvertirTransaccion.Value = True Then
            
              If Me.TxtTransaccion.Text = "" Then
                MsgBox "Se necesita un Numero de Transaccion", vbCritical, "Sistema Contable"
                Me.SSTab1.Enabled = True
                 Me.CmdConvertir.Enabled = True
                 Me.CmbTipoMoneda.Enabled = True
                Exit Sub
              End If
              
              If Not IsNumeric(Me.TxtTransaccion.Text) Then
                 MsgBox "El Numero Transaccion es de Tipo: Numerico", vbCritical, "Sistema Contable"
                 Me.SSTab1.Enabled = True
                 Me.CmdConvertir.Enabled = True
                 Me.CmbTipoMoneda.Enabled = True
                 Exit Sub
              End If
              
              
              
             '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE CON UNA TRANSACCION ESPECIFICA EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              FechaIni = Format(Me.DTFechaIni.Value, "yyyy-mm-dd")
              NumeroMovimiento = Me.TxtTransaccion.Text
              
              Me.DtaIndiceTransaccion.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion WHERE (NumeroMovimiento = " & NumeroMovimiento & ") AND (FechaTransaccion = CONVERT(DATETIME, '" & FechaIni & "', 102)) ORDER BY Nperiodo, NumeroMovimiento"
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                 If Not Me.DtaIndiceTransaccion.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text Then
                    Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                    NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                    Cod = ConvertirMovimiento(NTransaccion, Fecha, Me.CmbTipoMoneda.Text)
                   Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 End If
              

              
                
                
              
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop

              
            
            End If
            
            Me.SSTab1.Enabled = True
  
  MsgBox "Se Modificaron: " & Contador, vbInformation, "Sistema Contable"
  Me.CmbTipoMoneda.Enabled = True
  Me.CmdConvertir.Enabled = True

End Sub

Private Sub CmdCorregir_Click()
Dim Registros As Double, i As Double, Contador As Double
Dim NTransaccion As Double, Fecha As Date, FechaIni As String, FechaFin As String, NumeroMovimiento As Double
Dim Respuesta As Double

        Me.SSTab1.Enabled = False
        Me.CmdConvertir.Enabled = False
        Me.CmbTipoMoneda.Enabled = False
            If Me.CmbTipoMoneda.Text = "" Then
              MsgBox "Tiene que Seleccionar una Moneda", vbCritical, "Sistema Contable"
              Me.SSTab1.Enabled = True
              Me.CmdConvertir.Enabled = True
              Me.CmbTipoMoneda.Enabled = True
              Exit Sub
            End If
            

            Me.List1.Clear
            
            If Me.OptConvertirTodo.Value = True Then
              '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE TODAS LAS TRANSACCIONES EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              If Me.CmbTipoMoneda.Text = "Córdobas" Then
                Me.DtaIndiceTransaccion.RecordSource = "SELECT DISTINCT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento FROM IndiceTransaccion INNER JOIN Transacciones ON IndiceTransaccion.FechaTransaccion = Transacciones.FechaTransaccion AND IndiceTransaccion.NumeroMovimiento = Transacciones.NumeroMovimiento Where (Transacciones.TCambio = 1) ORDER BY IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento"
              Else
                Me.DtaIndiceTransaccion.RecordSource = "SELECT DISTINCT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento FROM IndiceTransaccion INNER JOIN Transacciones ON IndiceTransaccion.FechaTransaccion = Transacciones.FechaTransaccion AND IndiceTransaccion.NumeroMovimiento = Transacciones.NumeroMovimiento Where (Transacciones.TCambio <> 1) ORDER BY IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento"
              End If
              
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
              
                 Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                 NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                 Me.AdoIndice.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion WHERE  (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NTransaccion & ")"
                 Me.AdoIndice.Refresh
                 If Not Me.AdoIndice.Recordset.EOF Then

                                    Me.AdoIndice.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text
                                    Me.AdoIndice.Recordset.Update
                   Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 End If
              

              
                
                
              
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop
            ElseIf Me.OptConvertirRango.Value = True Then
              '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE CON UN RAGO DE FECHAS LAS TRANSACCIONES EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              
              FechaIni = Format(Me.DTFechaIni.Value, "yyyy-mm-dd")
              FechaFin = Format(Me.DTFechaFin.Value, "yyyy-mm-dd")
              
              If Me.CmbTipoMoneda.Text = "Córdobas" Then
              Me.DtaIndiceTransaccion.RecordSource = "SELECT DISTINCT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento FROM IndiceTransaccion INNER JOIN Transacciones ON IndiceTransaccion.FechaTransaccion = Transacciones.FechaTransaccion AND IndiceTransaccion.NumeroMovimiento = Transacciones.NumeroMovimiento " & _
                                                     "WHERE (Transacciones.TCambio = 1) AND (IndiceTransaccion.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) ORDER BY IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento"
              Else
              Me.DtaIndiceTransaccion.RecordSource = "SELECT DISTINCT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento FROM IndiceTransaccion INNER JOIN Transacciones ON IndiceTransaccion.FechaTransaccion = Transacciones.FechaTransaccion AND IndiceTransaccion.NumeroMovimiento = Transacciones.NumeroMovimiento " & _
                                                     "WHERE (Transacciones.TCambio <> 1) AND (IndiceTransaccion.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) ORDER BY IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento"
              End If
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                 Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                 NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                 Me.AdoIndice.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion WHERE  (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NTransaccion & ")"
                 Me.AdoIndice.Refresh
                 If Not Me.AdoIndice.Recordset.EOF Then

                                    Me.AdoIndice.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text
                                    Me.AdoIndice.Recordset.Update
                   Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 End If
              

              
                
                
              
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop
            
            
            
            
            
            ElseIf Me.OptConvertirTransaccion.Value = True Then
            
              If Me.TxtTransaccion.Text = "" Then
                MsgBox "Se necesita un Numero de Transaccion", vbCritical, "Sistema Contable"
                Me.SSTab1.Enabled = True
                 Me.CmdConvertir.Enabled = True
                 Me.CmbTipoMoneda.Enabled = True
                Exit Sub
              End If
              
              If Not IsNumeric(Me.TxtTransaccion.Text) Then
                 MsgBox "El Numero Transaccion es de Tipo: Numerico", vbCritical, "Sistema Contable"
                 Me.SSTab1.Enabled = True
                 Me.CmdConvertir.Enabled = True
                 Me.CmbTipoMoneda.Enabled = True
                 Exit Sub
              End If
              
              
              
             '///////////////////////////////////////////////////////////////////////////////////
              ''////////////////ESTA OPCION CONVIERTE CON UNA TRANSACCION ESPECIFICA EN UNA MONEDA/////
              '///////////////////////////////////////////////////////////////////////////////////
              FechaIni = Format(Me.DTFechaIni.Value, "yyyy-mm-dd")
              NumeroMovimiento = Me.TxtTransaccion.Text
              
              Me.DtaIndiceTransaccion.RecordSource = "SELECT DISTINCT IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento FROM IndiceTransaccion INNER JOIN Transacciones ON IndiceTransaccion.FechaTransaccion = Transacciones.FechaTransaccion AND IndiceTransaccion.NumeroMovimiento = Transacciones.NumeroMovimiento " & _
                                                     "WHERE (Transacciones.TCambio <> 1) AND (IndiceTransaccion.FechaTransaccion = CONVERT(DATETIME, '" & FechaIni & "', 102)) AND (IndiceTransaccion.NumeroMovimiento = " & NumeroMovimiento & ") ORDER BY IndiceTransaccion.FechaTransaccion, IndiceTransaccion.NumeroMovimiento"
              '"SELECT IndiceTransaccion.* From IndiceTransaccion WHERE (NumeroMovimiento = " & NumeroMovimiento & ") AND (FechaTransaccion = CONVERT(DATETIME, '" & FechaIni & "', 102)) ORDER BY Nperiodo, NumeroMovimiento"
              Me.DtaIndiceTransaccion.Refresh
              If Not Me.DtaIndiceTransaccion.Recordset.EOF Then
                Registros = Me.DtaIndiceTransaccion.Recordset.RecordCount
                Me.Barra.Max = Registros
                Me.Barra.Min = 0
                i = 0
                Contador = 0
                Me.LblProceso.Visible = True
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
              
              Else
               MsgBox "No Existen Transacciones", vbCritical, "Sistema Contable"
               Me.SSTab1.Enabled = True
               Me.CmdConvertir.Enabled = True
               Me.CmbTipoMoneda.Enabled = True
               Exit Sub
              End If
              Do While Not Me.DtaIndiceTransaccion.Recordset.EOF
                 
                 Fecha = Me.DtaIndiceTransaccion.Recordset("FechaTransaccion")
                 NTransaccion = Me.DtaIndiceTransaccion.Recordset("NumeroMovimiento")
                 Me.AdoIndice.RecordSource = "SELECT IndiceTransaccion.* From IndiceTransaccion WHERE  (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NTransaccion & ")"
                 Me.AdoIndice.Refresh
                 If Not Me.AdoIndice.Recordset.EOF Then

                                    Me.AdoIndice.Recordset("TipoMoneda") = Me.CmbTipoMoneda.Text
                                    Me.AdoIndice.Recordset.Update
                   Me.List1.AddItem ("Se Modifico Tran# " & NTransaccion & " Fecha: " & Fecha)
                   Contador = Contador + 1
                 End If
              

              
                
                
              
                DoEvents
                i = i + 1
                Me.Barra.Value = i
                Me.LblProceso.Caption = "Procesando " & i & de & " de " & Registros
                Me.DtaIndiceTransaccion.Recordset.MoveNext
              Loop

              
            
            End If
            
            Me.SSTab1.Enabled = True
  
  MsgBox "Se Modificaron: " & Contador, vbInformation, "Sistema Contable"
  Me.CmbTipoMoneda.Enabled = True
  Me.CmdConvertir.Enabled = True
  
 If Not Contador = 0 Then
    Respuesta = MsgBox("Es Necesario Ejecutar el Auditor, ¿Desea Ejecutarlo?", vbYesNo, "Sistema Contable")
    If Respuesta = 6 Then
       Me.SSTab1.Tab = 0
       If Me.OptConvertirTodo.Value = True Then
         Me.Option1.Value = True
         CmdAuditoria_Click
       ElseIf Me.OptConvertirRango.Value = True Then
         Me.Option2.Value = True
         Me.DTPFechaIni.Value = Me.DTFechaIni.Value
         Me.DTPFechaFin.Value = Me.DTFechaFin.Value
         CmdAuditoria_Click
       ElseIf Me.OptConvertirTransaccion.Value = True Then
         Me.Option2.Value = True
         Me.DTPFechaIni.Value = Me.DTFechaIni.Value
         Me.DTPFechaFin.Value = Me.DTFechaIni.Value
         CmdAuditoria_Click
       End If
    End If
 End If
End Sub

Private Sub CmdReparar_Click()
FrmReparar.Show 1
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdTasas_Click()
Dim SQL As String, i As Double, CantRegistros As Double

SQL = "SELECT  * From Tasas"
Me.AdoTasas.RecordSource = SQL
Me.AdoTasas.Refresh

Me.AdoTasas.Recordset.MoveLast
CantRegistros = Me.AdoTasas.Recordset.RecordCount
Me.AdoTasas.Recordset.MoveFirst

    With Barra
         .Min = 0
         .Max = CantRegistros
         .Value = 0
         i = 1
        Me.LblProceso.Visible = True
        Do While Not Me.AdoTasas.Recordset.EOF
            .Value = i
            DoEvents
            Me.LblProceso.Caption = "Procesando " & i & " de " & CantRegistros
            
            Me.AdoTasas.Recordset("MontoCordobas") = Format(Me.AdoTasas.Recordset("MontoCordobas"), "##,##0.00")
            Me.AdoTasas.Recordset.Update
            
          Me.AdoTasas.Recordset.MoveNext
          i = i + 1
        Loop
    End With

 If Not i = 0 Then
    Respuesta = MsgBox("Es Necesario Ejecutar el Auditor, ¿Desea Ejecutarlo?", vbYesNo, "Sistema Contable")
    If Respuesta = 6 Then
       Me.SSTab1.Tab = 0
       CmdAuditoria_Click
    End If
 End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Me.SSTab1.BackColor = RGB(219, 226, 242)


Me.DTPFechaFin.Value = Format(Now, "dd/mm/yyyy")
Me.DTPFechaIni.Value = Format(Now, "dd/mm/yyyy")
Me.DTFechaFin.Value = Format(Now, "dd/mm/yyyy")
Me.DTFechaIni.Value = Format(Now, "dd/mm/yyyy")

With Me.AdoPeriodos
 .ConnectionString = Conexion
End With

With Me.AdoTasas
 .ConnectionString = Conexion
End With

With Me.AdoIndice
 .ConnectionString = Conexion
End With


With Me.AdoCuentas
 .ConnectionString = Conexion
End With

With Me.AdoGrupos
 .ConnectionString = Conexion
End With

With Me.AdoConsulta
 .ConnectionString = Conexion
End With

With Me.DtaMovimientos

 .ConnectionString = Conexion
End With
'
With Me.DtaTransacciones

 .ConnectionString = Conexion
End With
'
With Me.DtaIndiceTransaccion

 .ConnectionString = Conexion
End With
'
With Me.DtaTasas

 .ConnectionString = Conexion
End With
End Sub

Private Sub OptConvertirRango_Click()
 Me.LblDesde.Visible = True
 Me.LblHasta.Visible = True
 Me.LblHasta.Caption = "Hasta"
 Me.DTFechaFin.Visible = True
 Me.DTFechaIni.Visible = True
 Me.TxtTransaccion.Visible = False
End Sub

Private Sub OptConvertirTodo_Click()
 Me.LblDesde.Visible = False
 Me.LblHasta.Visible = False
 Me.DTFechaFin.Visible = False
 Me.DTFechaIni.Visible = False
 Me.TxtTransaccion.Visible = False
End Sub

Private Sub OptConvertirTransaccion_Click()
Me.LblDesde.Visible = True
 Me.LblHasta.Visible = True
 Me.LblHasta.Caption = "Tran#"
 Me.DTFechaFin.Visible = False
 Me.DTFechaIni.Visible = True
 Me.TxtTransaccion.Visible = True
End Sub

Private Sub Option1_Click()
 Me.DTPFechaFin.Enabled = False
 Me.DTPFechaIni.Enabled = False
End Sub

Private Sub Option2_Click()
 Me.DTPFechaFin.Enabled = True
 Me.DTPFechaIni.Enabled = True
End Sub

Private Sub Option3_Click()
 Me.LblDesde.Visible = True
 Me.LblHasta.Visible = True
 Me.LblHasta.Caption = "Tran#"
 Me.DTFechaFin.Visible = False
 Me.DTFechaIni.Visible = True
 Me.TxtTransaccion.Visible = True
End Sub
