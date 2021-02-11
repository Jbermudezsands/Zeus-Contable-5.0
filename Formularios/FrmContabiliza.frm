VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmContabilizaFacturacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilizacion del Sistema de Factuacion"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoProcesosFacturacion 
      Height          =   375
      Left            =   8280
      Top             =   8280
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "AdoProcesosFacturacion"
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
   Begin MSAdodcLib.Adodc AdoBuscaFacturacion 
      Height          =   375
      Left            =   720
      Top             =   8280
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
      Caption         =   "AdoBuscaFacturacion"
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
   Begin MSAdodcLib.Adodc AdoConsultaFactura 
      Height          =   375
      Left            =   6840
      Top             =   8280
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
      Caption         =   "AdoConsultaFactura"
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
   Begin MSAdodcLib.Adodc AdoDetalleFactura 
      Height          =   375
      Left            =   7800
      Top             =   8880
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
      Caption         =   "AdoDetalleFactura"
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
   Begin MSAdodcLib.Adodc AdoProcesos 
      Height          =   375
      Left            =   1200
      Top             =   8880
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
      Caption         =   "AdoProcesos"
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
      Left            =   720
      Top             =   8880
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoConsulta"
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
   Begin MSAdodcLib.Adodc AdoCompras 
      Height          =   375
      Left            =   7800
      Top             =   8760
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoCompras"
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
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   375
      Left            =   3840
      Top             =   8640
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
      Caption         =   "AdoDatosEmpresa"
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
   Begin MSAdodcLib.Adodc AdoFacturacion 
      Height          =   450
      Left            =   600
      Top             =   8280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoFacturacion"
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
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   11520
      TabIndex        =   7
      Top             =   7560
      Width           =   1215
      _Version        =   786432
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   4
      TabHeight       =   520
      TabCaption(0)   =   "Facturacion"
      TabPicture(0)   =   "FrmContabiliza.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGridFacturacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GroupBox1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Compras"
      TabPicture(1)   =   "FrmContabiliza.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GroupBox2"
      Tab(1).Control(1)=   "TDBGridCompras"
      Tab(1).Control(2)=   "PushButton2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Cuentas x Cobrar y Pagar"
      TabPicture(2)   =   "FrmContabiliza.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GroupBox3"
      Tab(2).Control(1)=   "TDBGridCuentas"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Planilla Leche"
      TabPicture(3)   =   "FrmContabiliza.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GroupBox4"
      Tab(3).Control(1)=   "TDGridPlanillaLeche"
      Tab(3).ControlCount=   2
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   16
         Top             =   780
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   8916
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.CheckBox ChkDescripcionCompra 
            Caption         =   "Anexar Desc Prod Proveedor"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   4560
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   345
            Left            =   840
            TabIndex        =   30
            Top             =   3720
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   345
            Left            =   600
            TabIndex        =   17
            Top             =   2040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptCompras 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Compras"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptDevolucionCompra 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Devolucion de Compras"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptTransferenciaRecibida 
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Transferencia Recibida"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   345
            Left            =   600
            TabIndex        =   21
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton RadioButton4 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Pago a Proveedores"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdContabilizarCompras 
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   3240
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Contabilizar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdConsultarCompra 
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   2880
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Consultar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptCuenta 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cuenta"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdRecepcion 
            Height          =   375
            Left            =   360
            TabIndex        =   50
            Top             =   3240
            Visible         =   0   'False
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Contabilizar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkCheques 
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   4080
            Visible         =   0   'False
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Crear Cheque x Proveedor"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin VB.Label LblFechaCompra 
            Caption         =   "Feha:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   3720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   2400
            Width           =   495
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   9340
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.CheckBox ChkCtaCtoProducto 
            Caption         =   "Utilizar Cta Cto Prodto"
            Height          =   375
            Left            =   120
            TabIndex        =   74
            Top             =   5040
            Width           =   2055
         End
         Begin VB.CheckBox ChkDescripcion 
            Caption         =   "Anexar Desc Prod Clientes"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   4600
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Left            =   720
            TabIndex        =   10
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptFacturacion 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Facturacion"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptDevolucion 
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Devolucion de Ventas"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptTransferencia 
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1800
            Visible         =   0   'False
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Transferencia Enviada"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   345
            Left            =   720
            TabIndex        =   12
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptRecibos 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Recibos de Caja"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdContabilizar 
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   3720
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Contabilizar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdConsultar 
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   3240
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Consultar"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   345
            Left            =   840
            TabIndex        =   28
            Top             =   4200
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptSalidaBodega 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Salida Bodega"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label LblFecha 
            Caption         =   "Feha:"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   4080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   2280
            Width           =   495
         End
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridFacturacion 
         Bindings        =   "FrmContabiliza.frx":0070
         Height          =   5175
         Left            =   2520
         TabIndex        =   8
         Top             =   900
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9128
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   1
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha"
         Columns(0).DataField=   "Fecha_Factura"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Factura No"
         Columns(1).DataField=   "DescripcionMovimiento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nombre Cliente"
         Columns(2).DataField=   "VoucherNo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Sub Total"
         Columns(3).DataField=   "ChequeNo"
         Columns(3).DataWidth=   50
         Columns(3).NumberFormat=   "Standard"
         Columns(3).EditMask=   "##,##.##"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descuento"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4).EditMask=   "##,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "IVA"
         Columns(5).DataField=   "TCambio"
         Columns(5).NumberFormat=   "Standard"
         Columns(5).EditMask=   "##,##0.00"
         Columns(5).EditMaskUpdate=   -1  'True
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Neto Pagar"
         Columns(6).DataField=   "Debito"
         Columns(6).NumberFormat=   "Standard"
         Columns(6).EditMask=   "##,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   80
         Columns(7)._MaxComboItems=   5
         Columns(7).ValueItems(0)._DefaultItem=   0
         Columns(7).ValueItems(0).Value=   "0"
         Columns(7).ValueItems(0).Value.vt=   8
         Columns(7).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
         Columns(7).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
         Columns(7).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
         Columns(7).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
         Columns(7).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
         Columns(7).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
         Columns(7).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
         Columns(7).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
         Columns(7).ValueItems(0).DisplayValue.vt=   9
         Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems(1)._DefaultItem=   0
         Columns(7).ValueItems(1).Value=   "-1"
         Columns(7).ValueItems(1).Value.vt=   8
         Columns(7).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
         Columns(7).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
         Columns(7).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
         Columns(7).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
         Columns(7).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
         Columns(7).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
         Columns(7).ValueItems(1).DisplayValue.vt=   9
         Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems.Count=   2
         Columns(7).Caption=   "Contabilizar"
         Columns(7).DataField=   "Conciliada"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Movimientos"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Button=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8194"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1693"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=8194"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=1931"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1852"
         Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=8194"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=1931"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1852"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=8194"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(7).Width=1931"
         Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1852"
         Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   3
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(8)   =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&H0&,.bold=-1,.fontsize=825"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HD2D7E6&,.fgcolor=&HA00000&"
         _StyleDefs(24)  =   ":id=22,.bold=-1,.fontsize=1275,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(25)  =   ":id=22,.charset=0"
         _StyleDefs(26)  =   ":id=22,.fontname=Pristina"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HD2D7E6&,.fgcolor=&H0&,.bold=-1"
         _StyleDefs(28)  =   ":id=14,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(71)  =   "Named:id=33:Normal"
         _StyleDefs(72)  =   ":id=33,.parent=0"
         _StyleDefs(73)  =   "Named:id=34:Heading"
         _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   ":id=34,.wraptext=-1"
         _StyleDefs(76)  =   "Named:id=35:Footing"
         _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   "Named:id=36:Selected"
         _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=37:Caption"
         _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(82)  =   "Named:id=38:HighlightRow"
         _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(84)  =   "Named:id=39:EvenRow"
         _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(86)  =   "Named:id=40:OddRow"
         _StyleDefs(87)  =   ":id=40,.parent=33"
         _StyleDefs(88)  =   "Named:id=41:RecordSelector"
         _StyleDefs(89)  =   ":id=41,.parent=34"
         _StyleDefs(90)  =   "Named:id=42:FilterBar"
         _StyleDefs(91)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridCompras 
         Bindings        =   "FrmContabiliza.frx":008D
         Height          =   4815
         Left            =   -72480
         TabIndex        =   27
         Top             =   900
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8493
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha"
         Columns(0).DataField=   "Fecha_Factura"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Factura No"
         Columns(1).DataField=   "DescripcionMovimiento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nombre Cliente"
         Columns(2).DataField=   "VoucherNo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Sub Total"
         Columns(3).DataField=   "ChequeNo"
         Columns(3).DataWidth=   50
         Columns(3).NumberFormat=   "Standard"
         Columns(3).EditMask=   "##,##.##"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descuento"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4).EditMask=   "##,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "IVA"
         Columns(5).DataField=   "TCambio"
         Columns(5).NumberFormat=   "Standard"
         Columns(5).EditMask=   "##,##0.00"
         Columns(5).EditMaskUpdate=   -1  'True
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Neto Pagar"
         Columns(6).DataField=   "Debito"
         Columns(6).NumberFormat=   "Standard"
         Columns(6).EditMask=   "##,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   80
         Columns(7)._MaxComboItems=   5
         Columns(7).ValueItems(0)._DefaultItem=   0
         Columns(7).ValueItems(0).Value=   "0"
         Columns(7).ValueItems(0).Value.vt=   8
         Columns(7).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
         Columns(7).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
         Columns(7).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
         Columns(7).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
         Columns(7).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
         Columns(7).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
         Columns(7).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
         Columns(7).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
         Columns(7).ValueItems(0).DisplayValue.vt=   9
         Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems(1)._DefaultItem=   0
         Columns(7).ValueItems(1).Value=   "-1"
         Columns(7).ValueItems(1).Value.vt=   8
         Columns(7).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
         Columns(7).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
         Columns(7).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
         Columns(7).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
         Columns(7).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
         Columns(7).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
         Columns(7).ValueItems(1).DisplayValue.vt=   9
         Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems.Count=   2
         Columns(7).Caption=   "Contabilizar"
         Columns(7).DataField=   "Conciliada"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Movimientos"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8194"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1693"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8194"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1931"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1852"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8194"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1931"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1852"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8194"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=1931"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1852"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   3
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(8)   =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&H0&,.bold=-1,.fontsize=825"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HD2D7E6&,.fgcolor=&HA00000&"
         _StyleDefs(24)  =   ":id=22,.bold=-1,.fontsize=1275,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(25)  =   ":id=22,.charset=0"
         _StyleDefs(26)  =   ":id=22,.fontname=Pristina"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HD2D7E6&,.fgcolor=&H0&,.bold=-1"
         _StyleDefs(28)  =   ":id=14,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(71)  =   "Named:id=33:Normal"
         _StyleDefs(72)  =   ":id=33,.parent=0"
         _StyleDefs(73)  =   "Named:id=34:Heading"
         _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   ":id=34,.wraptext=-1"
         _StyleDefs(76)  =   "Named:id=35:Footing"
         _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   "Named:id=36:Selected"
         _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=37:Caption"
         _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(82)  =   "Named:id=38:HighlightRow"
         _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(84)  =   "Named:id=39:EvenRow"
         _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(86)  =   "Named:id=40:OddRow"
         _StyleDefs(87)  =   ":id=40,.parent=33"
         _StyleDefs(88)  =   "Named:id=41:RecordSelector"
         _StyleDefs(89)  =   ":id=41,.parent=34"
         _StyleDefs(90)  =   "Named:id=42:FilterBar"
         _StyleDefs(91)  =   ":id=42,.parent=33"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   36
         Top             =   780
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   8916
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin MSComCtl2.DTPicker DTPicker7 
            Height          =   345
            Left            =   720
            TabIndex        =   37
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptNotaDebito 
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notas de Debito"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptNotaCredito 
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notas de Credito"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker8 
            Height          =   345
            Left            =   720
            TabIndex        =   40
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.PushButton CmdContabilizarNotas 
            Height          =   375
            Left            =   360
            TabIndex        =   41
            Top             =   3600
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Contabilizar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdConsultaNota 
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   3240
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Consultar"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker9 
            Height          =   345
            Left            =   840
            TabIndex        =   43
            Top             =   4080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptNotaDebitoProveedor 
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Notas de Debito Proveedores"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptNotaCreditoProveedor 
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Notas de Credito Proveedores"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptPlanillaProductor 
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1920
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Planilla Productor"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkProductorCheque 
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   4560
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Crear Cheque x Proveedor"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin VB.Label Label7 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Feha:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   4200
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridCuentas 
         Bindings        =   "FrmContabiliza.frx":00A6
         Height          =   4935
         Left            =   -72480
         TabIndex        =   47
         Top             =   900
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8705
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   1
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Numero"
         Columns(0).DataField=   "Numero_Nota"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Fecha"
         Columns(1).DataField=   "Fecha_Nota"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Moneda"
         Columns(2).DataField=   "MonedaNota"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cliente"
         Columns(3).DataField=   "Nombre_Cliente"
         Columns(3).DataWidth=   50
         Columns(3).NumberFormat=   "Standard"
         Columns(3).EditMask=   "##,##.##"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descripcion"
         Columns(4).DataField=   "Descripcion"
         Columns(4).NumberFormat=   "Standard"
         Columns(4).EditMask=   "##,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Factura"
         Columns(5).DataField=   "Numero_Factura"
         Columns(5).EditMask=   "##,##0.00"
         Columns(5).EditMaskUpdate=   -1  'True
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Monto"
         Columns(6).DataField=   "Monto"
         Columns(6).NumberFormat=   "Standard"
         Columns(6).EditMask=   "##,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   80
         Columns(7)._MaxComboItems=   5
         Columns(7).ValueItems(0)._DefaultItem=   0
         Columns(7).ValueItems(0).Value=   "0"
         Columns(7).ValueItems(0).Value.vt=   8
         Columns(7).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
         Columns(7).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
         Columns(7).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
         Columns(7).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
         Columns(7).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
         Columns(7).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
         Columns(7).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
         Columns(7).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
         Columns(7).ValueItems(0).DisplayValue.vt=   9
         Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems(1)._DefaultItem=   0
         Columns(7).ValueItems(1).Value=   "-1"
         Columns(7).ValueItems(1).Value.vt=   8
         Columns(7).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
         Columns(7).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
         Columns(7).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
         Columns(7).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
         Columns(7).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
         Columns(7).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
         Columns(7).ValueItems(1).DisplayValue.vt=   9
         Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems.Count=   2
         Columns(7).Caption=   "Contabilizar"
         Columns(7).DataField=   "Marca"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "CodTipoNota"
         Columns(8).DataField=   "CodTipoNota"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Movimientos"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1588"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Button=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1588"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1508"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=3519"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3440"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8192"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2117"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2037"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=8194"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=1402"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1323"
         Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=8194"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=1931"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1852"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=8194"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(7).Width=1931"
         Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1852"
         Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(42)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   3
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(8)   =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&H0&,.bold=-1,.fontsize=825"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HD2D7E6&,.fgcolor=&HA00000&"
         _StyleDefs(24)  =   ":id=22,.bold=-1,.fontsize=1275,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(25)  =   ":id=22,.charset=0"
         _StyleDefs(26)  =   ":id=22,.fontname=Pristina"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HD2D7E6&,.fgcolor=&H0&,.bold=-1"
         _StyleDefs(28)  =   ":id=14,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(75)  =   "Named:id=33:Normal"
         _StyleDefs(76)  =   ":id=33,.parent=0"
         _StyleDefs(77)  =   "Named:id=34:Heading"
         _StyleDefs(78)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   ":id=34,.wraptext=-1"
         _StyleDefs(80)  =   "Named:id=35:Footing"
         _StyleDefs(81)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(82)  =   "Named:id=36:Selected"
         _StyleDefs(83)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(84)  =   "Named:id=37:Caption"
         _StyleDefs(85)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(86)  =   "Named:id=38:HighlightRow"
         _StyleDefs(87)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(88)  =   "Named:id=39:EvenRow"
         _StyleDefs(89)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(90)  =   "Named:id=40:OddRow"
         _StyleDefs(91)  =   ":id=40,.parent=33"
         _StyleDefs(92)  =   "Named:id=41:RecordSelector"
         _StyleDefs(93)  =   ":id=41,.parent=34"
         _StyleDefs(94)  =   "Named:id=42:FilterBar"
         _StyleDefs(95)  =   ":id=42,.parent=33"
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   -74760
         TabIndex        =   49
         Top             =   3900
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Contabilizar"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   5055
         Left            =   -74760
         TabIndex        =   56
         Top             =   840
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   8916
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.CheckBox Check1 
            Caption         =   "Anexar Desc Prod Proveedor"
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   4440
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker10 
            Height          =   345
            Left            =   840
            TabIndex        =   58
            Top             =   3480
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin MSComCtl2.DTPicker DTPicker11 
            Height          =   345
            Left            =   720
            TabIndex        =   59
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptRecepcion 
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Recepcion"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker12 
            Height          =   345
            Left            =   720
            TabIndex        =   61
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   80674817
            CurrentDate     =   40301
         End
         Begin XtremeSuiteControls.RadioButton OptPlanilla 
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Pago a Productores"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdContabilizarPlanilla 
            Height          =   375
            Left            =   360
            TabIndex        =   63
            Top             =   3000
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Contabilizar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   375
            Left            =   360
            TabIndex        =   64
            Top             =   2640
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Consultar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton4 
            Height          =   375
            Left            =   360
            TabIndex        =   65
            Top             =   2880
            Visible         =   0   'False
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Contabilizar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   3960
            Visible         =   0   'False
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Crear Cheque x Proveedor"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.RadioButton OptPlanillaTransportista 
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Pago a Transportista"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptLiquidacion 
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1320
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Liquidacion Leche"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label LblFechaRecepcion 
            Caption         =   "Feha:"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   3600
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Feha:"
            Height          =   255
            Left            =   5280
            TabIndex        =   67
            Top             =   3600
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin TrueOleDBGrid80.TDBGrid TDGridPlanillaLeche 
         Bindings        =   "FrmContabiliza.frx":00BC
         Height          =   4815
         Left            =   -72240
         TabIndex        =   70
         Top             =   960
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8493
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha"
         Columns(0).DataField=   "Fecha_Factura"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Factura No"
         Columns(1).DataField=   "DescripcionMovimiento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nombre Cliente"
         Columns(2).DataField=   "VoucherNo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Sub Total"
         Columns(3).DataField=   "ChequeNo"
         Columns(3).DataWidth=   50
         Columns(3).NumberFormat=   "Standard"
         Columns(3).EditMask=   "##,##.##"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descuento"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4).EditMask=   "##,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "IVA"
         Columns(5).DataField=   "TCambio"
         Columns(5).NumberFormat=   "Standard"
         Columns(5).EditMask=   "##,##0.00"
         Columns(5).EditMaskUpdate=   -1  'True
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Neto Pagar"
         Columns(6).DataField=   "Debito"
         Columns(6).NumberFormat=   "Standard"
         Columns(6).EditMask=   "##,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   80
         Columns(7)._MaxComboItems=   5
         Columns(7).ValueItems(0)._DefaultItem=   0
         Columns(7).ValueItems(0).Value=   "0"
         Columns(7).ValueItems(0).Value.vt=   8
         Columns(7).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
         Columns(7).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
         Columns(7).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
         Columns(7).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
         Columns(7).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
         Columns(7).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
         Columns(7).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
         Columns(7).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
         Columns(7).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
         Columns(7).ValueItems(0).DisplayValue.vt=   9
         Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems(1)._DefaultItem=   0
         Columns(7).ValueItems(1).Value=   "-1"
         Columns(7).ValueItems(1).Value.vt=   8
         Columns(7).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(7).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
         Columns(7).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
         Columns(7).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
         Columns(7).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
         Columns(7).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
         Columns(7).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
         Columns(7).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
         Columns(7).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
         Columns(7).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
         Columns(7).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
         Columns(7).ValueItems(1).DisplayValue.vt=   9
         Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(7).ValueItems.Count=   2
         Columns(7).Caption=   "Contabilizar"
         Columns(7).DataField=   "Conciliada"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "Movimientos"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8194"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1693"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8194"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1931"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1852"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8194"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1931"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1852"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8194"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=1931"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1852"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   3
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(8)   =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(9)   =   ":id=4,.fontname=MS Sans Serif"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&H0&,.bold=-1,.fontsize=825"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HD2D7E6&,.fgcolor=&HA00000&"
         _StyleDefs(24)  =   ":id=22,.bold=-1,.fontsize=1275,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(25)  =   ":id=22,.charset=0"
         _StyleDefs(26)  =   ":id=22,.fontname=Pristina"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HD2D7E6&,.fgcolor=&H0&,.bold=-1"
         _StyleDefs(28)  =   ":id=14,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(71)  =   "Named:id=33:Normal"
         _StyleDefs(72)  =   ":id=33,.parent=0"
         _StyleDefs(73)  =   "Named:id=34:Heading"
         _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   ":id=34,.wraptext=-1"
         _StyleDefs(76)  =   "Named:id=35:Footing"
         _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   "Named:id=36:Selected"
         _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=37:Caption"
         _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(82)  =   "Named:id=38:HighlightRow"
         _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(84)  =   "Named:id=39:EvenRow"
         _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(86)  =   "Named:id=40:OddRow"
         _StyleDefs(87)  =   ":id=40,.parent=33"
         _StyleDefs(88)  =   "Named:id=41:RecordSelector"
         _StyleDefs(89)  =   ":id=41,.parent=34"
         _StyleDefs(90)  =   "Named:id=42:FilterBar"
         _StyleDefs(91)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFDECE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   15495
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Contabilizando del Sistma de Facturacion"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   5505
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   15480
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "FrmContabiliza.frx":00D7
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
   End
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   7560
      Visible         =   0   'False
      Width           =   11295
      _Version        =   786432
      _ExtentX        =   19923
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoConsultaFacturacion 
      Height          =   375
      Left            =   720
      Top             =   8280
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
      Caption         =   "AdoConsultaFacturacion"
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
   Begin MSAdodcLib.Adodc AdoNota 
      Height          =   450
      Left            =   9720
      Top             =   8160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoNota"
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
   Begin MSAdodcLib.Adodc AdoRecepcion 
      Height          =   375
      Left            =   10560
      Top             =   8880
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoRecepcion"
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
   Begin MSAdodcLib.Adodc AdoNominas 
      Height          =   450
      Left            =   840
      Top             =   9120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoNominas"
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
   Begin MSAdodcLib.Adodc AdoBuscaNomina 
      Height          =   450
      Left            =   4680
      Top             =   9120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoBuscaNomina"
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
   Begin MSAdodcLib.Adodc AdoContraCuentaFacturacion 
      Height          =   450
      Left            =   3960
      Top             =   8160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
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
      Caption         =   "AdoContraCuentaFacturacion"
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
End
Attribute VB_Name = "FrmContabilizaFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConexionFacturacion As String

Private Sub CmdConsultaNota_Click()
Dim SqlString As String, FechaInicio As String, FechaFin As String
Dim TipoNota As String

If Me.OptNotaDebito.Value = True Then
 TipoNota = "Debito Clientes"
ElseIf Me.OptNotaCredito.Value = True Then
 TipoNota = "Credito Clientes"
ElseIf Me.OptNotaCreditoProveedor.Value = True Then
    TipoNota = "Credito Proveedores"
ElseIf Me.OptNotaDebitoProveedor.Value = True Then
    TipoNota = "Debito Proveedores"
ElseIf Me.OptPlanillaProductor.Value = True Then
    TipoNota = "PlanillaLeche"
    
End If

Me.DTPicker9.Value = Me.DTPicker8.Value

     Select Case TipoNota
            Case "Debito Proveedores"

         
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
            SqlString = "SELECT  IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto , IndiceNota.Marca, IndiceNota.Tipo_Nota AS CodTipoNota FROM IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB  " & _
                         "WHERE (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (NotaDebito.Tipo Like '%Debito Proveedores%') AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY IndiceNota.Fecha_Nota"
       
               Me.AdoNota.ConnectionString = ConexionFacturacion
                Me.AdoNota.RecordSource = SqlString
                Me.AdoNota.Refresh
                 If Not Me.AdoNota.Recordset.EOF Then
                   Me.CmdContabilizarNotas.Enabled = True
                   Me.DTPicker9.Visible = True
                Else
                   Me.CmdContabilizarNotas.Enabled = False
                   Me.DTPicker9.Visible = False
                End If
                
                Me.TDBGridCuentas.Columns(8).Visible = False
                
       Case "Debito Clientes"

         
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
            SqlString = "SELECT  IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto , IndiceNota.Marca, IndiceNota.Tipo_Nota AS CodTipoNota FROM IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB  " & _
                         "WHERE (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (NotaDebito.Tipo = 'Debito Clientes') AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY IndiceNota.Fecha_Nota"
       
               Me.AdoNota.ConnectionString = ConexionFacturacion
                Me.AdoNota.RecordSource = SqlString
                Me.AdoNota.Refresh
                 If Not Me.AdoNota.Recordset.EOF Then
                   Me.CmdContabilizarNotas.Enabled = True
                   Me.DTPicker9.Visible = True
                Else
                   Me.CmdContabilizarNotas.Enabled = False
                   Me.DTPicker9.Visible = False
                End If
                
                Me.TDBGridCuentas.Columns(8).Visible = False
       
       Case "Credito Clientes"

          
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
            SqlString = "SELECT  IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto , IndiceNota.Marca, IndiceNota.Tipo_Nota AS CodTipoNota FROM IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB  " & _
                         "WHERE (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (NotaDebito.Tipo = 'Credito Clientes') AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY IndiceNota.Fecha_Nota"
       
        Me.AdoNota.ConnectionString = ConexionFacturacion
        Me.AdoNota.RecordSource = SqlString
        Me.AdoNota.Refresh
         If Not Me.AdoNota.Recordset.EOF Then
           Me.CmdContabilizarNotas.Enabled = True
           Me.DTPicker9.Visible = True
        Else
           Me.CmdContabilizarNotas.Enabled = False
           Me.DTPicker9.Visible = False
        End If
        
        Me.TDBGridCuentas.Columns(8).Visible = False
        
        
       Case "Credito Proveedores"

          
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
            SqlString = "SELECT  IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto , IndiceNota.Marca, IndiceNota.Tipo_Nota AS CodTipoNota FROM IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB  " & _
                         "WHERE (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (NotaDebito.Tipo Like '%Credito Proveedores%') AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY IndiceNota.Fecha_Nota"
       
        Me.AdoNota.ConnectionString = ConexionFacturacion
        Me.AdoNota.RecordSource = SqlString
        Me.AdoNota.Refresh
         If Not Me.AdoNota.Recordset.EOF Then
           Me.CmdContabilizarNotas.Enabled = True
           Me.DTPicker9.Visible = True
        Else
           Me.CmdContabilizarNotas.Enabled = False
           Me.DTPicker9.Visible = False
        End If
        
        Me.TDBGridCuentas.Columns(8).Visible = False
        
        
       
       Case "PlanillaLeche"
           
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
           SqlString = "SELECT  NumPlanilla , FechaInicial , FechaFinal, Año, mes, Periodo, Marca From Nomina  " & _
                       "WHERE (FechaInicial <= CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (FechaFinal >= CONVERT(DATETIME, '" & FechaInicio & "', 102))"
       
        Me.AdoNota.ConnectionString = ConexionFacturacion
        Me.AdoNota.RecordSource = SqlString
        Me.AdoNota.Refresh
         If Not Me.AdoNota.Recordset.EOF Then
           Me.CmdContabilizarNotas.Enabled = True
           Me.DTPicker9.Visible = True
        Else
           Me.CmdContabilizarNotas.Enabled = False
           Me.DTPicker9.Visible = False
        End If
        Me.TDBGridCuentas.Columns(0).DataField = "NumPlanilla"
        Me.TDBGridCuentas.Columns(0).Caption = "NumPlanilla"
        Me.TDBGridCuentas.Columns(1).DataField = "FechaInicial"
        Me.TDBGridCuentas.Columns(1).Caption = "Inicio"
        Me.TDBGridCuentas.Columns(2).DataField = "FechaFinal"
        Me.TDBGridCuentas.Columns(2).Caption = "Fin"
        Me.TDBGridCuentas.Columns(3).DataField = "Año"
        Me.TDBGridCuentas.Columns(3).Caption = "Año"
        Me.TDBGridCuentas.Columns(4).DataField = "mes"
        Me.TDBGridCuentas.Columns(4).Caption = "Mes"
        Me.TDBGridCuentas.Columns(5).DataField = "Periodo"
        Me.TDBGridCuentas.Columns(5).Caption = "Periodo"
        Me.TDBGridCuentas.Columns(6).Visible = False
        Me.TDBGridCuentas.Columns(8).Visible = False
    End Select


 

End Sub

Private Sub CmdConsultar_Click()
Dim SqlString As String, FechaInicio As String, FechaFin As String
Dim TipoFactura As String

If Me.OptFacturacion.Value = True Then
 TipoFactura = "Factura"
ElseIf Me.OptRecibos.Value = True Then
 TipoFactura = Me.OptRecibos.Caption
ElseIf Me.OptSalidaBodega.Value = True Then
 TipoFactura = Me.OptSalidaBodega.Caption
ElseIf Me.OptDevolucion.Value = True Then
 TipoFactura = "Devolucion de Venta"
End If

Me.DTPicker5.Value = Me.DTPicker2.Value

     Select Case TipoFactura
          Case "Salida Bodega"
        
            Me.TDBGridFacturacion.Columns(0).DataField = "Fecha_Factura"
            Me.TDBGridFacturacion.Columns(1).DataField = "Numero_Factura"
            Me.TDBGridFacturacion.Columns(2).DataField = "Nombre_Cliente"
            Me.TDBGridFacturacion.Columns(3).DataField = "SubTotal"
            Me.TDBGridFacturacion.Columns(4).DataField = "Descuentos"
            Me.TDBGridFacturacion.Columns(5).DataField = "IVA"
            Me.TDBGridFacturacion.Columns(6).DataField = "NetoPagar"
            Me.TDBGridFacturacion.Columns(7).DataField = "Marca"
            Me.TDBGridFacturacion.Splits(0).Caption = "Listado de Facturas"
            
            
            FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
            SqlString = "SELECT Fecha_Factura, Numero_Factura, Nombre_Cliente, SubTotal, Descuentos, IVA, NetoPagar, Marca From Facturas  " & _
                       "WHERE (Contabilizado = 0)AND (Tipo_Factura = '" & TipoFactura & "') AND (Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Fecha_Factura"
             
            Me.AdoFacturacion.RecordSource = SqlString
            Me.AdoFacturacion.Refresh
            If Not Me.AdoFacturacion.Recordset.EOF Then
               Me.CmdContabilizar.Enabled = True
               Me.LblFecha.Visible = True
               Me.DTPicker5.Visible = True
            Else
               Me.CmdContabilizar.Enabled = False
               Me.LblFecha.Visible = False
               Me.DTPicker5.Visible = False
            End If
        Case "Factura"
        
            Me.TDBGridFacturacion.Columns(0).DataField = "Fecha_Factura"
            Me.TDBGridFacturacion.Columns(1).DataField = "Numero_Factura"
            Me.TDBGridFacturacion.Columns(2).DataField = "Nombre_Cliente"
            Me.TDBGridFacturacion.Columns(3).DataField = "SubTotal"
            Me.TDBGridFacturacion.Columns(4).DataField = "Descuentos"
            Me.TDBGridFacturacion.Columns(5).DataField = "IVA"
            Me.TDBGridFacturacion.Columns(6).DataField = "NetoPagar"
            Me.TDBGridFacturacion.Columns(7).DataField = "Marca"
            Me.TDBGridFacturacion.Splits(0).Caption = "Listado de Facturas"
            
            
            FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
            SqlString = "SELECT Fecha_Factura, Numero_Factura, Nombre_Cliente, SubTotal, Descuentos, IVA, NetoPagar, Marca From Facturas  " & _
                       "WHERE (Contabilizado = 0)AND (Tipo_Factura = '" & TipoFactura & "') AND (Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Fecha_Factura"
             
            Me.AdoFacturacion.RecordSource = SqlString
            Me.AdoFacturacion.Refresh
            If Not Me.AdoFacturacion.Recordset.EOF Then
               Me.CmdContabilizar.Enabled = True
               Me.LblFecha.Visible = True
               Me.DTPicker5.Visible = True
            Else
               Me.CmdContabilizar.Enabled = False
               Me.LblFecha.Visible = False
               Me.DTPicker5.Visible = False
            End If
       
        Case "Devolucion de Venta"
        
            Me.TDBGridFacturacion.Columns(0).DataField = "Fecha_Factura"
            Me.TDBGridFacturacion.Columns(1).DataField = "Numero_Factura"
            Me.TDBGridFacturacion.Columns(2).DataField = "Nombre_Cliente"
            Me.TDBGridFacturacion.Columns(3).DataField = "SubTotal"
            Me.TDBGridFacturacion.Columns(4).DataField = "Descuentos"
            Me.TDBGridFacturacion.Columns(5).DataField = "IVA"
            Me.TDBGridFacturacion.Columns(6).DataField = "NetoPagar"
            Me.TDBGridFacturacion.Columns(7).DataField = "Marca"
            Me.TDBGridFacturacion.Splits(0).Caption = "Listado de Facturas"
            
            
            FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
            SqlString = "SELECT Fecha_Factura, Numero_Factura, Nombre_Cliente, SubTotal, Descuentos, IVA, NetoPagar, Marca From Facturas  " & _
                       "WHERE (Contabilizado = 0)AND (Tipo_Factura = '" & TipoFactura & "') AND (Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Fecha_Factura"
             
            Me.AdoFacturacion.RecordSource = SqlString
            Me.AdoFacturacion.Refresh
            If Not Me.AdoFacturacion.Recordset.EOF Then
               Me.CmdContabilizar.Enabled = True
               Me.LblFecha.Visible = True
               Me.DTPicker5.Visible = True
            Else
               Me.CmdContabilizar.Enabled = False
               Me.LblFecha.Visible = False
               Me.DTPicker5.Visible = False
            End If
       
       Case "Recibos de Caja"
            Me.TDBGridFacturacion.Columns(0).DataField = "Fecha_Recibo"
            Me.TDBGridFacturacion.Columns(1).DataField = "CodReciboPago"
            Me.TDBGridFacturacion.Columns(2).DataField = "NombreCliente"
            Me.TDBGridFacturacion.Columns(3).DataField = "Nombre_Cajero"
            Me.TDBGridFacturacion.Columns(4).DataField = "Sub_Total"
            Me.TDBGridFacturacion.Columns(5).DataField = "Descuento"
            Me.TDBGridFacturacion.Columns(6).DataField = "Total"
            Me.TDBGridFacturacion.Columns(7).DataField = "Marca"
            Me.TDBGridFacturacion.Splits(0).Caption = "Listado de Recibos"
            
            
            FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
            SqlString = "SELECT  Recibo.Fecha_Recibo, Recibo.CodReciboPago, Recibo.NombreCliente, Cajeros.Nombre_Cajero, Recibo.Sub_Total, Recibo.Descuento, Recibo.Total, Recibo.Marca, Recibo.Contabilizado FROM  Recibo INNER JOIN Cajeros ON Recibo.Cod_Cajero = Cajeros.Cod_Cajero  " & _
                        "WHERE (Recibo.Fecha_Recibo BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Recibo.Contabilizado = 0) ORDER BY Recibo.Fecha_Recibo"
      
      
            Me.AdoFacturacion.RecordSource = SqlString
            Me.AdoFacturacion.Refresh
            If Not Me.AdoFacturacion.Recordset.EOF Then
               Me.CmdContabilizar.Enabled = True
               Me.LblFecha.Visible = True
               Me.DTPicker5.Visible = True
            Else
               Me.CmdContabilizar.Enabled = False
               Me.LblFecha.Visible = False
               Me.DTPicker5.Visible = False
            End If
            
            Me.TDBGridFacturacion.Columns(0).Caption = "Fecha Recibo"
            Me.TDBGridFacturacion.Columns(1).Caption = "No Recibo"
            Me.TDBGridFacturacion.Columns(2).Caption = "Nombre Cliente"
            Me.TDBGridFacturacion.Columns(3).Caption = "Cajero"
            Me.TDBGridFacturacion.Columns(4).Caption = "Sub Total"
            Me.TDBGridFacturacion.Columns(5).Caption = "Descuento"
            Me.TDBGridFacturacion.Columns(6).Caption = "Total"
            Me.TDBGridFacturacion.Columns(7).Caption = "Marca"
      
      
      End Select
 

  

 
           
End Sub

Private Sub CmdConsultarCompra_Click()
Dim SqlString As String, FechaInicio As String, FechaFin As String
Dim TipoFactura As String

If Me.OptCompras.Value = True Then
 TipoFactura = "Mercancia Recibida"
ElseIf Me.OptDevolucionCompra.Value = True Then
 TipoFactura = "Devolucion de Compra"
ElseIf Me.OptTransferenciaRecibida.Value = True Then
 TipoFactura = "Transferencia Recibida"
ElseIf Me.RadioButton4.Value = True Then
 TipoFactura = "Pago Proveedor"
ElseIf Me.OptCuenta.Value = True Then
 TipoFactura = "Cuenta"
 
End If

Me.DTPicker6.Value = Me.DTPicker4.Value

     Select Case TipoFactura
             Case "Cuenta"
        
            Me.TDBGridCompras.Columns(0).DataField = "Fecha_Compra"
            Me.TDBGridCompras.Columns(1).DataField = "Numero_Compra"
            Me.TDBGridCompras.Columns(2).DataField = "Nombre_Proveedor"
            Me.TDBGridCompras.Columns(3).DataField = "SubTotal"
            Me.TDBGridCompras.Columns(4).DataField = "Descuento"
            Me.TDBGridCompras.Columns(5).DataField = "IVA"
            Me.TDBGridCompras.Columns(6).DataField = "NetoPagar"
            Me.TDBGridCompras.Columns(7).DataField = "Marca"
            Me.TDBGridCompras.Splits(0).Caption = "Listado de Compras"
            
            
            FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
'            SqlString = "SELECT Fecha_Factura, Numero_Factura, Nombre_Cliente, SubTotal, Descuentos, IVA, NetoPagar, Marca From Facturas  " & _
                       "WHERE (Contabilizado = 0)AND (Tipo_Factura = '" & TipoFactura & "') AND (Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Fecha_Factura"
             SqlString = "SELECT  Fecha_Compra, Numero_Compra, Nombre_Proveedor, SubTotal, Descuento,IVA, NetoPagar,Marca From Compras  " & _
                         "WHERE   (Contabilizado = 0) AND (Tipo_Compra Like '%" & TipoFactura & "%') AND (Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102))"

        Case "Pago Proveedor"
        
        Me.TDBGridCompras.Columns(0).DataField = "Fecha_Recibo"
        Me.TDBGridCompras.Columns(1).DataField = "CodReciboPago"
        Me.TDBGridCompras.Columns(2).DataField = "Cod_Proveedor"
        Me.TDBGridCompras.Columns(3).DataField = "Nombre_Proveedor"
        Me.TDBGridCompras.Columns(4).DataField = "Sub_Total"
        Me.TDBGridCompras.Columns(5).DataField = "Descuento"
        Me.TDBGridCompras.Columns(6).DataField = "Total"
        Me.TDBGridCompras.Columns(7).DataField = "Marca"
        
        
        
        FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
        FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
        SqlString = "SELECT ReciboPago.Fecha_Recibo, ReciboPago.CodReciboPago, ReciboPago.Cod_Proveedor, Proveedor.Nombre_Proveedor, ReciboPago.Sub_Total, ReciboPago.Descuento , ReciboPago.Total, ReciboPago.Marca FROM ReciboPago INNER JOIN Proveedor ON ReciboPago.Cod_Proveedor = Proveedor.Cod_Proveedor  " & _
                    "WHERE (ReciboPago.Fecha_Recibo BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (ReciboPago.Contabilizado = 0)"
                
        Me.TDBGridCompras.Columns(0).Caption = "Fecha_Recibo"
        Me.TDBGridCompras.Columns(1).Caption = "CodReciboPago"
        Me.TDBGridCompras.Columns(2).Caption = "Cod_Proveedor"
        Me.TDBGridCompras.Columns(3).Caption = "Nombre_Proveedor"
        Me.TDBGridCompras.Columns(4).Caption = "Sub_Total"
        Me.TDBGridCompras.Columns(5).Caption = "Descuento"
        Me.TDBGridCompras.Columns(6).Caption = "Total"
        Me.TDBGridCompras.Columns(7).Caption = "Marca"
        
        Case "Recepcion"
        
            Me.TDBGridCompras.Columns(0).DataField = "Fecha_Compra"
            Me.TDBGridCompras.Columns(1).DataField = "Numero_Compra"
            Me.TDBGridCompras.Columns(2).DataField = "Nombre_Proveedor"
            Me.TDBGridCompras.Columns(3).DataField = "SubTotal"
            Me.TDBGridCompras.Columns(4).DataField = "Descuento"
            Me.TDBGridCompras.Columns(5).DataField = "IVA"
            Me.TDBGridCompras.Columns(6).DataField = "NetoPagar"
            Me.TDBGridCompras.Columns(7).DataField = "Marca"
            Me.TDBGridCompras.Splits(0).Caption = "Listado de Compras"
            
            
            FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
             SqlString = "SELECT Fecha as Fecha_Compra, NumeroRecepcion as Numero_Compra,Cod_Proveedor as Nombre_Proveedor, SubTotal, Marca From Recepcion " & _
                         "WHERE (Fecha BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY NumeroRecepcion"
       
       Case "Transferencia Recibida"
            Me.TDBGridCompras.Columns(0).DataField = "Fecha_Compra"
            Me.TDBGridCompras.Columns(1).DataField = "Numero_Compra"
            Me.TDBGridCompras.Columns(2).DataField = "Nombre_Proveedor"
            Me.TDBGridCompras.Columns(3).DataField = "SubTotal"
            Me.TDBGridCompras.Columns(4).DataField = "Descuento"
            Me.TDBGridCompras.Columns(5).DataField = "IVA"
            Me.TDBGridCompras.Columns(6).DataField = "NetoPagar"
            Me.TDBGridCompras.Columns(7).DataField = "Marca"
            Me.TDBGridCompras.Splits(0).Caption = "Listado de Compras"
            
            
            FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
'            SqlString = "SELECT Fecha_Factura, Numero_Factura, Nombre_Cliente, SubTotal, Descuentos, IVA, NetoPagar, Marca From Facturas  " & _
                       "WHERE (Contabilizado = 0)AND (Tipo_Factura = '" & TipoFactura & "') AND (Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Fecha_Factura"
             SqlString = "SELECT  Fecha_Compra, Numero_Compra, Nombre_Proveedor, SubTotal, Descuento,IVA, NetoPagar,Marca From Compras  " & _
                         "WHERE   (Contabilizado = 0) AND (Tipo_Compra = '" & TipoFactura & "') AND (Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102))"
        Case "Mercancia Recibida"
        
            Me.TDBGridCompras.Columns(0).DataField = "Fecha_Compra"
            Me.TDBGridCompras.Columns(1).DataField = "Numero_Compra"
            Me.TDBGridCompras.Columns(2).DataField = "Nombre_Proveedor"
            Me.TDBGridCompras.Columns(3).DataField = "SubTotal"
            Me.TDBGridCompras.Columns(4).DataField = "Descuento"
            Me.TDBGridCompras.Columns(5).DataField = "IVA"
            Me.TDBGridCompras.Columns(6).DataField = "NetoPagar"
            Me.TDBGridCompras.Columns(7).DataField = "Marca"
            Me.TDBGridCompras.Splits(0).Caption = "Listado de Compras"
            
            
            FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
'            SqlString = "SELECT Fecha_Factura, Numero_Factura, Nombre_Cliente, SubTotal, Descuentos, IVA, NetoPagar, Marca From Facturas  " & _
                       "WHERE (Contabilizado = 0)AND (Tipo_Factura = '" & TipoFactura & "') AND (Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Fecha_Factura"
             SqlString = "SELECT  Fecha_Compra, Numero_Compra, Nombre_Proveedor, SubTotal, Descuento,IVA, NetoPagar,Marca From Compras  " & _
                         "WHERE   (Contabilizado = 0) AND (Tipo_Compra = '" & TipoFactura & "') AND (Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102))"
        Case "Devolucion de Compra"
        
            Me.TDBGridCompras.Columns(0).DataField = "Fecha_Compra"
            Me.TDBGridCompras.Columns(1).DataField = "Numero_Compra"
            Me.TDBGridCompras.Columns(2).DataField = "Nombre_Proveedor"
            Me.TDBGridCompras.Columns(3).DataField = "SubTotal"
            Me.TDBGridCompras.Columns(4).DataField = "Descuento"
            Me.TDBGridCompras.Columns(5).DataField = "IVA"
            Me.TDBGridCompras.Columns(6).DataField = "NetoPagar"
            Me.TDBGridCompras.Columns(7).DataField = "Marca"
            Me.TDBGridCompras.Splits(0).Caption = "Listado de Compras"
            
            
            FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
'            SqlString = "SELECT Fecha_Factura, Numero_Factura, Nombre_Cliente, SubTotal, Descuentos, IVA, NetoPagar, Marca From Facturas  " & _
                       "WHERE (Contabilizado = 0)AND (Tipo_Factura = '" & TipoFactura & "') AND (Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY Fecha_Factura"
             SqlString = "SELECT  Fecha_Compra, Numero_Compra, Nombre_Proveedor, SubTotal, Descuento,IVA, NetoPagar,Marca From Compras  " & _
                         "WHERE   (Contabilizado = 0) AND (Tipo_Compra = '" & TipoFactura & "') AND (Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102))"
      End Select
 

  
 Me.AdoCompras.RecordSource = SqlString
 Me.AdoCompras.Refresh
 If Not Me.AdoCompras.Recordset.EOF Then
    Me.CmdContabilizarCompras.Enabled = True
    Me.CmdRecepcion.Enabled = True
    Me.LblFechaCompra.Visible = True
    Me.DTPicker6.Visible = True
 Else
    Me.CmdContabilizar.Enabled = False
    Me.LblFechaCompra.Visible = False
    Me.DTPicker6.Visible = False
    Me.CmdRecepcion.Enabled = False
 End If
End Sub

Private Sub CmdContabilizar_Click()

  Dim Periodo As Double, NumeroPeriodo As Double, FechaIni As String, FechaFin As String, EstadoPeriodo As String, NumeroTransaccion As Double
  Dim mes As Double, Año As Double, Moneda, Resultado As Boolean
  Dim SqlString As String, FechaInicio As String
  Dim TipoFactura As String, NumeroFactura As String, CodigoProducto As String, CodigoCuentaProducto As String, CodigoCliente As String
  Dim NombreCuenta As String, CodigoCuentaCliente As String, SubTotal As Double, Descuento As Double, Iva As Double, NetoPagar As Double
  Dim FechaFactura As Date, FechaVence As Date, DescripcionMovimiento As String, TasaIva As Double
  Dim MonedaFactura As String, Reg As Double, CodigoCuentaIngresos As String, CodCuentaEfectivo As String
  Dim CostoProducto As Double, CodigoCuentaCostos As String, CodigoCuentaInventario As String, CodigoCuentaOtros As String
  Dim CodigoCuentaMetodo As String, Pagado As Double, TotalRetencion As Double, Fuente As String, MonedaMovimiento As String
  Dim cn As New ADODB.Connection, TasaMovimiento As Double, TipoProducto As String, DescripcionProducto As String, UnidadMedida As String
  Dim rs As New ADODB.Recordset, DescripcionRecibo As String, SqlStringAnuladas As String
  Dim cmd As New ADODB.Command, TasaCambioFacturacion As Double
  
  CmdContabilizar.Enabled = False
  
  Reg = 1
  
  
            If Me.OptFacturacion.Value = True Then
             TipoFactura = "Factura"
            ElseIf Me.OptDevolucion.Value = True Then
             TipoFactura = "Devolucion de Venta"
            ElseIf Me.OptRecibos.Value = True Then
              TipoFactura = Me.OptRecibos.Caption
            ElseIf Me.OptSalidaBodega.Value = True Then
               TipoFactura = Me.OptSalidaBodega.Caption
            End If

                Select Case TipoFactura
                   Case "Salida Bodega"
                       FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
                        SqlString = "SELECT Facturas.Fecha_Factura, Facturas.Fecha_Vencimiento, Facturas.MonedaFactura, Facturas.Numero_Factura, Facturas.Cod_Cliente, Facturas.Nombre_Cliente, Facturas.Apellido_Cliente, Facturas.SubTotal, Facturas.IVA, Facturas.Pagado, Facturas.NetoPagar, Facturas.Descuentos, Facturas.Marca, Clientes.Cod_Cuenta_Cliente , Facturas.Contabilizado, Facturas.Tipo_Factura,Facturas.Observaciones FROM  Facturas INNER JOIN Clientes ON Facturas.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                    "WHERE (Facturas.Nombre_Cliente <> N'******CANCELADO') AND (Facturas.Contabilizado = 0) AND (Facturas.Tipo_Factura = '" & TipoFactura & "') AND (Facturas.Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Facturas.Marca = 1) ORDER BY Facturas.Fecha_Factura"
                        SqlStringAnuladas = "SELECT Facturas.Fecha_Factura, Facturas.Fecha_Vencimiento, Facturas.MonedaFactura, Facturas.Numero_Factura, Facturas.Cod_Cliente, Facturas.Nombre_Cliente, Facturas.Apellido_Cliente, Facturas.SubTotal, Facturas.IVA, Facturas.Pagado, Facturas.NetoPagar, Facturas.Descuentos, Facturas.Marca, Clientes.Cod_Cuenta_Cliente , Facturas.Contabilizado, Facturas.Tipo_Factura,Facturas.Observaciones FROM  Facturas INNER JOIN Clientes ON Facturas.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                    "WHERE (Facturas.Nombre_Cliente = '******CANCELADO') AND (Facturas.Contabilizado = 0) AND (Facturas.Tipo_Factura = '" & TipoFactura & "') AND (Facturas.Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Facturas.Marca = 1) ORDER BY Facturas.Fecha_Factura"
                       DescripcionMovimiento = "Movimientos Contables de Salida de Bodega"
                       Fuente = "SalidaBodega"
                   Case "Recibos de Caja"
                        FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
                        FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
                        SqlString = "SELECT  Recibo.Fecha_Recibo, Recibo.CodReciboPago, Recibo.NombreCliente, Cajeros.Nombre_Cajero, Recibo.Sub_Total, Recibo.Descuento, Recibo.Total, Recibo.Marca, Recibo.Contabilizado , Clientes.Cod_Cuenta_Cliente,Recibo.MonedaRecibo FROM Recibo INNER JOIN Cajeros ON Recibo.Cod_Cajero = Cajeros.Cod_Cajero INNER JOIN Clientes ON Recibo.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                    "WHERE (Recibo.Fecha_Recibo BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Recibo.Contabilizado = 0) AND (Recibo.Marca = 1) ORDER BY Recibo.Fecha_Recibo"
                        SqlStringAnuladas = "SELECT  Recibo.Fecha_Recibo, Recibo.CodReciboPago, Recibo.NombreCliente, Cajeros.Nombre_Cajero, Recibo.Sub_Total, Recibo.Descuento, Recibo.Total, Recibo.Marca, Recibo.Contabilizado , Clientes.Cod_Cuenta_Cliente,Recibo.MonedaRecibo FROM Recibo INNER JOIN Cajeros ON Recibo.Cod_Cajero = Cajeros.Cod_Cajero INNER JOIN Clientes ON Recibo.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                    "WHERE (Recibo.Fecha_Recibo BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Recibo.Contabilizado = 0) AND (Recibo.Marca = 1) ORDER BY Recibo.Fecha_Recibo"
                        DescripcionMovimiento = "Movimientos Contables Recibos de Caja"
                        Fuente = "Recibo"
                   Case "Factura"
                       FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
                        SqlString = "SELECT Facturas.Fecha_Factura, Facturas.Fecha_Vencimiento, Facturas.MonedaFactura, Facturas.Numero_Factura, Facturas.Cod_Cliente, Facturas.Nombre_Cliente, Facturas.Apellido_Cliente, Facturas.SubTotal, Facturas.IVA, Facturas.Pagado, Facturas.NetoPagar, Facturas.Descuentos, Facturas.Marca, Clientes.Cod_Cuenta_Cliente , Facturas.Contabilizado, Facturas.Tipo_Factura, Facturas.Observaciones FROM  Facturas INNER JOIN Clientes ON Facturas.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                    "WHERE (Facturas.Nombre_Cliente <> N'******CANCELADO') AND (Facturas.Contabilizado = 0) AND (Facturas.Tipo_Factura = '" & TipoFactura & "') AND (Facturas.Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Facturas.Marca = 1) ORDER BY Facturas.Fecha_Factura"
                        SqlStringAnuladas = "SELECT Facturas.Fecha_Factura, Facturas.Fecha_Vencimiento, Facturas.MonedaFactura, Facturas.Numero_Factura, Facturas.Cod_Cliente, Facturas.Nombre_Cliente, Facturas.Apellido_Cliente, Facturas.SubTotal, Facturas.IVA, Facturas.Pagado, Facturas.NetoPagar, Facturas.Descuentos, Facturas.Marca, Clientes.Cod_Cuenta_Cliente , Facturas.Contabilizado, Facturas.Tipo_Factura, Facturas.Observaciones FROM  Facturas INNER JOIN Clientes ON Facturas.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                            "WHERE (Facturas.Nombre_Cliente = '******CANCELADO') AND (Facturas.Contabilizado = 0) AND (Facturas.Tipo_Factura = '" & TipoFactura & "') AND (Facturas.Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Facturas.Marca = 1) ORDER BY Facturas.Fecha_Factura"
                   Case "Devolucion de Venta"
                       FechaInicio = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker2.Value, "yyyy-mm-dd")
                       SqlString = "SELECT Facturas.Fecha_Factura, Facturas.Fecha_Vencimiento, Facturas.MonedaFactura, Facturas.Numero_Factura, Facturas.Cod_Cliente, Facturas.Nombre_Cliente, Facturas.Apellido_Cliente, Facturas.SubTotal, Facturas.IVA, Facturas.Pagado, Facturas.NetoPagar, Facturas.Descuentos, Facturas.Marca, Clientes.Cod_Cuenta_Cliente , Facturas.Contabilizado, Facturas.Tipo_Factura FROM  Facturas INNER JOIN Clientes ON Facturas.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                    "WHERE (Facturas.Nombre_Cliente <> N'******CANCELADO') AND (Facturas.Contabilizado = 0) AND (Facturas.Tipo_Factura = '" & TipoFactura & "') AND (Facturas.Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Facturas.Marca = 1) ORDER BY Facturas.Fecha_Factura"
                       SqlStringAnuladas = "SELECT Facturas.Fecha_Factura, Facturas.Fecha_Vencimiento, Facturas.MonedaFactura, Facturas.Numero_Factura, Facturas.Cod_Cliente, Facturas.Nombre_Cliente, Facturas.Apellido_Cliente, Facturas.SubTotal, Facturas.IVA, Facturas.Pagado, Facturas.NetoPagar, Facturas.Descuentos, Facturas.Marca, Clientes.Cod_Cuenta_Cliente , Facturas.Contabilizado, Facturas.Tipo_Factura FROM  Facturas INNER JOIN Clientes ON Facturas.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                       "WHERE (Facturas.Nombre_Cliente = N'******CANCELADO') AND (Facturas.Contabilizado = 0) AND (Facturas.Tipo_Factura = '" & TipoFactura & "') AND (Facturas.Fecha_Factura BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Facturas.Marca = 1) ORDER BY Facturas.Fecha_Factura"
                 
                 End Select

                 '//////////////////////////////////////////////////////////////////////////////////////////////
                 '//////////SI LA CUENTA EXISTE AGREGO LOS ENCABEZADOS///////////////////////////////////////
                 '/////////////////////////////////////////////////////////////////////////////////////////////
                 

                         mes = Month(Me.DTPicker5.Value)
                         Año = Year(Me.DTPicker5.Value)
                         FechaIni = CDate("1/" & Month(Me.DTPicker5.Value) & "/" & Year(Me.DTPicker5.Value))
                         FechaFin = DateSerial(Año, mes + 1, 1 - 1)

                 
                         Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
                         Me.AdoConsulta.Refresh
                         If Not Me.AdoConsulta.Recordset.EOF Then
                           Periodo = Me.AdoConsulta.Recordset("Periodo")
                            NumeroPeriodo = Me.AdoConsulta.Recordset("NPeriodo")
                            EstadoPeriodo = Me.AdoConsulta.Recordset("EstadoPeriodo")
                      


                              Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                              Me.AdoConsulta.Recordset.Update
                              NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                              

                               
 
                             Me.AdoProcesos.RecordSource = SqlString
                             Me.AdoProcesos.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                       If Not Me.AdoProcesos.Recordset.EOF Then
                         Me.AdoProcesos.Recordset.MoveFirst
                         Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                       End If
                       
                       
                       Me.AdoProcesos.Refresh
                       Do While Not Me.AdoProcesos.Recordset.EOF
                             
                         Select Case TipoFactura
                              Case "Recibos de Caja"
                                FechaFactura = Me.AdoProcesos.Recordset("Fecha_Recibo")
                                FechaVence = Me.AdoProcesos.Recordset("Fecha_Recibo")
                                NumeroFactura = Me.AdoProcesos.Recordset("CodReciboPago")
                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Cliente")
                                
                                If CodigoCuentaCliente = "" Then
                                  MsgBox "No existe la cuenta del Cliente", vbCritical, "Zeus contable"
                                  Exit Sub
                                End If
                                
                                    SubTotal = 0
                                    Descuento = 0
                                    Iva = 0
                                    NetoPagar = 0
                                    Pagado = 0
                    
'                                Me.AdoConsultaFactura.RecordSource = "SELECT * From DetalleRecibo WHERE (CodReciboPago = '" & NumeroFactura & "') AND (Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) ORDER BY Fecha_Recibo"
                                Me.AdoConsultaFactura.RecordSource = "SELECT * FROM DetalleRecibo INNER JOIN MetodoPago ON DetalleRecibo.NombrePago = MetodoPago.NombrePago WHERE (DetalleRecibo.CodReciboPago = '" & NumeroFactura & "') AND (DetalleRecibo.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102))  ORDER BY DetalleRecibo.Fecha_Recibo"
                                Me.AdoConsultaFactura.Refresh
                                If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                     DescripcionRecibo = Me.AdoConsultaFactura.Recordset("Descripcion")
                                  If Not IsNull(Me.AdoConsultaFactura.Recordset("MontoPagado")) Then
                                    If Me.AdoConsultaFactura.Recordset("MontoPagado") = Me.AdoProcesos.Recordset("Sub_Total") Then
                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("Sub_Total")), "##,##0.00")
                                        If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                          Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                        End If
                                        NetoPagar = Format(Val(Me.AdoProcesos.Recordset("Total")), "##,##0.00")
                                    Else
                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("Sub_Total")), "##,##0.00")
                                        If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                         Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                        End If
                                        NetoPagar = Me.AdoConsultaFactura.Recordset("MontoPagado")
                                        


                                    End If
                               
                                
                                
                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                       MonedaFactura = Me.AdoProcesos.Recordset("MonedaRecibo")
                                       TasaMovimiento = 1
    
                                        SqlString = "SELECT * From TasaCambio WHERE (FechaTasa = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102))"
                                        Me.AdoBuscaFacturacion.RecordSource = SqlString
                                        Me.AdoBuscaFacturacion.Refresh
                                        If Not Me.AdoBuscaFacturacion.Recordset.EOF Then
                                          TasaCambio = Format(Val(Me.AdoBuscaFacturacion.Recordset("MontoTasa")), "##,##0.0000")
                                        Else
                                          TasaCambio = 1
                                        End If
                                   End If
                                 End If
                                 
                                      If MonedaFactura = "Cordobas" Then
                                        TasaCambio = 1
                                        MonedaMovimiento = "Córdobas"
                                      End If
                                 
                                 
                                      If Reg = 1 Then
                                         '////////////////////////////////////////////////////////////////
                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                         '/////////////////////////////////////////////////////////////////
                                         MonedaMovimiento = "Córdobas"
                                         Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker5.Value, "yyyy-mm-dd"), "Movimiento de Recibos", "Recibo", MonedaMovimiento)
                                         Reg = 2
                                      End If
                                      
                                       DescripcionMovimiento = "Registrando Recibo No " & NumeroFactura & "  " & DescripcionRecibo
                                      
                                      '///////////////////////////////////////////////////////////////////////////////////
                                      '///////////////////////GRABO EL MOVIMIENTO DEL CLIENTE ///////////////////////////
                                      '///////////////////////////////////////////////////////////////////////////////////
                                       Debito = 0
                                       Credito = SubTotal
                                       Credito = Format(SubTotal * TasaCambio, "##,##0.00")
                                       Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Format(Me.DTPicker5.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Recibo", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "Recibo")
                                      
                                      '///////////////////////////////////////////////////////////////////////////////////////
                                      '/////////////////////////////GRABA DETALLE DE RECIBO RETENCIONES///////////////////////////////////
                                      '///////////////////////////////////////////////////////////////////////////////////////
                                       TotalRetencion = 0
                                       Me.AdoConsultaFactura.RecordSource = "SELECT  DetalleRecibo.idDetalleRecibo, DetalleRecibo.CodReciboPago, DetalleRecibo.Fecha_Recibo, DetalleRecibo.Numero_Factura, DetalleRecibo.MontoPagado, DetalleRecibo.NombrePago, DetalleRecibo.Descripcion, DetalleRecibo.NumeroTarjeta, DetalleRecibo.FechaVence, DetalleRecibo.MontoFactura, DetalleRecibo.AplicaFactura, DetalleRecibo.SaldoFactura, DetalleRecibo.TasaCambio, MetodoPago.NombrePago AS Expr1, MetodoPago.TipoPago, MetodoPago.Cod_Cuenta , MetodoPago.Moneda  FROM  DetalleRecibo INNER JOIN MetodoPago ON DetalleRecibo.NombrePago = MetodoPago.NombrePago  " & _
                                                                            "WHERE  (DetalleRecibo.CodReciboPago = '" & NumeroFactura & "') AND (DetalleRecibo.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) AND (DetalleRecibo.MontoPagado < 0) ORDER BY DetalleRecibo.Fecha_Recibo"
                                       Me.AdoConsultaFactura.Refresh
                                      Do While Not Me.AdoConsultaFactura.Recordset.EOF
                                        CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                        NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                        Debito = Abs(Me.AdoConsultaFactura.Recordset("MontoPagado"))
                                        Debito = Format(Debito * TasaCambio, "##,##0.00")
                                        Credito = 0
'                                        DescripcionMovimiento = "Movimiento de Registro de Recibos No" & NumeroFactura
                                        Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Format(Me.DTPicker5.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Recibo", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "Recibo")
                                        Me.AdoConsultaFactura.Recordset.MoveNext
                                        TotalRetencion = Abs(Debito) + TotalRetencion
                                      Loop
                                      
                                      '///////////////////////////////////////////////////////////////////////////////////////
                                      '/////////////////////////////GRABA DETALLE DE RECIBO NO RETENCIONES///////////////////////////////////
                                      '///////////////////////////////////////////////////////////////////////////////////////
                                      
                                       Me.AdoConsultaFactura.RecordSource = "SELECT  DetalleRecibo.idDetalleRecibo, DetalleRecibo.CodReciboPago, DetalleRecibo.Fecha_Recibo, DetalleRecibo.Numero_Factura, DetalleRecibo.MontoPagado, DetalleRecibo.NombrePago, DetalleRecibo.Descripcion, DetalleRecibo.NumeroTarjeta, DetalleRecibo.FechaVence, DetalleRecibo.MontoFactura, DetalleRecibo.AplicaFactura, DetalleRecibo.SaldoFactura, DetalleRecibo.TasaCambio, MetodoPago.NombrePago AS Expr1, MetodoPago.TipoPago, MetodoPago.Cod_Cuenta , MetodoPago.Moneda  FROM  DetalleRecibo INNER JOIN MetodoPago ON DetalleRecibo.NombrePago = MetodoPago.NombrePago  " & _
                                                                            "WHERE  (DetalleRecibo.CodReciboPago = '" & NumeroFactura & "') AND (DetalleRecibo.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) AND (DetalleRecibo.MontoPagado > 0) ORDER BY DetalleRecibo.Fecha_Recibo"
                                       Me.AdoConsultaFactura.Refresh
                                      Do While Not Me.AdoConsultaFactura.Recordset.EOF
                                        CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                        NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                        Debito = (Me.AdoConsultaFactura.Recordset("MontoPagado") * TasaCambio) - TotalRetencion
'                                        Debito = Format(Debito * TasaCambio, "##,##0.00")
                                        Credito = 0
'
                                        DescripcionMovimiento = "Registrando Recibo No " & NumeroFactura & "  " & DescripcionRecibo & " Factura No. " & Me.AdoConsultaFactura.Recordset("Numero_Factura")
                                        Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Format(Me.DTPicker5.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Recibo", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "Recibo")
                                        Me.AdoConsultaFactura.Recordset.MoveNext
                                      Loop
                                      
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                rs.Open "UPDATE Recibo SET Contabilizado = 1 ,Activo = 0  WHERE (CodReciboPago = '" & NumeroFactura & "') ", ConexionFacturacion
                                      
                                      
                              
                              Case Else
                              
                              
                              '-----------------------------------------------------------------------------------------------------------
                              '----------------------------------------INICIO DE LAS FACTURAS --------------------------------------------
                              '-----------------------------------------------------------------------------------------------------------
                              
                                FechaFactura = Me.AdoProcesos.Recordset("Fecha_Factura")
                                FechaVence = Me.AdoProcesos.Recordset("Fecha_Vencimiento")
                                NumeroFactura = Me.AdoProcesos.Recordset("Numero_Factura")
                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Cliente")
                                MonedaFactura = Me.AdoProcesos.Recordset("MonedaFactura")
                                
                                If NumeroFactura = "02042" Then
                                  NumeroFactura = "02042"
                                End If
                                
                                    SubTotal = 0
                                    Descuento = 0
                                    Iva = 0
                                    NetoPagar = 0
                                    Pagado = 0
                                    TasaMovimiento = 1
                                    
                                    
                                        SqlString = "SELECT * From Detalle_Facturas WHERE (Numero_Factura = '" & NumeroFactura & "') AND (Tipo_Factura = '" & TipoFactura & "')"
                                        Me.AdoBuscaFacturacion.RecordSource = SqlString
                                        Me.AdoBuscaFacturacion.Refresh
                                        If Not Me.AdoBuscaFacturacion.Recordset.EOF Then
                                          If MonedaFactura = "Dolares" Then
'                                              TasaCambio = Format(Val(Me.AdoBuscaFacturacion.Recordset("TasaCambio")), "##,##0.0000")
                                               
                                               
                                               TasaCambioFacturacion = BuscaTasaCambioFacturacion(FechaFactura, ConexionFacturacion)
                                               
                                               If TasaCambioFacturacion = 0 Then
                                                 TasaCambio = BuscaTasaCambio(FechaFactura)
                                               Else
                                                 TasaCambio = TasaCambioFacturacion
                                               End If
                                               
                                          Else
                                              TasaCambio = 1
                                          End If
                                        Else
                                          TasaCambio = 0
'                                          Moneda = "Córdobas"
                                        End If
                    
                                'SUM(Detalle_Facturas.Descuento) AS Descuento
                                Me.AdoConsultaFactura.RecordSource = "SELECT SUM(Detalle_Facturas.Cantidad) AS Cantidad, SUM(Detalle_Facturas.Precio_Unitario) AS Precio_Unitario, SUM(Detalle_Facturas.Precio_Neto) AS Precio_Neto, SUM(Detalle_Facturas.Importe) AS Importe FROM Detalle_Facturas INNER JOIN Facturas ON Detalle_Facturas.Numero_Factura = Facturas.Numero_Factura AND Detalle_Facturas.Fecha_Factura = Facturas.Fecha_Factura AND Detalle_Facturas.Tipo_Factura = Facturas.Tipo_Factura  " & _
                                                              "WHERE (Facturas.Numero_Factura = '" & NumeroFactura & "') AND (Facturas.Nombre_Cliente <> N'******CANCELADO') AND (Facturas.Tipo_Factura = '" & TipoFactura & "')"
                                Me.AdoConsultaFactura.Refresh
                                If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                  If Not IsNull(Me.AdoConsultaFactura.Recordset("Importe")) Then
                                    If Format(Val(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00") = Format(Val(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00") Then
                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00")
                                        
                                        If IsNumeric(Me.AdoProcesos.Recordset("Descuentos")) Then
                                         Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuentos")), "##,##0.00")
                                        Else
                                         Descuento = 0
                                        End If
                                        
                                        Iva = Format(Val(Me.AdoProcesos.Recordset("IVA")), "##,##0.00")
                                       
                                        NetoPagar = Format(Val(Me.AdoProcesos.Recordset("NetoPagar")), "##,##0.00")
                                        
                                        Pagado = Format(Val(Me.AdoProcesos.Recordset("Pagado")), "##,##0.00")
                                        
                                    Else
                                        SubTotal = Format(Val(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00")
                                        Descuento = 0
                                        Iva = 0
                                        NetoPagar = 0
                                        Pagado = 0
                                    End If
                               
                                
                                
                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                       MonedaFactura = Me.AdoProcesos.Recordset("MonedaFactura")
    

                              
                                      If Reg = 1 Then
                                         '////////////////////////////////////////////////////////////////
                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                         '/////////////////////////////////////////////////////////////////
                                         Fuente = TipoFactura
                                         MonedaMovimiento = "Córdobas"
                                         DescripcionMovimiento = "Contabilizacion de Facturacion"
                                         Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker5.Value, "yyyy-mm-dd"), DescripcionMovimiento, Fuente, MonedaMovimiento)
                                         Reg = 2
                                      End If
                                           
                                        '////////////////////////////////////////////////////////////////////////////////////
                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                        '///////////////////////////////////////////////////////////////////////////////////////
                                           
                                        DescripcionMovimiento = ""
                                        If Me.ChkDescripcion.Value = 1 Then
                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Facturas.Numero_Factura = '" & NumeroFactura & "')"
                                          Me.AdoBuscaFacturacion.Refresh
                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                 DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                          Loop
                                        
                                        Else
                                        
                                          Select Case TipoFactura
                                            Case "Factura"
                                              DescripcionMovimiento = "Contabilizando Factura Numero " & NumeroFactura
                                            Case "Salida Bodega"
                                              DescripcionMovimiento = "Contabilizando Salida Bodega Numero " & NumeroFactura
                                            Case "Devolucion de Venta"
                                              DescripcionMovimiento = "Contabilizando Devolucion Venta Numero " & NumeroFactura
                                          
                                          End Select
                                          
                                        End If
                                        
'                                        DescripcionMovimiento = DescripcionMovimiento & " " & Me.AdoProcesos.Recordset("Observaciones")
                                           
                                           
                                     Credito = 0
                                     If NetoPagar = 0 Then
                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                            Me.AdoConsultaFactura.RecordSource = "SELECT  * FROM Detalle_MetodoFacturas INNER JOIN MetodoPago ON Detalle_MetodoFacturas.NombrePago = MetodoPago.NombrePago  " & _
                                                                                 "WHERE (Detalle_MetodoFacturas.Numero_Factura = '" & NumeroFactura & "')"
                                            Me.AdoConsultaFactura.Refresh
                                            
                                            If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                                CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
'                                                DescripcionMovimiento = "Registro de Facturacion Factura Numero " & NumeroFactura
                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                NetoPagar = SubTotal + Iva - Descuento
                                                Select Case TipoFactura
                                                  Case "Factura"
                                                    Debito = Format(NetoPagar * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                  Case "Devolucion de Venta"
                                                     Debito = 0
                                                     Credito = Format(NetoPagar * TasaCambio, "##,##0.00")
                                                  Case "Salida Bodega"
                                                    Debito = Format(NetoPagar * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                End Select

                                                '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////////////////BUSCO EL PAGO DE CONTADO///////////////////////////////////////////
                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                    If Pagado <> 0 Then
                                                      SqlString = "SELECT  Detalle_MetodoFacturas.Numero_Factura, Detalle_MetodoFacturas.Fecha_Factura, Detalle_MetodoFacturas.NombrePago, Detalle_MetodoFacturas.Monto, Detalle_MetodoFacturas.NumeroTarjeta , MetodoPago.Cod_Cuenta, MetodoPago.Moneda FROM  Detalle_MetodoFacturas INNER JOIN MetodoPago ON Detalle_MetodoFacturas.NombrePago = MetodoPago.NombrePago  " & _
                                                                  "WHERE (Detalle_MetodoFacturas.Numero_Factura = '" & NumeroFactura & "')"
                                                      Me.AdoConsultaFacturacion.RecordSource = SqlString
                                                      Me.AdoConsultaFacturacion.Refresh
                                                      Do While Not Me.AdoConsultaFacturacion.Recordset.EOF
                                                         CodigoCuentaMetodo = Me.AdoConsultaFacturacion.Recordset("Cod_Cuenta")
                                                         
                                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaMetodo)
                                                        Select Case TipoFactura
                                                            Case "Salida Bodega"
                                                             DescripcionMovimiento = "PAGO DE COMPRA" & NumeroFactura
                                                             Debito = Me.AdoConsultaFacturacion.Recordset("Monto") * TasaCambio
                                                             Credito = 0
                                                                 If DescripcionCuenta <> "Nulo" Then
'                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                 End If
                                                            Case "Factura"
                                                             DescripcionMovimiento = "PAGO DE COMPRA" & NumeroFactura
                                                             Debito = Me.AdoConsultaFacturacion.Recordset("Monto") * TasaCambio
                                                             Credito = 0
                                                                 If DescripcionCuenta <> "Nulo" Then
'                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                 End If
                                                        
                                                            Case "Devolucion de Venta"
                                                             DescripcionMovimiento = "PAGO DE DEVOLUCION" & NumeroFactura
                                                             Debito = 0
                                                             Credito = Me.AdoConsultaFacturacion.Recordset("Monto") * TasaCambio
                                                                 If DescripcionCuenta <> "Nulo" Then
'                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                 End If
                                                        End Select
                                                        
                                                        
                                                    
    
                                                         If Debito <> 0 Then
                                                             Resultado = GrabaDetalleFactura(CodigoCuentaMetodo, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", 1, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                         End If
                                                       
                                                       Me.AdoConsultaFacturacion.Recordset.MoveNext
                                                      Loop
                                                    
                                                End If
                                                

                                                
                                            Else
                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                NetoPagar = SubTotal + Iva - Descuento
                                                Select Case TipoFactura
                                                  Case "Salida Bodega"
                                                    Debito = Format(NetoPagar, "##,##0.00")
                                                    Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                    Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, Fuente, NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "SalidaBodega")
                                                  Case "Factura"
                                                    Debito = Format(NetoPagar, "##,##0.00")
                                                    Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                    'Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, Fuente, NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                  Case "Devolucion de Venta"
                                                     Debito = 0
                                                     Credito = Format(NetoPagar, "##,##0.00")
                                                     Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                     Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, Fuente, NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                End Select
                                            End If
                                  ElseIf NetoPagar < 0 Then
                                              '/////////////////AGREGO LA DIFERENICIA A OTROS INGRESOS /
                                               Me.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (TipoCuenta = 'Capital')"
                                               Me.AdoConsulta.Refresh
                                               If Not Me.AdoConsulta.Recordset.EOF Then
                                                CodigoCuentaOtros = Me.AdoConsulta.Recordset("CodCuentas")
                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaOtros)
                                               End If
'                                               DescripcionMovimiento = "Registro de AJUSTE Factura Numero " & NumeroFactura
                                               Resultado = GrabaDetalleFactura(CodigoCuentaOtros, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, 0, Abs(NetoPagar), "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                               
                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                            Me.AdoConsultaFactura.RecordSource = "SELECT  * FROM Detalle_MetodoFacturas INNER JOIN MetodoPago ON Detalle_MetodoFacturas.NombrePago = MetodoPago.NombrePago  " & _
                                                                                 "WHERE (Detalle_MetodoFacturas.Numero_Factura = '" & NumeroFactura & "')"
                                            Me.AdoConsultaFactura.Refresh
                                              CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                              NetoPagar = SubTotal + Iva - Descuento + Abs(NetoPagar)
                                              
                                               Select Case TipoFactura
                                                 Case "Salida Bodega"
                                                    Debito = Format(NetoPagar, "##,##0.00")
                                                    Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                    Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                  Case "Factura"
                                                    Debito = Format(NetoPagar, "##,##0.00")
                                                    Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                    Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                  Case "Devolucion de Venta"
                                                     Debito = 0
                                                     Credito = Format(NetoPagar, "##,##0.00")
                                                     Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                     Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                End Select
                                              DescripcionMovimiento = "Registro de Facturacion Factura Numero " & NumeroFactura
                                              DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                              
                                              
                                     ElseIf NetoPagar > 0 Then
                                               Select Case TipoFactura
                                                  Case "Salida Bodega"
                                                    Debito = Format(NetoPagar, "##,##0.00")
                                                    Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                    If Me.ChkCtaCtoProducto.Value = 0 Then
                                                      Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, Fuente, NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                    End If
                                                  Case "Factura"
                                                    Debito = Format(NetoPagar, "##,##0.00")
                                                    Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                    Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                  Case "Devolucion de Venta"
                                                     Debito = 0
                                                     Credito = Format(NetoPagar, "##,##0.00")
                                                     Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                     Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                End Select
'                                            DescripcionMovimiento = "Registro de Facturacion Factura Numero " & NumeroFactura
                                            DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                            
                                      End If
                                      
                                      
                                  
                              
                                     '//////////////////////////////////////////////////////////////////////////////////////////
                                    '/////////////////////CARGO LOS DETALLE DE LAS FACTURAS////////////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////////
                                     SqlString = "SELECT Detalle_Facturas.Numero_Factura, Detalle_Facturas.Fecha_Factura, Detalle_Facturas.Tipo_Factura, Detalle_Facturas.Cod_Producto,Detalle_Facturas.Descripcion_Producto, Detalle_Facturas.Cantidad, Detalle_Facturas.Precio_Unitario, Detalle_Facturas.Descuento,Detalle_Facturas.Precio_Neto, Detalle_Facturas.Importe, Detalle_Facturas.TasaCambio, Productos.Cod_Cuenta_Inventario,Productos.Cod_Cuenta_Costo, Productos.Cod_Cuenta_Ventas, Productos.Cod_Cuenta_GastoAjuste, Productos.Cod_Cuenta_IngresoAjuste,Productos.Costo_Promedio , Productos.Costo_Promedio_Dolar,Productos.Tipo_Producto,Productos.Unidad_Medida FROM Detalle_Facturas INNER JOIN Productos ON Detalle_Facturas.Cod_Producto = Productos.Cod_Productos  " & _
                                                 "WHERE (Detalle_Facturas.Numero_Factura = '" & NumeroFactura & "') AND (Detalle_Facturas.Tipo_Factura = '" & TipoFactura & "')"
                                     Me.AdoProcesosFacturacion.RecordSource = SqlString
                                     Me.AdoProcesosFacturacion.Refresh

                                        
                                       Do While Not Me.AdoProcesosFacturacion.Recordset.EOF
                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                CodigoCuentaProducto = BuscaCodigoProducto(CodigoProducto)
                                                Cantidad = Me.AdoProcesosFacturacion.Recordset("Cantidad")
                                                TipoProducto = Me.AdoProcesosFacturacion.Recordset("Tipo_Producto")
                                                DescripcionProducto = Me.AdoProcesosFacturacion.Recordset("Descripcion_Producto")
                                                UnidadMedida = Me.AdoProcesosFacturacion.Recordset("Unidad_Medida")
                                                
                                                

                                                
                                                If TipoProducto <> "Descuento" Then
                                                    If MonedaFactura = "Dolares" Then
                                                       CostoProducto = Cantidad * CalcularCostoPromedio(Trim(CodigoProducto), ConexionFacturacion)
'                                                       CostoProducto = CostoProducto / TasaCambio
                                                    Else
                                                       CostoProducto = Cantidad * CalcularCostoPromedio(Trim(CodigoProducto), ConexionFacturacion)
                                                    End If
                                                  
                                                Else
                                                  CostoProducto = 0
                                                End If
                                                
                                                
                                                
                                                

                                            
                                               '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                               '///////////////////////////CALCULO EL IVA DE CADA PRODUCTO//////////////////////////////////////////////
                                               '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                               
                                               

                                                   If Iva <> 0 Then
                                                         TasaIva = BuscaTasaIvaFactura(CodigoProducto)
                                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaIva)
                                                         QUIEN = "IVA"
    
    
                                                         DescripcionMovimiento = "IVA Ventas Factura No " & NumeroFactura
                                                         If DescripcionCuenta <> "Nulo" Then
    
                                                            Select Case TipoFactura
                                                               Case "Factura"
                                                                Debito = 0
'                                                                Credito = Iva
                                                                Credito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                                If Credito <> 0 Then
                                                                    Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                End If
                                                               Case "Devolucion de Venta"
                                                                Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
'                                                                Debito = Iva
                                                                Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                Credito = 0
                                                                If Debito <> 0 Then
                                                                    Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                End If
                                                             End Select
                                                         End If
                                                    End If
                                               
                                                QUIEN = ""
                                                
                                                '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////BUSCO LA CUENTA DE INGRESOS//////////////////////////////////////////////
                                                '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                                
                                                If Not IsNull(Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Ventas")) Then
                                                  CodigoCuentaIngresos = Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Ventas")
                                                End If
                                                
                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaIngresos)
                                                DescripcionMovimiento = "Ventas Prod " & CodigoProducto & " " & DescripcionProducto & " Cantidad: " & Cantidad & " U/M: " & UnidadMedida
'                                                If DescripcionCuenta <> "Nulo" Then
                                                    Select Case TipoFactura
                                                      Case "Factura"
                                                            Debito = 0
                                                            Credito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")), "##,##0.00")
                                                            Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                            If Credito <> 0 Then
                                                              If TipoProducto <> "Descuento" Then
                                                                 Resultado = GrabaDetalleFactura(CodigoCuentaIngresos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                              Else
                                                                 Credito = 0
                                                                 Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")), "##,##0.00")
                                                                 Debito = Abs(Debito * TasaCambio)
                                                                 Resultado = GrabaDetalleFactura(CodigoCuentaIngresos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                              End If
                                                            End If
                                                      Case "Devolucion de Venta"
                                                            Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")), "##,##0.00")
                                                            Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                            Credito = 0
                                                            If Debito <> 0 Then
                                                             Resultado = GrabaDetalleFactura(CodigoCuentaIngresos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                            End If
                                                    End Select

'                                                End If
                                         
                                         
                                                '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////BUSCO LA CUENTA DE COSTOS//////////////////////////////////////////////
                                                '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                                
                                                
                                                If Not IsNull(Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Costo")) Then
                                                  CodigoCuentaCostos = Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Costo")
                                                End If
                                                
                                                Select Case TipoFactura
                                                  Case "Factura"
                                                       Debito = Format(Val(CostoProducto), "##,##0.00")
                                                       Debito = Format(Debito, "##,##0.00")
                                                       Credito = 0
                                                       DescripcionMovimiento = "Costo Producto " & CodigoProducto & " " & DescripcionProducto & " Cantidad: " & Cantidad & " U/M: " & UnidadMedida
                                                       DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
                                                       
'                                                       If DescripcionCuenta <> "Nulo" Then
                                                        Resultado = GrabaDetalleFactura(CodigoCuentaCostos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
'                                                       End If
                                                  Case "Devolucion de Venta"
                                                       Debito = 0
                                                       Credito = Format(Val(CostoProducto), "##,##0.00")
                                                       Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                       DescripcionMovimiento = "Costo Producto " & CodigoProducto & " " & DescripcionProducto & " Cantidad: " & Cantidad & " U/M: " & UnidadMedida
                                                       DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
            
                                                       If DescripcionCuenta <> "Nulo" Then
                                                        Resultado = GrabaDetalleFactura(CodigoCuentaCostos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                       End If
                                                End Select

                                                
                                                
                                                 
                                                '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////////////////BUSCO LA CUENTA DE INVENTARIO///////////////////////////////////////////
                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                 
                                                If Not IsNull(Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")) Then
                                                  CodigoCuentaInventario = Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")
                                                End If
                                               
                                               Select Case TipoFactura
                                                  Case "Salida Bodega"
                                                    Debito = 0
                                                    Credito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")), "##,##0.00")
                                                    Credito = Format(Credito, "##,##0.00")
                                                    
                                                    DescripcionMovimiento = "Costo Producto " & CodigoProducto & " " & DescripcionProducto & " Cantidad: " & Cantidad & " U/M: " & UnidadMedida
                                                    DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
     
'                                                    If DescripcionCuenta <> "Nulo" Then
                                                     Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, Fuente, NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
'                                                    End If
                                                     If Me.ChkCtaCtoProducto.Value = 1 Then
                                                        If Not IsNull(Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Costo")) Then
                                                          CodigoCuentaCostos = Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Costo")
                                                        End If
                                                        Resultado = GrabaDetalleFactura(CodigoCuentaCostos, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Credito, Debito, Fuente, NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                        
                                                     End If
                                                  Case "Factura"
                                                    Debito = 0
                                                    Credito = Format(Val(CostoProducto), "##,##0.00")
                                                    Credito = Format(Credito, "##,##0.00")
                                                    DescripcionMovimiento = "Costo Producto " & CodigoProducto & " " & DescripcionProducto & " Cantidad: " & Cantidad & " U/M: " & UnidadMedida
                                                    DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
     
'                                                    If DescripcionCuenta <> "Nulo" Then
                                                     Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
'                                                    End If
                                                  Case "Devolucion de Venta"
                                                    Debito = Format(Val(CostoProducto), "##,##0.00")
                                                    Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                    Credito = 0
                                                    DescripcionMovimiento = "Costo Producto " & CodigoProducto & " " & DescripcionProducto & " Cantidad: " & Cantidad & " U/M: " & UnidadMedida
                                                    DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
     
                                                    If DescripcionCuenta <> "Nulo" Then
                                                     Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker5.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "VTAS", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                    End If
                                                End Select

                                     
                                        Me.AdoProcesosFacturacion.Recordset.MoveNext
                                     Loop
                              
                                   End If
                                End If
                              
                              
                              
                              
                              End Select
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                rs.Open "UPDATE Facturas SET Contabilizado = 1 ,Activo = 0  WHERE (Numero_Factura = '" & NumeroFactura & "') AND (Facturas.Tipo_Factura = '" & TipoFactura & "')", ConexionFacturacion
                                     

                                
                                
                                

                        Me.osProgress1.Value = Me.osProgress1.Value + 1
                        Me.AdoProcesos.Recordset.MoveNext
                    Loop
                    
                    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    '////////////////////////////////////CAMBIO EL VALOR DE CONTABILIZADO PARA LOS ANULADOS ///////////////////////////////////////////////
                    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    
                       Me.AdoProcesos.RecordSource = SqlStringAnuladas
                       Me.AdoProcesos.Refresh
                       Me.osProgress1.Visible = True
                       Me.osProgress1.Min = 0
                       Me.osProgress1.Value = 0
                       If Not Me.AdoProcesos.Recordset.EOF Then
                         Me.AdoProcesos.Recordset.MoveFirst
                         Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                       End If
                       
                       
                       Me.AdoProcesos.Refresh
                       Do While Not Me.AdoProcesos.Recordset.EOF
                       
                         
                       
                        Select Case TipoFactura
                           Case "Salida Bodega"
                                 NumeroFactura = Me.AdoProcesos.Recordset("Numero_Factura")
                                 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                 '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                  rs.Open "UPDATE Facturas SET Contabilizado = 1 ,Activo = 0  WHERE (Numero_Factura = '" & NumeroFactura & "') AND (Facturas.Tipo_Factura = '" & TipoFactura & "')", ConexionFacturacion
                           Case "Recibos de Caja"
                               NumeroFactura = Me.AdoProcesos.Recordset("CodReciboPago")
                               '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                               '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                               '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                 rs.Open "UPDATE Recibo SET Contabilizado = 1 ,Activo = 0  WHERE (CodReciboPago = '" & NumeroFactura & "') ", ConexionFacturacion
                          Case "Factura"
                                NumeroFactura = Me.AdoProcesos.Recordset("Numero_Factura")
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                 '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                  rs.Open "UPDATE Facturas SET Contabilizado = 1 ,Activo = 0  WHERE (Numero_Factura = '" & NumeroFactura & "') AND (Facturas.Tipo_Factura = '" & TipoFactura & "')", ConexionFacturacion
                          Case "Devolucion de Venta"
                                 NumeroFactura = Me.AdoProcesos.Recordset("Numero_Factura")
                                  '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                 '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                  rs.Open "UPDATE Facturas SET Contabilizado = 1 ,Activo = 0  WHERE (Numero_Factura = '" & NumeroFactura & "') AND (Facturas.Tipo_Factura = '" & TipoFactura & "')", ConexionFacturacion
                        End Select
                       
                       
                            
                           Me.osProgress1.Value = Me.osProgress1.Value + 1
                          Me.AdoProcesos.Recordset.MoveNext
                       Loop

                             
'

               End If
               
               
               CmdConsultar_Click

End Sub

Private Sub CmdContabilizarCompras_Click()

  Dim Periodo As Double, NumeroPeriodo As Double, FechaIni As String, FechaFin As String, EstadoPeriodo As String, NumeroTransaccion As Double
  Dim mes As Double, Año As Double, Moneda, Resultado As Boolean
  Dim SqlString As String, FechaInicio As String
  Dim TipoFactura As String, NumeroFactura As String, CodigoProducto As String, CodigoCuentaProducto As String, CodigoCliente As String
  Dim NombreCuenta As String, CodigoCuentaCliente As String, SubTotal As Double, Descuento As Double, Iva As Double, NetoPagar As Double
  Dim FechaFactura As Date, FechaVence As Date, DescripcionMovimiento As String, TasaIva As Double
  Dim MonedaFactura As String, Reg As Double, CodigoCuentaIngresos As String, CodCuentaEfectivo As String
  Dim CostoProducto As Double, CodigoCuentaCostos As String, CodigoCuentaInventario As String, CodigoCuentaOtros As String, CodigoCuentaMetodo As String
  Dim Pagado As Double, NumeroReferencia As String, MonedaMovimiento As String, TasaMovimiento As Double
  Dim cn As New ADODB.Connection, SuReferencia As String, CodigoCuentaBanco As String
  Dim rs As New ADODB.Recordset, Registro As Double, NumeroCompra As String
  Dim cmd As New ADODB.Command, Ret1Porc As Double, Ret2Porc As Double, MontoRetencion1 As Double, MontoRetencion2 As Double, Ret3Porc As Double, Ret4Porc As Double, MontoRetencion3 As Double, MontoRetencion4 As Double
  Dim MontoBanco As Double
  
  Reg = 1
  
  
  Me.CmdContabilizarCompras.Enabled = False
  
            If Me.OptCompras.Value = True Then
             TipoFactura = "Mercancia Recibida"
            ElseIf Me.OptDevolucionCompra.Value = True Then
             TipoFactura = "Devolucion de Compra"
            ElseIf Me.OptTransferenciaRecibida.Value = True Then
             TipoFactura = "Transferencia Recibida"
            ElseIf Me.RadioButton4.Value = True Then
               TipoFactura = "Pago Proveedor"
            ElseIf Me.OptCuenta.Value = True Then
               TipoFactura = "Cuenta"
            End If


                Select Case TipoFactura
                    Case "Cuenta"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Observaciones, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra Like '%" & TipoFactura & "%') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                    Case "Pago Proveedor"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       SqlString = "SELECT ReciboPago.Fecha_Recibo, ReciboPago.CodReciboPago, ReciboPago.Cod_Proveedor, Proveedor.Nombre_Proveedor, ReciboPago.Sub_Total, ReciboPago.Descuento , ReciboPago.Total, ReciboPago.Marca, Proveedor.Cod_Cuenta_Pagar,ReciboPago.MonedaRecibo, ReciboPago.Retencion1, ReciboPago.Retencion2, ReciboPago.Retencion3, ReciboPago.Retencion4, ReciboPago.Retencion5, ReciboPago.Retencion6, ReciboPago.Retencion7 FROM ReciboPago INNER JOIN Proveedor ON ReciboPago.Cod_Proveedor = Proveedor.Cod_Proveedor  " & _
                                   "WHERE (ReciboPago.Fecha_Recibo BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (ReciboPago.Marca = 1) AND (ReciboPago.Contabilizado = 0)"
                
                    Case "Transferencia Recibida"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                       
                    Case "Recepcion"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                    Case "Mercancia Recibida"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Observaciones, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                   Case "Devolucion de Compra"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       
                        SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                 
                 End Select

                 '//////////////////////////////////////////////////////////////////////////////////////////////
                 '//////////SI LA CUENTA EXISTE AGREGO LOS ENCABEZADOS///////////////////////////////////////
                 '/////////////////////////////////////////////////////////////////////////////////////////////
                 

                         mes = Month(Me.DTPicker6.Value)
                         Año = Year(Me.DTPicker6.Value)
                         FechaIni = CDate("1/" & Month(Me.DTPicker6.Value) & "/" & Year(Me.DTPicker6.Value))
                         FechaFin = DateSerial(Año, mes + 1, 1 - 1)

                 
                         Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
                         Me.AdoConsulta.Refresh
                         If Not Me.AdoConsulta.Recordset.EOF Then
                           Periodo = Me.AdoConsulta.Recordset("Periodo")
                            NumeroPeriodo = Me.AdoConsulta.Recordset("NPeriodo")
                            EstadoPeriodo = Me.AdoConsulta.Recordset("EstadoPeriodo")
                      


                              Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                              Me.AdoConsulta.Recordset.Update
                              NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                              

                               
 
                             Me.AdoProcesos.RecordSource = SqlString
                             Me.AdoProcesos.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                             If Not Me.AdoProcesos.Recordset.EOF Then
                             Me.AdoProcesos.Recordset.MoveFirst
                             Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                             End If
                             Me.AdoProcesos.Refresh
                             
                             Beneficiario = ""
                             Ret1Porc = 0
                             Ret2Porc = 0
                             Ret3Porc = 0
                             Ret4Porc = 0
                             
                             
                             
                        Do While Not Me.AdoProcesos.Recordset.EOF
                           
                           Select Case TipoFactura
                              Case "Pago Proveedor"
                              
                                MontoBanco = 0
                                FechaFactura = Me.AdoProcesos.Recordset("Fecha_Recibo")
                                FechaVence = Me.AdoProcesos.Recordset("Fecha_Recibo")
                                NumeroFactura = Me.AdoProcesos.Recordset("CodReciboPago")
                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                
                                Ret1Porc = Mid(Me.AdoProcesos.Recordset("Retencion1"), 1, Len(Me.AdoProcesos.Recordset("Retencion1")) - 1)
                                Ret2Porc = Mid(Me.AdoProcesos.Recordset("Retencion2"), 1, Len(Me.AdoProcesos.Recordset("Retencion2")) - 1)
                                Ret3Porc = Mid(Me.AdoProcesos.Recordset("Retencion3"), 1, Len(Me.AdoProcesos.Recordset("Retencion3")) - 1)
                                Ret4Porc = Mid(Me.AdoProcesos.Recordset("Retencion4"), 1, Len(Me.AdoProcesos.Recordset("Retencion4")) - 1)
                                
                                If CodigoCuentaCliente = "" Then
                                  MsgBox "No existe la cuenta del Proveedor", vbCritical, "Zeus contable"
                                  Exit Sub
                                End If
                                
                                    SubTotal = 0
                                    Descuento = 0
                                    Iva = 0
                                    NetoPagar = 0
                                    Pagado = 0
                    

                                Me.AdoConsultaFactura.RecordSource = "SELECT DetalleReciboPago.idDetallePago, DetalleReciboPago.CodReciboPago, DetalleReciboPago.Fecha_Recibo, DetalleReciboPago.Numero_Compra, DetalleReciboPago.MontoPagado , Detalle_MetodoPagoProveedores.NombrePago, Detalle_MetodoPagoProveedores.Monto, MetodoPago.Cod_Cuenta FROM DetalleReciboPago INNER JOIN ReciboPago ON DetalleReciboPago.CodReciboPago = ReciboPago.CodReciboPago AND DetalleReciboPago.Fecha_Recibo = ReciboPago.Fecha_Recibo INNER JOIN Detalle_MetodoPagoProveedores ON DetalleReciboPago.CodReciboPago = Detalle_MetodoPagoProveedores.CodReciboPago INNER JOIN MetodoPago ON Detalle_MetodoPagoProveedores.NombrePago = MetodoPago.NombrePago  " & _
                                                                     "WHERE (DetalleReciboPago.CodReciboPago = '" & NumeroFactura & "') AND (DetalleReciboPago.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) ORDER BY DetalleReciboPago.Fecha_Recibo"
                                Me.AdoConsultaFactura.Refresh
                                If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                     DescripcionRecibo = "Pago a Proveedores"  'Me.AdoConsultaFactura.Recordset("Descripcion")
                                  If Not IsNull(Me.AdoConsultaFactura.Recordset("MontoPagado")) Then
                                    If Me.AdoConsultaFactura.Recordset("MontoPagado") = Me.AdoProcesos.Recordset("Sub_Total") Then
                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("Sub_Total")), "##,##0.00")
                                        If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                          Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                        End If
                                        NetoPagar = Format(Val(Me.AdoProcesos.Recordset("Total")), "##,##0.00")
                                    Else
                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("Sub_Total")), "##,##0.00")
                                        If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                         Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                        End If
                                        NetoPagar = Me.AdoConsultaFactura.Recordset("MontoPagado")
                                        


                                    End If
                               
                                
                                
                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                       MonedaFactura = Me.AdoProcesos.Recordset("MonedaRecibo")
                                       TasaMovimiento = 1
    
                                        SqlString = "SELECT * From TasaCambio WHERE (FechaTasa = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102))"
                                        Me.AdoBuscaFacturacion.RecordSource = SqlString
                                        Me.AdoBuscaFacturacion.Refresh
                                        If Not Me.AdoBuscaFacturacion.Recordset.EOF Then
                                          TasaCambio = Format(Val(Me.AdoBuscaFacturacion.Recordset("MontoTasa")), "##,##0.0000")
                                        Else
                                          TasaCambio = 1
                                        End If
                                   End If
                                 End If
                                 
                                      If MonedaFactura = "Cordobas" Then
                                        TasaCambio = 1
                                        MonedaMovimiento = "Córdobas"
                                      End If
                                 
                                 
                                      If Reg = 1 Then
                                         '////////////////////////////////////////////////////////////////
                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                         '/////////////////////////////////////////////////////////////////
                                         MonedaMovimiento = "Córdobas"
                                         If Me.ChkCheques.Value = xtpChecked Then
                                           Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker6.Value, "yyyy-mm-dd"), "Movimiento de Pago a Proveedores", "CHEQUE", MonedaMovimiento)
                                         Else
                                           Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker6.Value, "yyyy-mm-dd"), "Movimiento de Pago a Proveedores", "Pago", MonedaMovimiento)
                                         End If
                                         Reg = 2
                                      End If
                                      
                                       DescripcionMovimiento = "Registrando Recibo No " & NumeroFactura & "  " & DescripcionRecibo
                                      

                                      '///////////////////////////////////////////////////////////////////////////////////////
                                      '/////////////////////////////GRABA DETALLE DE RECIBO PAGOS///////////////////////////////////
                                      '///////////////////////////////////////////////////////////////////////////////////////
                                       TotalRetencion = 0
                                       Registro = 1
                                       CodigoCuentaBanco = ""
'                                       Me.AdoConsultaFactura.RecordSource = "SELECT  DetalleReciboPago.idDetallePago, DetalleReciboPago.CodReciboPago, DetalleReciboPago.Fecha_Recibo, DetalleReciboPago.Numero_Compra, DetalleReciboPago.MontoPagado  FROM  DetalleRecibo INNER JOIN MetodoPago ON DetalleRecibo.NombrePago = MetodoPago.NombrePago  " & _
'                                                                            "WHERE  (DetalleRecibo.CodReciboPago = '" & NumeroFactura & "') AND (DetalleRecibo.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) AND (DetalleRecibo.MontoPagado < 0) ORDER BY DetalleRecibo.Fecha_Recibo"
                                       Me.AdoConsultaFactura.RecordSource = "SELECT DetalleReciboPago.idDetallePago, DetalleReciboPago.CodReciboPago, DetalleReciboPago.Fecha_Recibo, DetalleReciboPago.Numero_Compra, DetalleReciboPago.MontoPagado , Detalle_MetodoPagoProveedores.NombrePago, Detalle_MetodoPagoProveedores.Monto, MetodoPago.Cod_Cuenta FROM DetalleReciboPago INNER JOIN ReciboPago ON DetalleReciboPago.CodReciboPago = ReciboPago.CodReciboPago AND DetalleReciboPago.Fecha_Recibo = ReciboPago.Fecha_Recibo INNER JOIN Detalle_MetodoPagoProveedores ON DetalleReciboPago.CodReciboPago = Detalle_MetodoPagoProveedores.CodReciboPago INNER JOIN MetodoPago ON Detalle_MetodoPagoProveedores.NombrePago = MetodoPago.NombrePago  " & _
                                                                            "WHERE (DetalleReciboPago.CodReciboPago = '" & NumeroFactura & "') AND (DetalleReciboPago.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) ORDER BY DetalleReciboPago.Fecha_Recibo"
                                       Me.AdoConsultaFactura.Refresh
                                   Do While Not Me.AdoConsultaFactura.Recordset.EOF
                                      
                                       NumeroCompra = Me.AdoConsultaFactura.Recordset("Numero_Compra")
                                      
                                       Debito = Abs(Me.AdoConsultaFactura.Recordset("MontoPagado"))
                                       Debito = Format(Debito * TasaCambio, "##,##0.00")
                                       Credito = 0
                                       CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                        
'                                        DescripcionMovimiento = "Movimiento de Registro de Recibos No" & NumeroFactura
                                        If Me.ChkCheques.Value = xtpChecked Then
                                         Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroCompra, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                        Else
                                         Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Pago", NumeroCompra, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ReciboPago")
                                        End If
                                        

                                        CodigoCuentaBanco = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                                               
                                       
                                        Me.AdoConsultaFactura.Recordset.MoveNext
                                        TotalRetencion = Abs(Debito) + TotalRetencion
                                     Loop
                                     
                                     
                                     
                                      
                                     If Registro = 1 Then
                                                     
                                             MontoBanco = SubTotal - (SubTotal * (Ret1Porc / 100)) - (SubTotal * (Ret2Porc / 100)) - (SubTotal * (Ret3Porc / 100)) - (SubTotal * (Ret4Porc / 100))
                                            
                                             NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                             Debito = 0
                                             Credito = SubTotal
                                             Credito = Format(MontoBanco * TasaCambio, "##,##0.00")
                                              
                                             
                                            '///////////////////////////////////////////////////////////////////////////////////
                                            '///////////////////////GRABO EL MOVIMIENTO DEL BANCO ///////////////////////////
                                            '///////////////////////////////////////////////////////////////////////////////////
                                             If Me.ChkCheques.Value = xtpChecked Then
                                               Beneficiario = Me.AdoProcesos.Recordset("Nombre_Proveedor")
                                               Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             Else
                                               Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Pago", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ReciboPago")
                                             End If
                                             Registro = 2
                                       
                                       
                                             '////////////////////////////////////////////////////////////////////////////
                                             '//////////////////////////AGREGO RETENCIONES /////////////////////////////////
                                             '////////////////////////////////////////////////////////////////////////////////
                                             If Ret1Porc <> 0 Then
                                              CodigoCuentaBanco = BuscaCuentaImpuestos(Ret1Porc & "%")
                                              NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                              Debito = 0
                                             
                                              Credito = Format(SubTotal * (Ret1Porc / 100) * TasaCambio, "##,##0.00")
                                              Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                             
                                             If Ret2Porc <> 0 Then
                                               CodigoCuentaBanco = BuscaCuentaImpuestos(Ret2Porc & "%")
                                               NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                               Debito = 0
                                               Credito = Format(SubTotal * (Ret2Porc / 100) * TasaCambio, "##,##0.00")
                                               Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                             
                                            If Ret3Porc <> 0 Then
                                              CodigoCuentaBanco = BuscaCuentaImpuestos(Ret3Porc & "%")
                                              NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                              Debito = 0
                                             
                                              Credito = Format(SubTotal * (Ret3Porc / 100) * TasaCambio, "##,##0.00")
                                              Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                       
                                            If Ret4Porc <> 0 Then
                                              CodigoCuentaBanco = BuscaCuentaImpuestos(Ret4Porc & "%")
                                              NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                              Debito = 0
                                             
                                              Credito = Format(SubTotal * (Ret4Porc / 100) * TasaCambio, "##,##0.00")
                                              Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                       End If

                                      

                                      
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                rs.Open "UPDATE ReciboPago SET Contabilizado = 1 ,Activo = 0  WHERE (CodReciboPago = '" & NumeroFactura & "') ", ConexionFacturacion
                                      
                                      
                              
                              Case Else
                             
                             
                             
                                                FechaFactura = Me.AdoProcesos.Recordset("Fecha_Compra")
                                                If Not IsNull(Me.AdoProcesos.Recordset("Fecha_Vencimiento")) Then
                                                  FechaVence = Me.AdoProcesos.Recordset("Fecha_Vencimiento")
                                                End If
                                                NumeroFactura = Me.AdoProcesos.Recordset("Numero_Compra")
                                                MonedaFactura = Me.AdoProcesos.Recordset("MonedaCompra")
                                                If Not IsNull(Me.AdoProcesos.Recordset("Su_Referencia")) Then
                                                   NumeroReferencia = "Compra No:" & NumeroFactura & " " & "Referencia: " & Me.AdoProcesos.Recordset("Su_Referencia")
                                                Else
                                                   NumeroReferencia = "Compra No:" & NumeroFactura
                                                End If
                                                
                                                If Not IsNull(Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")) Then
                                                  CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                                Else
                                                  CodigoCuentaCliente = "2121"
                                                End If
                                                
                                                Select Case TipoFactura
                                                  Case "Cuenta"
'                                                     codigocuentacliente =
                                                
                                                End Select
                                                
                                                
                                                    SubTotal = 0
                                                    Descuento = 0
                                                    Iva = 0
                                                    NetoPagar = 0
                                                    TasaCambio = 1
                                                    TasaMovimiento = 1
                                                    
                                                    
                                                    SqlString = "SELECT * From Detalle_Compras WHERE (Numero_Compra = '" & NumeroFactura & "') AND (Tipo_Compra = '" & TipoFactura & "')"
                                                    Me.AdoBuscaFacturacion.RecordSource = SqlString
                                                    Me.AdoBuscaFacturacion.Refresh
                                                    
                                                    If MonedaFactura = "Dolares" Then
                                                         TasaCambio = Format(Val(Me.AdoBuscaFacturacion.Recordset("TasaCambio")), "##,##0.000000")
                                                    Else
                                                         TasaCambio = 1
                                                    End If
                                                
                
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////////////////BUSCO LOS DETALLES DE LACOMPRA ////////////////////////////////////////////////////
                                                '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                Me.AdoConsultaFactura.RecordSource = "SELECT SUM(Detalle_Compras.Cantidad) AS Cantidad, SUM(Detalle_Compras.Precio_Unitario) AS Precio_Unitario, SUM(Detalle_Compras.Descuento) AS Descuento, SUM(Detalle_Compras.Precio_Neto) AS Precio_Neto, SUM(Detalle_Compras.Importe) AS Importe FROM  Compras INNER JOIN Detalle_Compras ON Compras.Numero_Compra = Detalle_Compras.Numero_Compra AND Compras.Fecha_Compra = Detalle_Compras.Fecha_Compra And Compras.Tipo_Compra = Detalle_Compras.Tipo_Compra  " & _
                                                                                      "WHERE (Compras.Numero_Compra = '" & NumeroFactura & "') AND (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Tipo_Compra = '" & TipoFactura & "')"
                                                Me.AdoConsultaFactura.Refresh
                                                If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                                  If Not IsNull(Me.AdoConsultaFactura.Recordset("Importe")) Then
                                                    If Format(CDbl(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00") = Format(CDbl(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00") Then
                                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00")
                                                        SubTotal = Format(SubTotal * TasaCambio, "##,##0.00")
                                                        If Not IsNull(Me.AdoProcesos.Recordset("Descuento")) Then
                                                         If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                                          Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                                          Descuento = Format(Descuento * TasaCambio, "##,##0.00")
                                                         End If
                                                        End If
                                                        Iva = Format(Val(Me.AdoProcesos.Recordset("IVA")), "##,##0.00")
                                                        Iva = Format(Iva * TasaCambio, "##,##0.00")
                                                        NetoPagar = Format(Val(Me.AdoProcesos.Recordset("NetoPagar")), "##,##0.00")
                                                        NetoPagar = Format(NetoPagar * TasaCambio, "##,##0.00")
                                                        Pagado = Format(Val(Me.AdoProcesos.Recordset("Pagado")), "##,##0.00")
                                                        Pagado = Format(Pagado * TasaCambio, "##,##0.00")
                                                    Else
                                                        SubTotal = Format(Val(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00")
                                                        SubTotal = Format(SubTotal * TasaCambio, "##,##0.00")
                                                        Descuento = 0
                                                        Iva = 0
                                                        NetoPagar = 0
                                                        Pagado = 0
                                                    End If
                                               
                                                
                                                
                                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                       MonedaFactura = Me.AdoProcesos.Recordset("MonedaCompra")
                                                       
                
                
                    
                
                                                
                                                      If Reg = 1 Then
                                                         '////////////////////////////////////////////////////////////////
                                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                                         '/////////////////////////////////////////////////////////////////
                                                         MonedaMovimiento = "Córdobas"
                                                         Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker6.Value, "yyyy-mm-dd"), "Movimiento de Compras", "Comp", MonedaMovimiento)
                                                         Reg = 2
                                                      End If
                                                           
                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                        '////////////AGREGO LA CUENTA DEL PROVEEDOR//////////////////////////////////////
                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                           
                                                           Credito = 0
                                                           Debito = 0
                                                          If NetoPagar = 0 Then
                                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                                                                                      
                                                            Me.AdoConsultaFactura.RecordSource = "SELECT Detalle_MetodoCompras.Numero_Compra, Detalle_MetodoCompras.Fecha_Compra, Detalle_MetodoCompras.Tipo_Compra, Detalle_MetodoCompras.NombrePago, Detalle_MetodoCompras.Monto, Detalle_MetodoCompras.NumeroTarjeta, Detalle_MetodoCompras.FechaVence , MetodoPago.TipoPago, MetodoPago.Cod_Cuenta, MetodoPago.Moneda FROM Detalle_MetodoCompras INNER JOIN MetodoPago ON Detalle_MetodoCompras.NombrePago = MetodoPago.NombrePago  " & _
                                                                                                 "WHERE (Detalle_MetodoCompras.Numero_Compra = '" & NumeroFactura & "')"
                                                            Me.AdoConsultaFactura.Refresh
                                                            If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                                                CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                                
                                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                                NetoPagar = SubTotal + Iva - Descuento
                                                                Select Case TipoFactura
                                                                  Case "Transferencia Recibida"
                                                                       DescripcionMovimiento = "Registro de Transferencia Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                                                                  
                                                                  
                '                                                       DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Mercancia Recibida"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                                                                  
                                                                  
                '                                                       DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de Devolucion Numero " & NumeroFactura
                                                                       Debito = Format(NetoPagar, "##,##0.00")
                                                                       Credito = 0
                                                                End Select
                                                            Else
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                                NetoPagar = SubTotal + Iva - Descuento
                                                                Debito = Format(NetoPagar, "##,##0.00")
                                                            End If
                                                           ElseIf NetoPagar < 0 Then
                                                              '/////////////////AGREGO LA DIFERENICIA A OTROS INGRESOS /
                                                               Me.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (TipoCuenta = 'Capital')"
                                                               Me.AdoConsulta.Refresh
                                                               If Not Me.AdoConsulta.Recordset.EOF Then
                                                                CodigoCuentaOtros = Me.AdoConsulta.Recordset("CodCuentas")
                                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaOtros)
                                                               End If
                                                               
                                                               
                                                               Select Case TipoFactura
                                                                  Case "Transferencia Recibida"
                                                                       DescripcionMovimiento = "Registro de AJUSTE Transferencia Recibida Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Mercancia Recibida"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Devolucion Compra Numero " & NumeroFactura
                                                                       Debito = Format(NetoPagar, "##,##0.00")
                                                                       Credito = 0
                                                                End Select
                                                               
                                                               
                                                               Resultado = GrabaDetalleFactura(CodigoCuentaOtros, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, 0, Abs(NetoPagar), "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                               
                                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                                            Me.AdoConsultaFactura.RecordSource = "SELECT  * FROM Detalle_MetodoFacturas INNER JOIN MetodoPago ON Detalle_MetodoFacturas.NombrePago = MetodoPago.NombrePago  " & _
                                                                                                 "WHERE (Detalle_MetodoFacturas.Numero_Factura = '" & NumeroFactura & "')"
                                                            Me.AdoConsultaFactura.Refresh
                '                                              CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                              NetoPagar = SubTotal + Iva - Descuento + Abs(NetoPagar)
                                                              
                                                              Debito = Format(NetoPagar, "##,##0.00")
                                                              DescripcionMovimiento = "Registro de Facturacion Factura Numero " & NumeroFactura
                                                              DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                              
                                                           ElseIf NetoPagar > 0 Then
                                                           
                                              
                                                            DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                           
                                                           Select Case TipoFactura
                                                            Case "Transferencia Recibida"
                                                                   DescripcionMovimiento = "Registro de Transferencia Recibida Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = 0
                                                                   Credito = Format(NetoPagar, "##,##0.00")
                
                                                                  '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             Case "Mercancia Recibida"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Compra Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = 0
                                                                   Credito = Format(NetoPagar, "##,##0.00")
                                                                   
                                                                  DescripcionMovimiento = DescripcionMovimiento & "," & Me.AdoProcesos.Recordset("Observaciones")
                                                                  '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             
                                                              Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Compra Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = 0
                                                                   Credito = Format(NetoPagar, "##,##0.00")
                                                                   
                                                                  DescripcionMovimiento = DescripcionMovimiento & "," & Me.AdoProcesos.Recordset("Observaciones")
                                                                  '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             
                                                             Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Devolucion Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = Format(NetoPagar, "##,##0.00")
                                                                   Credito = 0
                                                        '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             End Select
                                                                
                
                                                           End If
                 
'++++++++++++++++++++++++++++++++++++++++++ DETALLE DE MOVIMIENTOS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                                    '//////////////////////////////////////////////////////////////////////////////////////////
                                                    '/////////////////////CARGO LOS DETALLE DE LAS FACTURAS////////////////////////////////////
                                                    '//////////////////////////////////////////////////////////////////////////////////////////
                                                     
                                                     If TipoFactura = "Cuenta" Then
                                                        SqlString = "SELECT  Numero_Compra, Fecha_Compra, Tipo_Compra, Cod_Producto, Cantidad, Precio_Unitario, Descuento, Precio_Neto, Importe, TasaCambio From Detalle_Compras WHERE (Numero_Compra = '" & NumeroFactura & "') AND (Tipo_Compra LIKE '%" & TipoFactura & "%')"
                                                     
                                                     Else
                                                         SqlString = "SELECT  Detalle_Compras.Numero_Compra, Detalle_Compras.Fecha_Compra, Detalle_Compras.Tipo_Compra, Detalle_Compras.Cod_Producto, Detalle_Compras.Cantidad, Detalle_Compras.Precio_Unitario, Detalle_Compras.Descuento, Detalle_Compras.Precio_Neto, Detalle_Compras.Importe, Productos.Descripcion_Producto, Productos.Cod_Cuenta_Inventario, Productos.Cod_Cuenta_Costo, Productos.Cod_Cuenta_Ventas, Productos.Cod_Cuenta_GastoAjuste , Productos.Cod_Cuenta_IngresoAjuste, Detalle_Compras.TasaCambio, Productos.Costo_Promedio, Productos.Costo_Promedio_Dolar  FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos  " & _
                                                                     "WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "') AND (Detalle_Compras.Tipo_Compra Like '%" & TipoFactura & "%')"
                                                     End If
                                                     Me.AdoProcesosFacturacion.RecordSource = SqlString
                                                     Me.AdoProcesosFacturacion.Refresh
                                                     Do While Not Me.AdoProcesosFacturacion.Recordset.EOF
                                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                If TipoFactura = "Cuenta" Then
                                                                  CodigoCuentaProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                Else
                                                                  CodigoCuentaProducto = BuscaCodigoProducto(CodigoProducto)
                                                                End If
                                                                If MonedaFactura = "Dolares" Then
                '                                                 TasaCambio = BuscaTasaCambio(Fecha)
                                                                   TasaCambio = Format(Val(Me.AdoProcesosFacturacion.Recordset("TasaCambio")), "##,##0.0000")
                                                                Else
                                                                    TasaCambio = 1
                                                               End If
                                                                
                                                                
                                                                If MonedaFactura = "Dolares" Then
                                                                  CostoProducto = Me.AdoProcesosFacturacion.Recordset("Importe")
                '                                                   CostoProducto = Format(Me.AdoProcesosFacturacion.Recordset("Importe") * TasaCambio, "##,##0.00")
                                                                Else
                                                                   CostoProducto = Format(Me.AdoProcesosFacturacion.Recordset("Importe"), "##,##0.00")
                                                                End If
                                                            
                                                               '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                               '///////////////////////////CALCULO EL IVA DE CADA PRODUCTO//////////////////////////////////////////////
                                                               '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                
                                                                If Iva <> 0 Then
                                                                 TasaIva = BuscaTasaIva(CodigoProducto)
                                                                 DescripcionCuenta = BuscaCuenta(CodigoCuentaIva)
                                                                 
                                                                 QUIEN = ""
                                                                 
                                                                 If DescripcionCuenta <> "Nulo" Then
                                                                  
                                                                  Select Case TipoFactura
                                                                    Case "Transferencia Recibida"
                                                                         DescripcionMovimiento = "IVA Ventas Compra No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                         Credito = 0
                                                                         Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                            If Debito <> 0 Then
                                                                                Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                            End If
                                                                    Case "Mercancia Recibida"
                                                                         DescripcionMovimiento = "IVA Ventas Compra No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                         Credito = 0
                                                                         Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                            If Debito <> 0 Then
                                                                                Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                            End If
                                                                            
                                                                     Case "Cuenta"
                                                                         DescripcionMovimiento = "IVA Ventas Compra No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                         Credito = 0
                                                                         Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                            If Debito <> 0 Then
                                                                                Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                            End If
                                                                    Case "Devolucion de Compra"
                                                                      DescripcionMovimiento = "IVA Ventas Devolucion No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                      Debito = 0
                                                                      Credito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                      Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                                     If Credito <> 0 Then
                                                                         Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                     End If
                                                                    End Select
                
                
                                                                   End If
                                                                 End If
                                                                
                                                              QUIEN = ""
                                                                
                                                                 
                                                                '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '///////////////////////////////////////BUSCO LA CUENTA DE INVENTARIO///////////////////////////////////////////
                                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                If TipoFactura <> "Cuenta" Then
                                                                    If Not IsNull(Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")) Then
                                                                      CodigoCuentaInventario = Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")
                                                                    End If
                                                                Else
                                                                    CodigoCuentaInventario = CodigoCuentaProducto
                                                                End If
                                                                
                                                                 
                                                                 DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
                                                                   Select Case TipoFactura
                                                                    Case "Transferencia Recibida"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                          Credito = 0
                                                                         'If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         'End If
                                                                    Case "Mercancia Recibida"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                          Credito = 0
                                                                         'If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         'End If
                                                                
                                                                    Case "Cuenta"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                          Credito = 0
                                                                         'If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         'End If
                                                                    Case "Devolucion de Compra"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = 0
                                                                          Credito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Credito = Format(Val(CostoProducto) * TasaCambio, "##,##0.00")
                                                                         If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         End If
                                                                    End Select
                            
                            
                                                        Me.AdoProcesosFacturacion.Recordset.MoveNext
                                                     Loop
                                                     
                                                     
                                                                 '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '///////////////////////////////////////BUSCO LA CUENTA DE INVENTARIO///////////////////////////////////////////
                                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                If Pagado <> 0 Then
                                                                  SqlString = "SELECT Detalle_MetodoCompras.Numero_Compra, Detalle_MetodoCompras.Fecha_Compra, Detalle_MetodoCompras.NombrePago, Detalle_MetodoCompras.Monto,Detalle_MetodoCompras.NumeroTarjeta , MetodoPago.Cod_Cuenta, MetodoPago.Moneda FROM  Detalle_MetodoCompras INNER JOIN MetodoPago ON Detalle_MetodoCompras.NombrePago = MetodoPago.NombrePago " & _
                                                                              "WHERE (Detalle_MetodoCompras.Numero_Compra = '" & NumeroFactura & "') "
                                                                  Me.AdoConsultaFacturacion.RecordSource = SqlString
                                                                  Me.AdoConsultaFacturacion.Refresh
                                                                  Do While Not Me.AdoConsultaFacturacion.Recordset.EOF
                                                                     CodigoCuentaMetodo = Me.AdoConsultaFacturacion.Recordset("Cod_Cuenta")
                                                                     
                                                                     DescripcionCuenta = BuscaCuenta(CodigoCuentaMetodo)
                                                                    Select Case TipoFactura
                                                                        Case "Mercancia Recibida"
                                                                         DescripcionMovimiento = "PAGO DE COMPRA" & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                         Debito = 0
                                                                         Credito = Me.AdoConsultaFacturacion.Recordset("Monto")
                                                                             If DescripcionCuenta <> "Nulo" Then
                                                                              Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                             End If
                                                                        Case "Cuenta"
                                                                         DescripcionMovimiento = "PAGO DE COMPRA" & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                         Debito = 0
                                                                         Credito = Me.AdoConsultaFacturacion.Recordset("Monto")
                                                                             If DescripcionCuenta <> "Nulo" Then
                                                                              Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                             End If
                                                                        Case "Devolucion de Compra"
                                                                         DescripcionMovimiento = "PAGO DE DEVOLUCION" & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                         Debito = Me.AdoConsultaFacturacion.Recordset("Monto")
                                                                         Credito = 0
                                                                             If DescripcionCuenta <> "Nulo" Then
                                                                              Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                             End If
                                                                    End Select
                                                                    
                                                                    
                                                                
                
                                                                     If Debito <> 0 Then
                                                                         Resultado = GrabaDetalleFactura(CodigoCuentaMetodo, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "VTAS", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                     End If
                                                                 
                                                                    Me.AdoConsultaFacturacion.Recordset.MoveNext
                                                                  Loop
                                                                
                                                                End If
                                                                
                                                                
                                                                
                                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                rs.Open "UPDATE Compras SET Contabilizado = 1 ,Activo = 0  WHERE (Numero_Compra = '" & NumeroFactura & "') ", ConexionFacturacion
                                                     
                                                     
                                                     
                                                  End If
                                                
                                                End If
                
                                End Select
                                                                                       
                                Me.osProgress1.Value = Me.osProgress1.Value + 1
                                Me.AdoProcesos.Recordset.MoveNext
                            Loop
                             
'

                         End If

End Sub

Private Sub CmdContabilizarNotas_Click()
  Dim Periodo As Double, NumeroPeriodo As Double, FechaIni As String, FechaFin As String, EstadoPeriodo As String, NumeroTransaccion As Double
  Dim mes As Double, Año As Double, Moneda, Resultado As Boolean
  Dim SqlString As String, FechaInicio As String
  Dim TipoFactura As String, NumeroFactura As String, CodigoProducto As String, CodigoCuentaProducto As String, CodigoCliente As String
  Dim NombreCuenta As String, CodigoCuentaCliente As String, SubTotal As Double, Descuento As Double, Iva As Double, NetoPagar As Double
  Dim Fecha As Date, FechaVence As Date, DescripcionMovimiento As String, TasaIva As Double
  Dim MonedaFactura As String, Reg As Double, CodigoCuentaIngresos As String, CodCuentaEfectivo As String
  Dim CostoProducto As Double, CodigoCuentaCostos As String, CodigoCuentaInventario As String, CodigoCuentaOtros As String, CodigoCuentaMetodo As String
  Dim Pagado As Double, NumeroReferencia As String, MonedaMovimiento As String, TasaMovimiento As Double
  Dim cn As New ADODB.Connection, SuReferencia As String, NumeroNota As String, CodigoCuentaNota As String
  Dim rs As New ADODB.Recordset, CodigoTipoNota As String
  Dim cmd As New ADODB.Command, Fuente As String
  
  Reg = 1
  Monto = 0
  
  Me.CmdContabilizarNotas.Enabled = False
  
If Me.OptNotaDebito.Value = True Then
 TipoNota = "Debito Clientes"
ElseIf Me.OptNotaCredito.Value = True Then
 TipoNota = "Credito Clientes"
ElseIf Me.OptNotaCreditoProveedor.Value = True Then
    TipoNota = "Credito Proveedores"
ElseIf Me.OptNotaDebitoProveedor.Value = True Then
    TipoNota = "Debito Proveedores"
ElseIf Me.OptPlanillaProductor.Value = True Then
    TipoNota = "PlanillaLeche"
End If

     Select Case TipoNota
       Case "Debito Clientes"

         
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
            SqlString = "SELECT IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto , IndiceNota.Marca, NotaDebito.CuentaContable, Clientes.Cod_Cuenta_Cliente, IndiceNota.Tipo_Nota AS CodTipoNota, Clientes.RUC, " & _
                        "IndiceNota.Observaciones FROM IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB INNER JOIN Clientes ON IndiceNota.Cod_Cliente = Clientes.Cod_Cliente  " & _
                        "WHERE (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (NotaDebito.Tipo = 'Debito Clientes') AND (IndiceNota.Marca = 1) AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY IndiceNota.Fecha_Nota"
       
       Case "Credito Clientes"

          
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
                        SqlString = "SELECT IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto , IndiceNota.Marca, NotaDebito.CuentaContable, Clientes.Cod_Cuenta_Cliente, IndiceNota.Tipo_Nota AS CodTipoNota,Clientes.RUC, " & _
                                    "IndiceNota.Observaciones FROM IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB INNER JOIN Clientes ON IndiceNota.Cod_Cliente = Clientes.Cod_Cliente  " & _
                                    "WHERE (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (NotaDebito.Tipo = 'Credito Clientes') AND (IndiceNota.Marca = 1) AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) ORDER BY IndiceNota.Fecha_Nota"
      
      
       Case "Debito Proveedores"

         
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
            SqlString = "SELECT  IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto, IndiceNota.Marca, NotaDebito.CuentaContable, IndiceNota.Tipo_Nota AS CodTipoNota, IndiceNota.Observaciones, Proveedor.Cod_Proveedor, Proveedor.Cod_Cuenta_Proveedor , Proveedor.RUC, Proveedor.Cod_Cuenta_Pagar As Cod_Cuenta_Cliente FROM  IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN  NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB INNER JOIN Proveedor ON IndiceNota.Cod_Cliente = Proveedor.Cod_Proveedor  " & _
                        "WHERE  (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (NotaDebito.Tipo Like '%Debito Proveedores%') AND (IndiceNota.Marca = 1) ORDER BY IndiceNota.Fecha_Nota"
       
       Case "Credito Proveedores"

         
            
            FechaInicio = Format(Me.DTPicker7.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker8.Value, "yyyy-mm-dd")
            SqlString = "SELECT  IndiceNota.Numero_Nota, IndiceNota.Fecha_Nota, IndiceNota.MonedaNota, IndiceNota.Nombre_Cliente, Detalle_Nota.Descripcion, Detalle_Nota.Numero_Factura, Detalle_Nota.Monto, IndiceNota.Marca, NotaDebito.CuentaContable, IndiceNota.Tipo_Nota AS CodTipoNota, IndiceNota.Observaciones, Proveedor.Cod_Proveedor, Proveedor.Cod_Cuenta_Proveedor , Proveedor.RUC, Proveedor.Cod_Cuenta_Pagar As Cod_Cuenta_Cliente FROM  IndiceNota INNER JOIN Detalle_Nota ON IndiceNota.Numero_Nota = Detalle_Nota.Numero_Nota AND IndiceNota.Fecha_Nota = Detalle_Nota.Fecha_Nota AND IndiceNota.Tipo_Nota = Detalle_Nota.Tipo_Nota INNER JOIN  NotaDebito ON IndiceNota.Tipo_Nota = NotaDebito.CodigoNB INNER JOIN Proveedor ON IndiceNota.Cod_Cliente = Proveedor.Cod_Proveedor  " & _
                        "WHERE  (IndiceNota.Nombre_Cliente <> '*******ANULADO*******') AND (IndiceNota.Contabilizado = 0) AND (IndiceNota.Fecha_Nota BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (NotaDebito.Tipo Like '%Credito Proveedores%') AND (IndiceNota.Marca = 1) ORDER BY IndiceNota.Fecha_Nota"
      
      End Select
    
    
                 '//////////////////////////////////////////////////////////////////////////////////////////////
                 '//////////SI LA CUENTA EXISTE AGREGO LOS ENCABEZADOS///////////////////////////////////////
                 '/////////////////////////////////////////////////////////////////////////////////////////////
                 

                         mes = Month(Me.DTPicker6.Value)
                         Año = Year(Me.DTPicker6.Value)
                         FechaIni = CDate("1/" & Month(Me.DTPicker9.Value) & "/" & Year(Me.DTPicker9.Value))
                         FechaFin = DateSerial(Año, mes + 1, 1 - 1)

                 
                         Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
                         Me.AdoConsulta.Refresh
                         If Not Me.AdoConsulta.Recordset.EOF Then
                           Periodo = Me.AdoConsulta.Recordset("Periodo")
                            NumeroPeriodo = Me.AdoConsulta.Recordset("NPeriodo")
                            EstadoPeriodo = Me.AdoConsulta.Recordset("EstadoPeriodo")
                      


                              Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                              Me.AdoConsulta.Recordset.Update
                              NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                              

                               
 
                             Me.AdoProcesos.RecordSource = SqlString
                             Me.AdoProcesos.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                             If Not Me.AdoProcesos.Recordset.EOF Then
                             Me.AdoProcesos.Recordset.MoveLast
                             Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                             End If
                             Me.AdoProcesos.Refresh
                             Do While Not Me.AdoProcesos.Recordset.EOF
                                Fecha = Me.AdoProcesos.Recordset("Fecha_Nota")
                                NumeroNota = Me.AdoProcesos.Recordset("Numero_Nota")
                                NumeroFactura = Me.AdoProcesos.Recordset("Numero_Factura")
                                MonedaFactura = Me.AdoProcesos.Recordset("MonedaNota")
                                CodigoTipoNota = Me.AdoProcesos.Recordset("CodTipoNota")
                                NumeroReferencia = NumeroNota

                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Cliente")
                                CodigoCuentaNota = Me.AdoProcesos.Recordset("CuentaContable")
                                
                                If MonedaFactura = "Dolares" Then
                                   TasaCambio = BuscaTasaCambio(Fecha)
                                Else
                                   TasaCambio = 1
                                End If
                             
                                      If Reg = 1 Then
                                        Select Case TipoNota
                                         Case "Debito Clientes"
                                           Fuente = "NDB"
                                         Case "Credito Clientes"
                                           Fuente = "NC"
                                        End Select
                                         '////////////////////////////////////////////////////////////////
                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                         '/////////////////////////////////////////////////////////////////
                                         MonedaMovimiento = "Córdobas"
                                         Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker9.Value, "yyyy-mm-dd"), "Movimiento de Nota Debito/Credito", Fuente, MonedaMovimiento)
                                         Reg = 2
                                      End If
                                      
                                 Select Case TipoNota
                                  Case "Debito Proveedores"
                                   '-------------------------AGREGO EL MOVIMIENTO DEBITO --------------------------------
                                         DescripcionMovimiento = "Nota de Debito No " & NumeroNota & " RUC:" & Me.AdoProcesos.Recordset("RUC") & ", " & Me.AdoProcesos.Recordset("Observaciones")
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                         TasaMovimiento = 1
                                         Credito = 0
                                         Debito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                         If Debito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "NDB", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
                                    '---------------------------AGREGO EL MOVIMIENTO CREDITO -------------------------------
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaNota)
                                         Credito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = 0
                                         Credito = Format(Credito * TasaCambio, "##,##0.00")
                                         If Credito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaNota, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "NDB", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
                                         
                                   Case "Debito Clientes"
                                   '-------------------------AGREGO EL MOVIMIENTO DEBITO --------------------------------
                                         DescripcionMovimiento = "Nota de Debito No " & NumeroNota & " RUC:" & Me.AdoProcesos.Recordset("RUC") & ", " & Me.AdoProcesos.Recordset("Observaciones")
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                         TasaMovimiento = 1
                                         Credito = 0
                                         Debito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                         If Debito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "NDB", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
                                    '---------------------------AGREGO EL MOVIMIENTO CREDITO -------------------------------
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaNota)
                                         Credito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = 0
                                         Credito = Format(Credito * TasaCambio, "##,##0.00")
                                         If Credito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaNota, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "NDB", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
'                                         rs.Open "UPDATE [IndiceNota]  Set [Activo] = 0 ,[Contabilizado] = 1 WHERE (IndiceNota.Numero_Nota = '" & NumeroNota & "') AND (NotaDebito.Tipo = 'Debito Clientes')", ConexionFacturacion
                                   
                                   Case "Credito Clientes"
                                   
                                    '-------------------------AGREGO EL MOVIMIENTO CREDITO --------------------------------
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaNota)
                                         TasaMovimiento = 1
                                         DescripcionMovimiento = "Nota de Credito No " & NumeroNota & " RUC:" & Me.AdoProcesos.Recordset("RUC") & ", " & Me.AdoProcesos.Recordset("Observaciones")
                                         Credito = 0
                                         Debito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                         If Debito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaNota, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "NC", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
                                    '---------------------------AGREGO EL MOVIMIENTO CREDITO -------------------------------
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                         Credito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = 0
                                         Credito = Format(Credito * TasaCambio, "##,##0.00")
                                         If Credito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "NC", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
'                                         rs.Open "UPDATE [IndiceNota]  Set [Activo] = 0 ,[Contabilizado] = 1 WHERE (IndiceNota.Numero_Nota = '" & NumeroNota & "') AND (NotaDebito.Tipo = 'Credito Clientes')"
                                  
                                   Case "Credito Proveedores"
                                   
                                    '-------------------------AGREGO EL MOVIMIENTO CREDITO --------------------------------
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaNota)
                                         TasaMovimiento = 1
                                         DescripcionMovimiento = "Nota de Credito No " & NumeroNota & " RUC:" & Me.AdoProcesos.Recordset("RUC") & ", " & Me.AdoProcesos.Recordset("Observaciones")
                                         Credito = 0
                                         Debito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                         If Debito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaNota, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "NC", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
                                    '---------------------------AGREGO EL MOVIMIENTO CREDITO -------------------------------
                                         DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                         Credito = Format(Val(Me.AdoProcesos.Recordset("Monto")), "##,##0.00")
                                         Debito = 0
                                         Credito = Format(Credito * TasaCambio, "##,##0.00")
                                         If Credito <> 0 Then
                                             Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker9.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "NC", NumeroFactura, Fecha, Descuento, Fecha, CodigoCuentaCliente, "FacturaVenta")
                                         End If
                                  
                                  End Select
                             
                               '/////////////////////////////////////////////////ACTUALIZO LAS NOTAS DE DEBITOS //////////////////////////////////
                             
                                                                             '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                rs.Open "UPDATE [IndiceNota]  Set [Activo] = 0 ,[Contabilizado] = 1 WHERE (IndiceNota.Numero_Nota = '" & NumeroNota & "') AND (IndiceNota.Tipo_Nota = '" & CodigoTipoNota & "')", ConexionFacturacion
                               
                               Me.osProgress1.Value = Me.osProgress1.Value + 1
                               Me.AdoProcesos.Recordset.MoveNext
                             Loop
                             
                                Me.AdoNota.Refresh
                                 If Not Me.AdoNota.Recordset.EOF Then
                                   Me.CmdContabilizarNotas.Enabled = True
                                   Me.DTPicker9.Visible = True
                                Else
                                   Me.CmdContabilizarNotas.Enabled = False
                                   Me.DTPicker9.Visible = False
                                End If
                                
                                Me.TDBGridCuentas.Columns(8).Visible = False
                    End If
    
End Sub


Private Sub CmdContabilizarPlanilla_Click()
 
  Dim Periodo As Double, NumeroPeriodo As Double, FechaIni As String, FechaFin As String, EstadoPeriodo As String, NumeroTransaccion As Double
  Dim mes As Double, Año As Double, Moneda, Resultado As Boolean
  Dim SqlString As String, FechaInicio As String
  Dim TipoFactura As String, NumeroFactura As String, CodigoProducto As String, CodigoCuentaProducto As String, CodigoCliente As String
  Dim NombreCuenta As String, CodigoCuentaCliente As String, SubTotal As Double, Descuento As Double, Iva As Double, NetoPagar As Double
  Dim FechaFactura As Date, FechaVence As Date, DescripcionMovimiento As String, TasaIva As Double
  Dim MonedaFactura As String, Reg As Double, CodigoCuentaIngresos As String, CodCuentaEfectivo As String
  Dim CostoProducto As Double, CodigoCuentaCostos As String, CodigoCuentaInventario As String, CodigoCuentaOtros As String, CodigoCuentaMetodo As String
  Dim Pagado As Double, NumeroReferencia As String, MonedaMovimiento As String, TasaMovimiento As Double
  Dim cn As New ADODB.Connection, SuReferencia As String, CodigoCuentaBanco As String
  Dim rs As New ADODB.Recordset, Registro As Double, NumeroCompra As String
  Dim cmd As New ADODB.Command, Ret1Porc As Double, Ret2Porc As Double, MontoRetencion1 As Double, MontoRetencion2 As Double, Ret3Porc As Double, Ret4Porc As Double, MontoRetencion3 As Double, MontoRetencion4 As Double
  Dim MontoBanco As Double, Directorio As String, NombreProductor As String
  Dim CtaxPagar As String, Cuenta_Banco As String, Cuenta_IR As String, Cuenta_Bolsa As String, Cuenta_Anticipo As String, Cuenta_Pulperia As String
  Dim Cuenta_Transporte As String, Cuenta_Trazabilidad As String, Cuenta_Veterinario As String, Cuenta_Inseminacion As String
  Dim Cuenta_Otras As String, NumeroNomina As String, Cuenta_Debito As String, Cuenta_Credito As String, CodigoProductor As String
  
  Dim MontoNominaPagar As Double, IR As Double, DeduccionPolicia As Double
  Dim Anticipo As Double, DeduccionTransporte As Double, Pulperia As Double, Inseminacion As Double, ProductosVeterinarios As Double, Trazabilidad As Double
  Dim OtrasDeducciones As Double, Bolsa As Double
  
  Reg = 1
  
  
  Me.CmdContabilizarPlanilla.Enabled = False
  
    If Me.OptRecepcion.Value = True Then
     TipoFactura = "Recepcion"
    ElseIf Me.OptPlanilla.Value = True Then
     TipoFactura = "Pago Proveedor"
    ElseIf Me.OptPlanillaTransportista.Value = True Then
     TipoFactura = "Pago Transportista"
    ElseIf Me.OptLiquidacion.Value = True Then
     TipoFactura = "LiquidacionLeche"
    End If


                Select Case TipoFactura
                     Case "LiquidacionLeche"
                       FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
                       SqlString = "SELECT DISTINCT LiquidacionLeche.NumeroLiquidacion, MIN(LiquidacionLeche.FechaInicio) AS FechaInicio, MAX(LiquidacionLeche.FechaFin) AS FechaFin, SUM(DetalleLiquidacionLeche.Total_Ingresos - DetalleLiquidacionLeche.Total_Deducciones) AS NetoPagar, MAX(DetalleLiquidacionLeche.Codigo_Productor) AS Codigo_Productor, MAX(Clientes.Cod_Cliente) AS Cod_Cliente, MAX(Clientes.Cod_Cuenta_Banco) AS Cod_Cuenta_Banco FROM LiquidacionLeche INNER JOIN  Clientes ON LiquidacionLeche.Cod_Proveedor = Clientes.Cod_Cliente INNER JOIN DetalleLiquidacionLeche ON LiquidacionLeche.NumeroLiquidacion = DetalleLiquidacionLeche.NumeroLiquidacion  " & _
                                    "Where (LiquidacionLeche.Contabilizado = 0) And (LiquidacionLeche.Marca = 1) GROUP BY LiquidacionLeche.NumeroLiquidacion"
                    Case "Pago Transportista"
                       FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
                       SqlString = "SELECT NominaTransportista.NumPlanilla, NominaTransportista.FechaFinal, NominaTransportista.FechaInicial, Conductor.Codigo As CodProductor, Conductor.Nombre As NombreProductor, Detalle_NominaTransportista.PrecioVenta,  Detalle_NominaTransportista.Total, Detalle_NominaTransportista.TotalIngresos  " & _
                                   ", Conductor.Cuenta_Contable AS Cod_Cuenta_Pagar, Detalle_NominaTransportista.IR, Detalle_NominaTransportista.DeduccionPolicia, Detalle_NominaTransportista.Anticipo, Detalle_NominaTransportista.DeduccionTransporte, Detalle_NominaTransportista.Pulperia, Detalle_NominaTransportista.Inseminacion, Detalle_NominaTransportista.ProductosVeterinarios, Detalle_NominaTransportista.Trazabilidad, Detalle_NominaTransportista.OtrasDeducciones, Detalle_NominaTransportista.Bolsa, Detalle_NominaTransportista.Nombres,  Conductor.Cuenta_Banco, Conductor.Cuenta_IR, Conductor.Cuenta_Bolsa, Conductor.Cuenta_Anticipo, Conductor.Cuenta_Pulperia, Conductor.Cuenta_Transporte, Conductor.Cuenta_Inseminacion, Conductor.Cuenta_Trazabilidad, Conductor.Cuenta_Veterinario, Conductor.Cuenta_Otras, Conductor.Cuenta_GastoPlanilla, " & _
                                   "Detalle_NominaTransportista.TotalIngresos - (Detalle_NominaTransportista.IR + Detalle_NominaTransportista.DeduccionPolicia + Detalle_NominaTransportista.Anticipo + Detalle_NominaTransportista.DeduccionTransporte + Detalle_NominaTransportista.Pulperia + Detalle_NominaTransportista.Inseminacion + Detalle_NominaTransportista.ProductosVeterinarios + Detalle_NominaTransportista.OtrasDeducciones + Detalle_NominaTransportista.Trazabilidad) AS NetoPagar, NominaTransportista.Marca , NominaTransportista.Contabilizado  FROM  NominaTransportista INNER JOIN  Detalle_NominaTransportista ON NominaTransportista.NumPlanilla = Detalle_NominaTransportista.NumNomina INNER JOIN Conductor ON Detalle_NominaTransportista.CodigoTransportista = Conductor.Codigo Where (NominaTransportista.Marca = 1) And (NominaTransportista.Contabilizado = 0)"

                    Case "Pago Proveedor"
                       FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
                       SqlString = "SELECT Nomina.NumPlanilla, Nomina.FechaFinal, Nomina.FechaInicial, Productor.CodProductor, Productor.NombreProductor + ' ' + Productor.ApellidoProductor As NombreProductor, Detalle_Nomina.PrecioVenta,  Detalle_Nomina.Total, Detalle_Nomina.TotalIngresos, Productor.Cod_Cuenta_Proveedor AS Cod_Cuenta_Pagar, Detalle_Nomina.IR, Detalle_Nomina.DeduccionPolicia, Detalle_Nomina.Anticipo, Detalle_Nomina.DeduccionTransporte, Detalle_Nomina.Pulperia, Detalle_Nomina.Inseminacion, Detalle_Nomina.ProductosVeterinarios, Detalle_Nomina.Trazabilidad, Detalle_Nomina.OtrasDeducciones, Detalle_Nomina.Bolsa, Detalle_Nomina.Nombres,  Productor.Cuenta_Banco, Productor.Cuenta_IR, Productor.Cuenta_Bolsa, Productor.Cuenta_Anticipo, Productor.Cuenta_Pulperia, Productor.Cuenta_Transporte, Productor.Cuenta_Inseminacion, Productor.Cuenta_Trazabilidad, Productor.Cuenta_Veterinario, Productor.Cuenta_Otras, Productor.Cuenta_GastoPlanilla, " & _
                                   "Detalle_Nomina.TotalIngresos - (Detalle_Nomina.IR + Detalle_Nomina.DeduccionPolicia + Detalle_Nomina.Anticipo + Detalle_Nomina.DeduccionTransporte + Detalle_Nomina.Pulperia + Detalle_Nomina.Inseminacion + Detalle_Nomina.ProductosVeterinarios + Detalle_Nomina.OtrasDeducciones + Detalle_Nomina.Trazabilidad) AS NetoPagar, Nomina.Marca , Nomina.Contabilizado  FROM  Nomina INNER JOIN  Detalle_Nomina ON Nomina.NumPlanilla = Detalle_Nomina.NumNomina INNER JOIN Productor ON Detalle_Nomina.CodProductor = Productor.CodProductor AND Detalle_Nomina.TipoProductor = Productor.TipoProductor  Where (Nomina.Marca = 1) And (Nomina.Contabilizado = 0)"
                
                       
                    Case "Recepcion"
                       FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
'                       SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
'                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                        SqlString = "SELECT Compras.Numero_Compra, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado, Compras.Marca , Compras.Contabilizado, Compras.Tipo_Compra, Compras.Su_Referencia, Compras.Nuestra_Referencia ,  Productor.Cod_Cuenta_Proveedor AS Cod_Cuenta_Pagar , Compras.Observaciones FROM  Compras INNER JOIN  Productor ON Compras.Cod_Proveedor = Productor.CodProductor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                    
                    Case "Mercancia Recibida"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Observaciones, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
                 
                 End Select

                 '//////////////////////////////////////////////////////////////////////////////////////////////
                 '//////////SI LA CUENTA EXISTE AGREGO LOS ENCABEZADOS///////////////////////////////////////
                 '/////////////////////////////////////////////////////////////////////////////////////////////
                 

                         mes = Month(Me.DTPicker10.Value)
                         Año = Year(Me.DTPicker10.Value)
                         FechaIni = CDate("1/" & Month(Me.DTPicker10.Value) & "/" & Year(Me.DTPicker10.Value))
                         FechaFin = DateSerial(Año, mes + 1, 1 - 1)

                 
                         Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
                         Me.AdoConsulta.Refresh
                         If Not Me.AdoConsulta.Recordset.EOF Then
                           Periodo = Me.AdoConsulta.Recordset("Periodo")
                            NumeroPeriodo = Me.AdoConsulta.Recordset("NPeriodo")
                            EstadoPeriodo = Me.AdoConsulta.Recordset("EstadoPeriodo")
                      

                  If TipoFactura = "Recepcion" Then
                        Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                        Me.AdoConsulta.Recordset.Update
                        NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                  Else
                  
                     ExisteCodigo = True
                  
                     '///////////////////////////////////////VERIFICO SI LAS CUENTAS EXISTEN EN EL SISTEMA CONTABLE
                        Directorio = App.Path + "\Cuentas.txt"
                        Open Directorio For Output As #1
                           Print #1, "Zeus Contable"
                           Print #1, "Contabilizar Nominas"
                           Print #1, ""
                           
                            Me.AdoBuscaFacturacion.RecordSource = SqlString
                            Me.AdoBuscaFacturacion.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                             If Not Me.AdoBuscaFacturacion.Recordset.EOF Then
                                Me.AdoBuscaFacturacion.Recordset.MoveFirst
                                Me.osProgress1.Max = Me.AdoBuscaFacturacion.Recordset.RecordCount
                             End If
                             Me.AdoBuscaFacturacion.Refresh
                             
                             Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                            
                              If TipoFactura = "LiquidacionLeche" Then
                                 CodigoProductor = Me.AdoBuscaFacturacion.Recordset("Cod_Cliente")
                              
                              Else
                            
                                 CodigoProductor = Me.AdoBuscaFacturacion.Recordset("CodProductor")
                            
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cod_Cuenta_Pagar")) Then CtaxPagar = Me.AdoBuscaFacturacion.Recordset("Cod_Cuenta_Pagar")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Banco")) Then Cuenta_Banco = Me.AdoBuscaFacturacion.Recordset("Cuenta_Banco")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_IR")) Then Cuenta_IR = Me.AdoBuscaFacturacion.Recordset("Cuenta_IR")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Bolsa")) Then Cuenta_Bolsa = Me.AdoBuscaFacturacion.Recordset("Cuenta_Bolsa")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Anticipo")) Then Cuenta_Anticipo = Me.AdoBuscaFacturacion.Recordset("Cuenta_Anticipo")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Pulperia")) Then Cuenta_Pulperia = Me.AdoBuscaFacturacion.Recordset("Cuenta_Pulperia")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Transporte")) Then Cuenta_Transporte = Me.AdoBuscaFacturacion.Recordset("Cuenta_Transporte")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Inseminacion")) Then Cuenta_Inseminacion = Me.AdoBuscaFacturacion.Recordset("Cuenta_Inseminacion")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Trazabilidad")) Then Cuenta_Trazabilidad = Me.AdoBuscaFacturacion.Recordset("Cuenta_Trazabilidad")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Veterinario")) Then Cuenta_Veterinario = Me.AdoBuscaFacturacion.Recordset("Cuenta_Veterinario")
                                 If Not IsNull(Me.AdoBuscaFacturacion.Recordset("Cuenta_Otras")) Then Cuenta_Otras = Me.AdoBuscaFacturacion.Recordset("Cuenta_Otras")

                           
                                If ValidarCuentas(CtaxPagar) = False Then Print #1, "CuentaXPagar " & Cod_Cuenta_Pagar & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Banco) = False Then Print #1, "CuentaBanco " & Cuenta_Banco & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_IR) = False Then Print #1, "Cuenta IR " & Cuenta_IR & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Bolsa) = False Then Print #1, "Cuenta Bolsa " & Cuenta_Bolsa & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Pulperia) = False Then Print #1, "Cuenta Fondo " & Cuenta_Pulperia & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Transporte) = False Then Print #1, "CuentaTransporte: " & Cuenta_Transporte & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Inseminacion) = False Then Print #1, "Cuenta Inseminacion " & Cuenta_Inseminacion & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Trazabilidad) = False Then Print #1, "Cuenta Trazabilidad " & Cuenta_Trazabilidad & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Veterinario) = False Then Print #1, "Cuenta Veterinario " & Cuenta_Veterinario & " Productor: " & CodigoProductor; ExisteCodigo = False
                                If ValidarCuentas(Cuenta_Otras) = False Then Print #1, "Otras Cuentas " & Cuenta_Otras & " Productor: " & CodigoProductor; ExisteCodigo = False
                                
                                 Cuenta_Debito = ""
                                 Cuenta_Credito = ""
                                '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<BUSCO LAS CONTRA CUENTAS DE SALDOS >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<>>
                                Me.AdoContraCuentaFacturacion.RecordSource = "SELECT CuentaCredito, CuentaDebito From ContraCuentaPlanillaLeche WHERE (CuentaDebito = '" & Cuenta_Banco & "')"
                                Me.AdoContraCuentaFacturacion.Refresh
                                If Not Me.AdoContraCuentaFacturacion.Recordset.EOF Then
                                
                                  Cuenta_Debito = Me.AdoContraCuentaFacturacion.Recordset("CuentaDebito")
                                  Cuenta_Credito = Me.AdoContraCuentaFacturacion.Recordset("CuentaCredito")
                                  
                                  If ValidarCuentas(Cuenta_Debito) = False Then Print #1, Cuenta_Debito & " Productor: " & CodigoProductor; ExisteCodigo = False
                                  If ValidarCuentas(Cuenta_Credito) = False Then Print #1, Cuenta_Credito & " Productor: " & CodigoProductor; ExisteCodigo = False
                                
                                Else
'                                  Print #1, "No Existen las contra Cuentas " & Cuenta_Banco & " Productor: " & CodigoProductor
'                                  ExisteCodigo = False
                                End If
                           
                             End If
                           
                               Me.AdoBuscaFacturacion.Recordset.MoveNext
                           Loop
                       Close #1
                              
                              
                         If ExisteCodigo = False Then
                           MsgBox "No existen Cuentas", vbCritical, "Sistema Contable"
                        
                           Abrir = "notepad.exe " & Directorio
                           Shell Abrir
                           Exit Sub
           
                        End If
                           
                  End If

                               
 
                             Me.AdoProcesos.RecordSource = SqlString
                             Me.AdoProcesos.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                             If Not Me.AdoProcesos.Recordset.EOF Then
                             Me.AdoProcesos.Recordset.MoveFirst
                             Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                             End If
                             Me.AdoProcesos.Refresh
                             
                             Beneficiario = ""
                             Ret1Porc = 0
                             Ret2Porc = 0
                             Ret3Porc = 0
                             Ret4Porc = 0
                             
                             
                             
                        Do While Not Me.AdoProcesos.Recordset.EOF
                           
                           Select Case TipoFactura
                           
                              Case "LiquidacionLeche"
                                FechaFactura = Me.AdoProcesos.Recordset("FechaInicio")
                                FechaVence = Me.AdoProcesos.Recordset("FechaFin")
                                NumeroNomina = Me.AdoProcesos.Recordset("NumeroLiquidacion")
                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cliente")
                                NetoPagar = Me.AdoProcesos.Recordset("NetoPagar")
                                
                                If Not IsNull(Me.AdoProcesos.Recordset("Cod_Cuenta_Banco")) Then
                                      Cuenta_Banco = Me.AdoProcesos.Recordset("Cod_Cuenta_Banco")
                                End If
                                
                                If CodigoCuentaCliente = "" Then
                                  MsgBox "No existe la cuenta del Cliente", vbCritical, "Zeus contable"
                                  Exit Sub
                                End If
                                
                                Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                                Me.AdoConsulta.Recordset.Update
                                NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                                
                                
                               If Reg = 1 Then
                                   '////////////////////////////////////////////////////////////////
                                   '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                   '///////////////////////////////////////////////////////////////
                                   MonedaNomina = "Córdobas"
                                    Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker10.Value, "yyyy-mm-dd"), "Movimiento de Nominas Liquidacion", "LiquidLeche", "Córdobas")
                                    Reg = 2
                                End If
                                
                                DescripcionMovimiento = "Pago de la Liquidacion No" & NumeroNomina & " Desde " & Me.AdoProcesos.Recordset("FechaInicio") & " Hasta " & Me.AdoProcesos.Recordset("FechaFin")
                                
                               '////////////////////////////////////CONTRACUENTA DE BANCO //////////////////////////////////////////////////////////////////
                              
                                Me.AdoContraCuentaFacturacion.RecordSource = "SELECT CuentaCredito, CuentaDebito From ContraCuentaPlanillaLeche WHERE (CuentaDebito = '" & Cuenta_Banco & "')"
                                Me.AdoContraCuentaFacturacion.Refresh
                                If Not Me.AdoContraCuentaFacturacion.Recordset.EOF Then
                                
                                  Cuenta_Debito = Me.AdoContraCuentaFacturacion.Recordset("CuentaDebito")
                                  Cuenta_Credito = Me.AdoContraCuentaFacturacion.Recordset("CuentaCredito")
                                  
                               
                                    Credito = 0
                                    If NetoPagar <> 0 Then
                                    NumeroFactura = "-"
                                    Resultado = GrabaDetalleNomina(Cuenta_Banco, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, NetoPagar, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                    End If
                                
                                    Debito = 0
                                    If NetoPagar <> 0 Then
                                    NumeroFactura = "-"
                                    Resultado = GrabaDetalleNomina(CodigoCuentaCliente, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, NetoPagar, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                    End If
                                End If


                                
                                Me.AdoConsultaFacturacion.RecordSource = "SELECT DISTINCT NumeroLiquidacion, Contabilizado From LiquidacionLeche WHERE  (NumeroLiquidacion = '" & NumeroNomina & "')"
                                Me.AdoConsultaFacturacion.Refresh
                                If Not Me.AdoConsultaFacturacion.Recordset.EOF Then
                                   Me.AdoConsultaFacturacion.Recordset("Contabilizado") = 1
                                   Me.AdoConsultaFacturacion.Recordset.Update
                                End If
                                
                                Me.PushButton3() = True
                           

                           
                              Case "Pago Transportista"
                              
                                
                              
                                MontoBanco = 0
                                FechaFactura = Me.AdoProcesos.Recordset("FechaFinal")
                                FechaVence = Me.AdoProcesos.Recordset("FechaFinal")
                                NumeroNomina = Me.AdoProcesos.Recordset("NumPlanilla")
                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                NombreProductor = Me.AdoProcesos.Recordset("NombreProductor")
                                
                                If CodigoCuentaCliente = "" Then
                                  MsgBox "No existe la cuenta del productor", vbCritical, "Zeus contable"
                                  Exit Sub
                                End If
                                
                              
                                
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")) Then
                                      CtaxPagar = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                 End If
                                   If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Banco")) Then
                                      Cuenta_Banco = Me.AdoProcesos.Recordset("Cuenta_Banco")
                                 End If
                                
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_IR")) Then
                                      Cuenta_IR = Me.AdoProcesos.Recordset("Cuenta_IR")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Bolsa")) Then
                                      Cuenta_Bolsa = Me.AdoProcesos.Recordset("Cuenta_Bolsa")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Anticipo")) Then
                                      Cuenta_Anticipo = Me.AdoProcesos.Recordset("Cuenta_Anticipo")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Pulperia")) Then
                                      Cuenta_Pulperia = Me.AdoProcesos.Recordset("Cuenta_Pulperia")
                                 End If
                                 
                                                                  
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Transporte")) Then
                                      Cuenta_Transporte = Me.AdoProcesos.Recordset("Cuenta_Transporte")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Inseminacion")) Then
                                      Cuenta_Inseminacion = Me.AdoProcesos.Recordset("Cuenta_Inseminacion")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Trazabilidad")) Then
                                      Cuenta_Trazabilidad = Me.AdoProcesos.Recordset("Cuenta_Trazabilidad")
                                 End If
                              
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Veterinario")) Then
                                      Cuenta_Veterinario = Me.AdoProcesos.Recordset("Cuenta_Veterinario")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Otras")) Then
                                      Cuenta_Otras = Me.AdoProcesos.Recordset("Cuenta_Otras")
                                End If
                                
                                
                               '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<BUSCO LAS CONTRA CUENTAS DE SALDOS >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<>>
                               Me.AdoContraCuentaFacturacion.RecordSource = "SELECT CuentaCredito, CuentaDebito From ContraCuentaPlanillaLeche WHERE (CuentaDebito = '" & Cuenta_Banco & "')"
                               Me.AdoContraCuentaFacturacion.Refresh
                               If Not Me.AdoContraCuentaFacturacion.Recordset.EOF Then
                                 Cuenta_Debito = Me.AdoContraCuentaFacturacion.Recordset("CuentaDebito")
                                 Cuenta_Credito = Me.AdoContraCuentaFacturacion.Recordset("CuentaCredito")
                               End If
                                
                                
                              Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                              Me.AdoConsulta.Recordset.Update
                              NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                                
                                
                               If Reg = 1 Then
                                   '////////////////////////////////////////////////////////////////
                                   '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                   '///////////////////////////////////////////////////////////////
                                   MonedaNomina = "Córdobas"
                                    Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker10.Value, "yyyy-mm-dd"), "Movimiento de Nominas Transportista", "CHEQUE", "Córdobas")
                                    Reg = 2
                                End If
                                 
                                 
                                 
'                                Me.AdoBuscaNomina.RecordSource = "SELECT  * FROM  Nomina INNER JOIN  TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina Where (Nomina.NumNomina = " & NumNomina & ")"
'                                Me.AdoBuscaNomina.Refresh
'                                If Not Me.AdoBuscaNomina.Recordset.EOF Then
'                                  DescripcionMovimiento = "Registrando Nomina " & NumeroFactura
'                                  DescripcionMovimiento = "Registrando Nominas No " & NumNomina
'                                End If
'
                                 
                                MontoNominaPagar = Format(Me.AdoProcesos.Recordset("TotalIngresos"), "##,##0.00")
                                NetoPagar = Format(Me.AdoProcesos.Recordset("NetoPagar"), "##,##0.00")
                                IR = Format(Me.AdoProcesos.Recordset("IR"), "##,##0.00")
                                DeduccionPolicia = Format(Me.AdoProcesos.Recordset("DeduccionPolicia"), "##,##0.00")
                                Anticipo = Format(Me.AdoProcesos.Recordset("Anticipo"), "##,##0.00")
                                DeduccionTransporte = Format(Me.AdoProcesos.Recordset("DeduccionTransporte"), "##,##0.00")
                                Pulperia = Format(Me.AdoProcesos.Recordset("Pulperia"), "##,##0.00")
                                Inseminacion = Format(Me.AdoProcesos.Recordset("Inseminacion"), "##,##0.00")
                                ProductosVeterinarios = Format(Me.AdoProcesos.Recordset("ProductosVeterinarios"), "##,##0.00")
                                Trazabilidad = Format(Me.AdoProcesos.Recordset("Trazabilidad"), "##,##0.00")
                                OtrasDeducciones = Format(Me.AdoProcesos.Recordset("OtrasDeducciones"), "##,##0.00")
                                Bolsa = Format(Me.AdoProcesos.Recordset("Bolsa"), "##,##0.00")

                                NetoPagar = Format(MontoNominaPagar - IR - DeduccionPolicia - Anticipo - DeduccionTransporte - Pulperia - Inseminacion - ProductosVeterinarios - Trazabilidad - OtrasDeducciones - Bolsa, "##,##0.00")
                                
                                NombreProductor = Me.AdoProcesos.Recordset("NombreProductor")
                                DescripcionMovimiento = "Pago de Planilla Transportista No" & Me.AdoProcesos.Recordset("NumPlanilla") & " Desde " & Me.AdoProcesos.Recordset("FechaInicial") & " Hasta " & Me.AdoProcesos.Recordset("FechaFinal")
                                
                                '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                '//////////////////////////CREO LAS CONTRA CUENTAS PARA FONDOS DE PLANILLA //////////////////////////////
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                
                                '////////////////////////////////////CONTRACUENTA DE BANCO //////////////////////////////////////////////////////////////////
                                
                                                               '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<BUSCO LAS CONTRA CUENTAS DE SALDOS >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<>>
                               Me.AdoContraCuentaFacturacion.RecordSource = "SELECT CuentaCredito, CuentaDebito From ContraCuentaPlanillaLeche WHERE (CuentaDebito = '" & Cuenta_Banco & "')"
                               Me.AdoContraCuentaFacturacion.Refresh
                               If Not Me.AdoContraCuentaFacturacion.Recordset.EOF Then
                                 Cuenta_Debito = Me.AdoContraCuentaFacturacion.Recordset("CuentaDebito")
                                 Cuenta_Credito = Me.AdoContraCuentaFacturacion.Recordset("CuentaCredito")
                                
                                Credito = 0
                                If NetoPagar <> 0 Then
                                NumeroFactura = "-"
                                Resultado = GrabaDetalleNomina(Cuenta_Debito, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, NetoPagar, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                End If
                                
                                Debito = 0
                                If NetoPagar <> 0 Then
                                NumeroFactura = "-"
                                Resultado = GrabaDetalleNomina(Cuenta_Credito, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, NetoPagar, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                End If
                               
                               
                              End If

                                
                                
                                
                                '///////////////////////////CREO LOS REGISTROS CONTABLES X PAGAR/////////////////////////////////
                               
                                Credito = 0
                                NumeroFactura = "-"
                                If MontoNominaPagar <> 0 Then
                                Resultado = GrabaDetalleNomina(CtaxPagar, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoNominaPagar, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                '////////////////////////////BUSCO LAS CUENTAS DE INCENTIVOS ///////////////////////////////////////
                                End If
                             
                                '////////////////////////////////////CUENTA DE BANCO //////////////////////////////////////////////////////////////////
                                Debito = 0
                                
                                If NetoPagar <> 0 Then
                                NumeroFactura = "#######"
                                Resultado = GrabaDetalleNomina(Cuenta_Banco, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, NetoPagar, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If IR <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_IR, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, IR, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Bolsa <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Bolsa, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Bolsa, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Anticipo <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Anticipo, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Anticipo, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Pulperia <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Pulperia, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Pulperia, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If DeduccionTransporte <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Transporte, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, DeduccionTransporte, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Inseminacion <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Inseminacion, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Inseminacion, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Trazabilidad <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Trazabilidad, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Trazabilidad, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If ProductosVeterinarios <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Veterinario, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, ProductosVeterinarios, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If OtrasDeducciones <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Otras, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, OtrasDeducciones, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Reg = 1
                                
                                
                                Me.AdoConsultaFacturacion.RecordSource = "SELECT  NumPlanilla, Contabilizado From NominaTransportista WHERE  (NumPlanilla = '" & NumeroNomina & "')"
                                Me.AdoConsultaFacturacion.Refresh
                                If Not Me.AdoConsultaFacturacion.Recordset.EOF Then
                                   Me.AdoConsultaFacturacion.Recordset("Contabilizado") = 1
                                   Me.AdoConsultaFacturacion.Recordset.Update
                                End If
                                
                                Me.PushButton3() = True
                           
                           
                         Case "Pago Proveedor"
                              
                                
                              
                                MontoBanco = 0
                                FechaFactura = Me.AdoProcesos.Recordset("FechaFinal")
                                FechaVence = Me.AdoProcesos.Recordset("FechaFinal")
                                NumeroNomina = Me.AdoProcesos.Recordset("NumPlanilla")
                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                NombreProductor = Me.AdoProcesos.Recordset("NombreProductor")
                                
                                If CodigoCuentaCliente = "" Then
                                  MsgBox "No existe la cuenta del productor", vbCritical, "Zeus contable"
                                  Exit Sub
                                End If
                                
                              
                                
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")) Then
                                      CtaxPagar = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                 End If
                                   If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Banco")) Then
                                      Cuenta_Banco = Me.AdoProcesos.Recordset("Cuenta_Banco")
                                 End If
                                
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_IR")) Then
                                      Cuenta_IR = Me.AdoProcesos.Recordset("Cuenta_IR")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Bolsa")) Then
                                      Cuenta_Bolsa = Me.AdoProcesos.Recordset("Cuenta_Bolsa")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Anticipo")) Then
                                      Cuenta_Anticipo = Me.AdoProcesos.Recordset("Cuenta_Anticipo")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Pulperia")) Then
                                      Cuenta_Pulperia = Me.AdoProcesos.Recordset("Cuenta_Pulperia")
                                 End If
                                 
                                                                  
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Transporte")) Then
                                      Cuenta_Transporte = Me.AdoProcesos.Recordset("Cuenta_Transporte")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Inseminacion")) Then
                                      Cuenta_Inseminacion = Me.AdoProcesos.Recordset("Cuenta_Inseminacion")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Trazabilidad")) Then
                                      Cuenta_Trazabilidad = Me.AdoProcesos.Recordset("Cuenta_Trazabilidad")
                                 End If
                              
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Veterinario")) Then
                                      Cuenta_Veterinario = Me.AdoProcesos.Recordset("Cuenta_Veterinario")
                                 End If
                                 
                                 If Not IsNull(Me.AdoProcesos.Recordset("Cuenta_Otras")) Then
                                      Cuenta_Otras = Me.AdoProcesos.Recordset("Cuenta_Otras")
                                End If
                                
                                
                               '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<BUSCO LAS CONTRA CUENTAS DE SALDOS >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<>>
                               Me.AdoContraCuentaFacturacion.RecordSource = "SELECT CuentaCredito, CuentaDebito From ContraCuentaPlanillaLeche WHERE (CuentaDebito = '" & Cuenta_Banco & "')"
                               Me.AdoContraCuentaFacturacion.Refresh
                               If Not Me.AdoContraCuentaFacturacion.Recordset.EOF Then
                                 Cuenta_Debito = Me.AdoContraCuentaFacturacion.Recordset("CuentaDebito")
                                 Cuenta_Credito = Me.AdoContraCuentaFacturacion.Recordset("CuentaCredito")
                               End If
                                
                                
                              Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                              Me.AdoConsulta.Recordset.Update
                              NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                                
                                
                               If Reg = 1 Then
                                   '////////////////////////////////////////////////////////////////
                                   '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                   '///////////////////////////////////////////////////////////////
                                   MonedaNomina = "Córdobas"
                                    Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker10.Value, "yyyy-mm-dd"), "Movimiento de Nominas", "CHEQUE", "Córdobas")
                                    Reg = 2
                                End If
                                 
                                 
                                 
'                                Me.AdoBuscaNomina.RecordSource = "SELECT  * FROM  Nomina INNER JOIN  TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina Where (Nomina.NumNomina = " & NumNomina & ")"
'                                Me.AdoBuscaNomina.Refresh
'                                If Not Me.AdoBuscaNomina.Recordset.EOF Then
'                                  DescripcionMovimiento = "Registrando Nomina " & NumeroFactura
'                                  DescripcionMovimiento = "Registrando Nominas No " & NumNomina
'                                End If
'

                                
                                MontoNominaPagar = Format(Me.AdoProcesos.Recordset("TotalIngresos"), "##,##0.00")
                                NetoPagar = Format(Me.AdoProcesos.Recordset("NetoPagar"), "##,##0.00")
                                IR = Format(Me.AdoProcesos.Recordset("IR"), "##,##0.00")
                                DeduccionPolicia = Format(Me.AdoProcesos.Recordset("DeduccionPolicia"), "##,##0.00")
                                Anticipo = Format(Me.AdoProcesos.Recordset("Anticipo"), "##,##0.00")
                                DeduccionTransporte = Format(Me.AdoProcesos.Recordset("DeduccionTransporte"), "##,##0.00")
                                Pulperia = Format(Me.AdoProcesos.Recordset("Pulperia"), "##,##0.00")
                                Inseminacion = Format(Me.AdoProcesos.Recordset("Inseminacion"), "##,##0.00")
                                ProductosVeterinarios = Format(Me.AdoProcesos.Recordset("ProductosVeterinarios"), "##,##0.00")
                                Trazabilidad = Format(Me.AdoProcesos.Recordset("Trazabilidad"), "##,##0.00")
                                OtrasDeducciones = Format(Me.AdoProcesos.Recordset("OtrasDeducciones"), "##,##0.00")
                                Bolsa = Format(Me.AdoProcesos.Recordset("Bolsa"), "##,##0.00")

                                NetoPagar = Format(MontoNominaPagar - IR - DeduccionPolicia - Anticipo - DeduccionTransporte - Pulperia - Inseminacion - ProductosVeterinarios - Trazabilidad - OtrasDeducciones - Bolsa, "##,##0.00")
                                
                                NombreProductor = Me.AdoProcesos.Recordset("NombreProductor")
                                DescripcionMovimiento = "Pago de Planilla por acopio de Leche No" & Me.AdoProcesos.Recordset("NumPlanilla") & " Desde " & Me.AdoProcesos.Recordset("FechaInicial") & " Hasta " & Me.AdoProcesos.Recordset("FechaFinal")
                                
                                '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                '//////////////////////////CREO LAS CONTRA CUENTAS PARA FONDOS DE PLANILLA //////////////////////////////
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                
                                '////////////////////////////////////CONTRACUENTA DE BANCO //////////////////////////////////////////////////////////////////
                                                               '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<BUSCO LAS CONTRA CUENTAS DE SALDOS >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<>>
                               Me.AdoContraCuentaFacturacion.RecordSource = "SELECT CuentaCredito, CuentaDebito From ContraCuentaPlanillaLeche WHERE (CuentaDebito = '" & Cuenta_Banco & "')"
                               Me.AdoContraCuentaFacturacion.Refresh
                               If Not Me.AdoContraCuentaFacturacion.Recordset.EOF Then
                                 Cuenta_Debito = Me.AdoContraCuentaFacturacion.Recordset("CuentaDebito")
                                 Cuenta_Credito = Me.AdoContraCuentaFacturacion.Recordset("CuentaCredito")
                               
                                    Credito = 0
                                    If NetoPagar <> 0 Then
                                    NumeroFactura = "-"
                                    Resultado = GrabaDetalleNomina(Cuenta_Debito, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, NetoPagar, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                    End If
                                    
                                    Debito = 0
                                    If NetoPagar <> 0 Then
                                    NumeroFactura = "-"
                                    Resultado = GrabaDetalleNomina(Cuenta_Credito, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, NetoPagar, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                    End If
                               
                               End If
                               
                               

                                
                                
                                
                                '///////////////////////////CREO LOS REGISTROS CONTABLES X PAGAR/////////////////////////////////
                               
                                Credito = 0
                                NumeroFactura = "-"
                                If MontoNominaPagar <> 0 Then
                                Resultado = GrabaDetalleNomina(CtaxPagar, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, MontoNominaPagar, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                '////////////////////////////BUSCO LAS CUENTAS DE INCENTIVOS ///////////////////////////////////////
                                End If
                             
                                '////////////////////////////////////CUENTA DE BANCO //////////////////////////////////////////////////////////////////
                                Debito = 0
                                
                                If NetoPagar <> 0 Then
                                NumeroFactura = "#######"
                                Resultado = GrabaDetalleNomina(Cuenta_Banco, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, NetoPagar, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, NombreProductor)
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If IR <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_IR, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, IR, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Bolsa <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Bolsa, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Bolsa, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Anticipo <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Anticipo, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Anticipo, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Pulperia <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Pulperia, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Pulperia, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If DeduccionTransporte <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Transporte, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, DeduccionTransporte, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Inseminacion <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Inseminacion, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Inseminacion, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If Trazabilidad <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Trazabilidad, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Trazabilidad, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If ProductosVeterinarios <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Veterinario, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, ProductosVeterinarios, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Debito = 0
                                NumeroFactura = "-"
                                If OtrasDeducciones <> 0 Then
                                Resultado = GrabaDetalleNomina(Cuenta_Otras, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, OtrasDeducciones, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "-")
                                End If
                                
                                Reg = 1
                                 
          
                              Case Else
                             
                             
                             
                                                FechaFactura = Me.AdoProcesos.Recordset("Fecha_Compra")
                                                If Not IsNull(Me.AdoProcesos.Recordset("Fecha_Vencimiento")) Then
                                                  FechaVence = Me.AdoProcesos.Recordset("Fecha_Vencimiento")
                                                End If
                                                NumeroFactura = Me.AdoProcesos.Recordset("Numero_Compra")
                                                MonedaFactura = Me.AdoProcesos.Recordset("MonedaCompra")
                                                If Not IsNull(Me.AdoProcesos.Recordset("Su_Referencia")) Then
                                                   NumeroReferencia = "Compra No:" & NumeroFactura & " " & "Referencia: " & Me.AdoProcesos.Recordset("Su_Referencia")
                                                Else
                                                   NumeroReferencia = "Compra No:" & NumeroFactura
                                                End If
                                                
                                                If Not IsNull(Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")) Then
                                                  CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                                Else
                                                  CodigoCuentaCliente = "2121"
                                                End If
                                                
                                                Select Case TipoFactura
                                                  Case "Cuenta"
'                                                     codigocuentacliente =
                                                
                                                End Select
                                                
                                                
                                                    SubTotal = 0
                                                    Descuento = 0
                                                    Iva = 0
                                                    NetoPagar = 0
                                                    TasaCambio = 1
                                                    TasaMovimiento = 1
                                                    
                                                    
                                                    SqlString = "SELECT * From Detalle_Compras WHERE (Numero_Compra = '" & NumeroFactura & "') AND (Tipo_Compra = '" & TipoFactura & "')"
                                                    Me.AdoBuscaFacturacion.RecordSource = SqlString
                                                    Me.AdoBuscaFacturacion.Refresh
                                                    
                                                    If MonedaFactura = "Dolares" Then
                                                         TasaCambio = Format(Val(Me.AdoBuscaFacturacion.Recordset("TasaCambio")), "##,##0.000000")
                                                    Else
                                                         TasaCambio = 1
                                                    End If
                                                
                
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////////////////BUSCO LOS DETALLES DE LACOMPRA ////////////////////////////////////////////////////
                                                '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                Me.AdoConsultaFactura.RecordSource = "SELECT SUM(Detalle_Compras.Cantidad) AS Cantidad, SUM(Detalle_Compras.Precio_Unitario) AS Precio_Unitario, SUM(Detalle_Compras.Descuento) AS Descuento, SUM(Detalle_Compras.Precio_Neto) AS Precio_Neto, SUM(Detalle_Compras.Importe) AS Importe FROM  Compras INNER JOIN Detalle_Compras ON Compras.Numero_Compra = Detalle_Compras.Numero_Compra AND Compras.Fecha_Compra = Detalle_Compras.Fecha_Compra And Compras.Tipo_Compra = Detalle_Compras.Tipo_Compra  " & _
                                                                                      "WHERE (Compras.Numero_Compra = '" & NumeroFactura & "') AND (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Tipo_Compra = '" & TipoFactura & "')"
                                                Me.AdoConsultaFactura.Refresh
                                                If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                                  If Not IsNull(Me.AdoConsultaFactura.Recordset("Importe")) Then
                                                    If Format(CDbl(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00") = Format(CDbl(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00") Then
                                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00")
                                                        SubTotal = Format(SubTotal * TasaCambio, "##,##0.00")
                                                        If Not IsNull(Me.AdoProcesos.Recordset("Descuento")) Then
                                                         If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                                          Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                                          Descuento = Format(Descuento * TasaCambio, "##,##0.00")
                                                         End If
                                                        End If
                                                        Iva = Format(Val(Me.AdoProcesos.Recordset("IVA")), "##,##0.00")
                                                        Iva = Format(Iva * TasaCambio, "##,##0.00")
                                                        NetoPagar = Format(Val(Me.AdoProcesos.Recordset("NetoPagar")), "##,##0.00")
                                                        NetoPagar = Format(NetoPagar * TasaCambio, "##,##0.00")
                                                        Pagado = Format(Val(Me.AdoProcesos.Recordset("Pagado")), "##,##0.00")
                                                        Pagado = Format(Pagado * TasaCambio, "##,##0.00")
                                                    Else
                                                        SubTotal = Format(Val(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00")
                                                        SubTotal = Format(SubTotal * TasaCambio, "##,##0.00")
                                                        Descuento = 0
                                                        Iva = 0
                                                        NetoPagar = 0
                                                        Pagado = 0
                                                    End If
                                               
                                                
                                                
                                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                       MonedaFactura = Me.AdoProcesos.Recordset("MonedaCompra")
                                                       
                
                
                    
                
                                                
                                                      If Reg = 1 Then
                                                         '////////////////////////////////////////////////////////////////
                                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                                         '/////////////////////////////////////////////////////////////////
                                                         MonedaMovimiento = "Córdobas"
                                                         Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker10.Value, "yyyy-mm-dd"), "Movimiento de Compras", "Recepcion", MonedaMovimiento)
                                                         Reg = 2
                                                      End If
                                                           
                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                        '////////////AGREGO LA CUENTA DEL PROVEEDOR//////////////////////////////////////
                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                           
                                                           Credito = 0
                                                           Debito = 0
                                                          If NetoPagar = 0 Then
                                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                                                                                      
                                                            Me.AdoConsultaFactura.RecordSource = "SELECT Detalle_MetodoCompras.Numero_Compra, Detalle_MetodoCompras.Fecha_Compra, Detalle_MetodoCompras.Tipo_Compra, Detalle_MetodoCompras.NombrePago, Detalle_MetodoCompras.Monto, Detalle_MetodoCompras.NumeroTarjeta, Detalle_MetodoCompras.FechaVence , MetodoPago.TipoPago, MetodoPago.Cod_Cuenta, MetodoPago.Moneda FROM Detalle_MetodoCompras INNER JOIN MetodoPago ON Detalle_MetodoCompras.NombrePago = MetodoPago.NombrePago  " & _
                                                                                                 "WHERE (Detalle_MetodoCompras.Numero_Compra = '" & NumeroFactura & "')"
                                                            Me.AdoConsultaFactura.Refresh
                                                            If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                                                CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                                
                                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                                NetoPagar = SubTotal + Iva - Descuento
                                                                Select Case TipoFactura
                                                                  Case "Transferencia Recibida"
                                                                       DescripcionMovimiento = "Registro de Transferencia Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                                                                  
                                                                  
                '                                                       DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Recepcion"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                                                                  
                                                                  
                '                                                       DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de Devolucion Numero " & NumeroFactura
                                                                       Debito = Format(NetoPagar, "##,##0.00")
                                                                       Credito = 0
                                                                End Select
                                                            Else
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                                NetoPagar = SubTotal + Iva - Descuento
                                                                Debito = Format(NetoPagar, "##,##0.00")
                                                            End If
                                                           ElseIf NetoPagar < 0 Then
                                                              '/////////////////AGREGO LA DIFERENICIA A OTROS INGRESOS /
                                                               Me.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (TipoCuenta = 'Capital')"
                                                               Me.AdoConsulta.Refresh
                                                               If Not Me.AdoConsulta.Recordset.EOF Then
                                                                CodigoCuentaOtros = Me.AdoConsulta.Recordset("CodCuentas")
                                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaOtros)
                                                               End If
                                                               
                                                               
                                                               Select Case TipoFactura
                                                                  Case "Transferencia Recibida"
                                                                       DescripcionMovimiento = "Registro de AJUSTE Transferencia Recibida Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Recepcion"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Devolucion Compra Numero " & NumeroFactura
                                                                       Debito = Format(NetoPagar, "##,##0.00")
                                                                       Credito = 0
                                                                End Select
                                                               
                                                               
                                                               Resultado = GrabaDetalleFactura(CodigoCuentaOtros, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, 0, Abs(NetoPagar), "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                               
                                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                                            Me.AdoConsultaFactura.RecordSource = "SELECT  * FROM Detalle_MetodoFacturas INNER JOIN MetodoPago ON Detalle_MetodoFacturas.NombrePago = MetodoPago.NombrePago  " & _
                                                                                                 "WHERE (Detalle_MetodoFacturas.Numero_Factura = '" & NumeroFactura & "')"
                                                            Me.AdoConsultaFactura.Refresh
                '                                              CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                              NetoPagar = SubTotal + Iva - Descuento + Abs(NetoPagar)
                                                              
                                                              Debito = Format(NetoPagar, "##,##0.00")
                                                              DescripcionMovimiento = "Registro de Facturacion Factura Numero " & NumeroFactura
                                                              DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                              
                                                           ElseIf NetoPagar > 0 Then
                                                           
                                              
                                                            DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                           
                                                           Select Case TipoFactura
 
                                                             Case "Recepcion"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Compra Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = 0
                                                                   Credito = Format(NetoPagar, "##,##0.00")
                                                                   
                                                                  DescripcionMovimiento = DescripcionMovimiento & "," & Me.AdoProcesos.Recordset("Observaciones")
                                                                  '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
 
                                                             
                                                             Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Devolucion Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = Format(NetoPagar, "##,##0.00")
                                                                   Credito = 0
                                                        '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             End Select
                                                                
                
                                                           End If
                 
'++++++++++++++++++++++++++++++++++++++++++ DETALLE DE MOVIMIENTOS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                                    '//////////////////////////////////////////////////////////////////////////////////////////
                                                    '/////////////////////CARGO LOS DETALLE DE LAS FACTURAS////////////////////////////////////
                                                    '//////////////////////////////////////////////////////////////////////////////////////////
                                                     
                                                     If TipoFactura = "Cuenta" Then
                                                        SqlString = "SELECT  Numero_Compra, Fecha_Compra, Tipo_Compra, Cod_Producto, Cantidad, Precio_Unitario, Descuento, Precio_Neto, Importe, TasaCambio From Detalle_Compras WHERE (Numero_Compra = '" & NumeroFactura & "') AND (Tipo_Compra LIKE '%" & TipoFactura & "%')"
                                                     
                                                     Else
                                                         SqlString = "SELECT  Detalle_Compras.Numero_Compra, Detalle_Compras.Fecha_Compra, Detalle_Compras.Tipo_Compra, Detalle_Compras.Cod_Producto, Detalle_Compras.Cantidad, Detalle_Compras.Precio_Unitario, Detalle_Compras.Descuento, Detalle_Compras.Precio_Neto, Detalle_Compras.Importe, Productos.Descripcion_Producto, Productos.Cod_Cuenta_Inventario, Productos.Cod_Cuenta_Costo, Productos.Cod_Cuenta_Ventas, Productos.Cod_Cuenta_GastoAjuste , Productos.Cod_Cuenta_IngresoAjuste, Detalle_Compras.TasaCambio, Productos.Costo_Promedio, Productos.Costo_Promedio_Dolar  FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos  " & _
                                                                     "WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "') AND (Detalle_Compras.Tipo_Compra Like '%" & TipoFactura & "%')"
                                                     End If
                                                     Me.AdoProcesosFacturacion.RecordSource = SqlString
                                                     Me.AdoProcesosFacturacion.Refresh
                                                     Do While Not Me.AdoProcesosFacturacion.Recordset.EOF
                                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                If TipoFactura = "Cuenta" Then
                                                                  CodigoCuentaProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                Else
                                                                  CodigoCuentaProducto = BuscaCodigoProducto(CodigoProducto)
                                                                End If
                                                                If MonedaFactura = "Dolares" Then
                '                                                 TasaCambio = BuscaTasaCambio(Fecha)
                                                                   TasaCambio = Format(Val(Me.AdoProcesosFacturacion.Recordset("TasaCambio")), "##,##0.0000")
                                                                Else
                                                                    TasaCambio = 1
                                                               End If
                                                                
                                                                
                                                                If MonedaFactura = "Dolares" Then
                                                                  CostoProducto = Me.AdoProcesosFacturacion.Recordset("Importe")
                '                                                   CostoProducto = Format(Me.AdoProcesosFacturacion.Recordset("Importe") * TasaCambio, "##,##0.00")
                                                                Else
                                                                   CostoProducto = Format(Me.AdoProcesosFacturacion.Recordset("Importe"), "##,##0.00")
                                                                End If
                                                            
                                                               '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                               '///////////////////////////CALCULO EL IVA DE CADA PRODUCTO//////////////////////////////////////////////
                                                               '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                
                                                                If Iva <> 0 Then
                                                                 TasaIva = BuscaTasaIva(CodigoProducto)
                                                                 DescripcionCuenta = BuscaCuenta(CodigoCuentaIva)
                                                                 
                                                                 QUIEN = ""
                                                                 
                                                                 If DescripcionCuenta <> "Nulo" Then
                                                                  
                                                                  Select Case TipoFactura
 
                                                                    Case "Recepcion"
                                                                         DescripcionMovimiento = "IVA Ventas Compra No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                         Credito = 0
                                                                         Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                            If Debito <> 0 Then
                                                                                Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                            End If
                                                                            

                                                                    End Select
                
                
                                                                   End If
                                                                 End If
                                                                
                                                              QUIEN = ""
                                                                
                                                                 
                                                                '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '///////////////////////////////////////BUSCO LA CUENTA DE INVENTARIO///////////////////////////////////////////
                                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                If TipoFactura <> "Cuenta" Then
                                                                    If Not IsNull(Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")) Then
                                                                      CodigoCuentaInventario = Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")
                                                                    End If
                                                                Else
                                                                    CodigoCuentaInventario = CodigoCuentaProducto
                                                                End If
                                                                
                                                                 
                                                                 DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
                                                                   Select Case TipoFactura
                                                                        
                                                                    Case "Recepcion"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                          Credito = 0
                                                                         'If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         'End If
                                                                    End Select
                            
                            
                                                        Me.AdoProcesosFacturacion.Recordset.MoveNext
                                                     Loop
                                                     
                                                     
                                                                 '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '///////////////////////////////////////BUSCO LA CUENTA DE INVENTARIO///////////////////////////////////////////
                                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                If Pagado <> 0 Then
                                                                  SqlString = "SELECT Detalle_MetodoCompras.Numero_Compra, Detalle_MetodoCompras.Fecha_Compra, Detalle_MetodoCompras.NombrePago, Detalle_MetodoCompras.Monto,Detalle_MetodoCompras.NumeroTarjeta , MetodoPago.Cod_Cuenta, MetodoPago.Moneda FROM  Detalle_MetodoCompras INNER JOIN MetodoPago ON Detalle_MetodoCompras.NombrePago = MetodoPago.NombrePago " & _
                                                                              "WHERE (Detalle_MetodoCompras.Numero_Compra = '" & NumeroFactura & "') "
                                                                  Me.AdoConsultaFacturacion.RecordSource = SqlString
                                                                  Me.AdoConsultaFacturacion.Refresh
                                                                  Do While Not Me.AdoConsultaFacturacion.Recordset.EOF
                                                                     CodigoCuentaMetodo = Me.AdoConsultaFacturacion.Recordset("Cod_Cuenta")
                                                                     
                                                                     DescripcionCuenta = BuscaCuenta(CodigoCuentaMetodo)
                                                                    Select Case TipoFactura
                                                                        Case "Recepcion"
                                                                         DescripcionMovimiento = "PAGO DE COMPRA" & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                         Debito = 0
                                                                         Credito = Me.AdoConsultaFacturacion.Recordset("Monto")
                                                                             If DescripcionCuenta <> "Nulo" Then
                                                                              Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                             End If

                                                                    End Select
                                                                    
                                                                    
                                                                
                
                                                                     If Debito <> 0 Then
                                                                         Resultado = GrabaDetalleFactura(CodigoCuentaMetodo, Me.DTPicker10.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "VTAS", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                     End If
                                                                 
                                                                    Me.AdoConsultaFacturacion.Recordset.MoveNext
                                                                  Loop
                                                                
                                                                End If
                                                                
                                                                
                                                                
                                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                rs.Open "UPDATE Compras SET Contabilizado = 1 ,Activo = 0  WHERE (Numero_Compra = '" & NumeroFactura & "') ", ConexionFacturacion
                                                     
                                                     
                                                     
                                                  End If
                                                
                                                End If
                
                                End Select
                                                                                       
                                Me.osProgress1.Value = Me.osProgress1.Value + 1
                                Me.AdoProcesos.Recordset.MoveNext
                            Loop
                             
                             If TipoFactura = "Pago Proveedor" Then
                             

                                If NumeroNomina <> "" Then
                                
                                      '///////////////////////////////////////////ACTUALIZO LA NOMINA ////////////////////////////////////////////////////////////
                                   Me.AdoBuscaFacturacion.RecordSource = "SELECT  * From Nomina WHERE (NumPlanilla = " & NumeroNomina & ")"
                                   Me.AdoBuscaFacturacion.Refresh
                                   If Not Me.AdoBuscaFacturacion.Recordset.EOF Then
                                      Me.AdoBuscaFacturacion.Recordset("Contabilizado") = True
                                      Me.AdoBuscaFacturacion.Recordset.Update
                                      PushButton3_Click
                                   End If
                                   
                               End If
                             
                             End If
'

                         End If

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
With Me.AdoDatosEmpresa
   .ConnectionString = Conexion
   .RecordSource = "SELECT  * From DatosEmpresa"
   .Refresh
End With

If Not IsNull(Me.AdoDatosEmpresa.Recordset("ConexionFacturacion")) Then
  ConexionFacturacion = Me.AdoDatosEmpresa.Recordset("ConexionFacturacion")
Else
  MsgBox "No Existe conexion con la Facturacion", vbCritical, "Zeus Contabilidad"
  Unload Me
End If

With Me.AdoNota
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoFacturacion
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoCompras
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoProcesosFacturacion
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoConsultaFacturacion
   .ConnectionString = ConexionFacturacion
End With


With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.AdoConsultaFactura
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoBuscaFacturacion
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoProcesos
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoDetalleFactura
   .ConnectionString = ConexionFacturacion
End With


With Me.AdoRecepcion
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoNominas
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoBuscaNomina
   .ConnectionString = ConexionFacturacion
End With

With Me.AdoContraCuentaFacturacion
   .ConnectionString = Conexion
End With

'SSTab1.TabVisible(0) = False
'SSTab1.TabVisible(2) = False

Me.DTPicker1.Value = Now
Me.DTPicker2.Value = Now
Me.DTPicker3.Value = Now
Me.DTPicker4.Value = Now
Me.DTPicker5.Value = Now
Me.DTPicker6.Value = Now
Me.DTPicker7.Value = Now
Me.DTPicker8.Value = Now
Me.DTPicker9.Value = Now
Me.DTPicker10.Value = Now
Me.DTPicker11.Value = Now
Me.DTPicker12.Value = Now

End Sub

Private Sub OptCompras_Click()
'  If Me.RadioButton1.Value = True Then
'    Me.CmdRecepcion.Visible = True
'    Me.CmdContabilizarCompras.Visible = False
'  Else
'    Me.CmdContabilizarCompras.Visible = True
'    Me.CmdRecepcion.Visible = False
'  End If
End Sub

Private Sub OptDevolucion_Click()
  If Me.OptDevolucion.Value = True Then
   Me.ChkDescripcion.Visible = True
  End If
End Sub

Private Sub OptFacturacion_Click()
  If Me.OptFacturacion.Value = True Then
        Me.ChkDescripcion.Visible = True
  End If
End Sub

Private Sub OptRecibos_Click()
  If Me.OptRecibos.Value = True Then
    Me.ChkDescripcion.Visible = False
  End If
End Sub

Private Sub OptSalidaBodega_Click()
  If Me.OptSalidaBodega.Value = True Then
    Me.ChkDescripcion.Visible = True
    
  End If
End Sub

Private Sub PushButton1_Click()

  Dim Periodo As Double, NumeroPeriodo As Double, FechaIni As String, FechaFin As String, EstadoPeriodo As String, NumeroTransaccion As Double
  Dim mes As Double, Año As Double, Moneda, Resultado As Boolean
  Dim SqlString As String, FechaInicio As String
  Dim TipoFactura As String, NumeroFactura As String, CodigoProducto As String, CodigoCuentaProducto As String, CodigoCliente As String
  Dim NombreCuenta As String, CodigoCuentaCliente As String, SubTotal As Double, Descuento As Double, Iva As Double, NetoPagar As Double
  Dim FechaFactura As Date, FechaVence As Date, DescripcionMovimiento As String, TasaIva As Double
  Dim MonedaFactura As String, Reg As Double, CodigoCuentaIngresos As String, CodCuentaEfectivo As String
  Dim CostoProducto As Double, CodigoCuentaCostos As String, CodigoCuentaInventario As String, CodigoCuentaOtros As String, CodigoCuentaMetodo As String
  Dim Pagado As Double, NumeroReferencia As String, MonedaMovimiento As String, TasaMovimiento As Double
  Dim cn As New ADODB.Connection, SuReferencia As String, CodigoCuentaBanco As String
  Dim rs As New ADODB.Recordset, Registro As Double, NumeroCompra As String
  Dim cmd As New ADODB.Command, Ret1Porc As Double, Ret2Porc As Double, MontoRetencion1 As Double, MontoRetencion2 As Double, Ret3Porc As Double, Ret4Porc As Double, MontoRetencion3 As Double, MontoRetencion4 As Double
  Dim MontoBanco As Double



    If Me.OptRecepcion.Value = True Then
     TipoFactura = "Recepcion"
    ElseIf Me.OptPlanilla.Value = True Then
     TipoFactura = "Planilla"
    End If
    
    Select Case TipoFactura
    
                     Case "Recepcion"
                       FechaInicio = Format(Me.DTPicker3.Value, "yyyy-mm-dd")
                       FechaFin = Format(Me.DTPicker4.Value, "yyyy-mm-dd")
                       SqlString = "SELECT  Proveedor.Cod_Cuenta_Proveedor, Compras.Numero_Compra, Compras.Observaciones, Compras.Fecha_Compra, Compras.MonedaCompra, Compras.Cod_Proveedor, Compras.Nombre_Proveedor, Compras.Apellido_Proveedor, Compras.Fecha_Vencimiento, Compras.SubTotal, Compras.Descuento, Compras.IVA, Compras.NetoPagar, Compras.Pagado , Compras.Marca, Compras.Contabilizado, Compras.Tipo_Compra, Proveedor.Cod_Cuenta_Pagar, Proveedor.Cod_Cuenta_Cobrar,Compras.Su_Referencia, Compras.Nuestra_Referencia FROM Proveedor INNER JOIN Compras ON Proveedor.Cod_Proveedor = Compras.Cod_Proveedor  " & _
                                    "WHERE  (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (Compras.Marca = 1) AND (Compras.Contabilizado = 0) AND (Compras.Tipo_Compra = '" & TipoFactura & "') ORDER BY Compras.Fecha_Compra, Compras.Numero_Compra"
        
    
    End Select
    

                 '//////////////////////////////////////////////////////////////////////////////////////////////
                 '//////////SI LA CUENTA EXISTE AGREGO LOS ENCABEZADOS///////////////////////////////////////
                 '/////////////////////////////////////////////////////////////////////////////////////////////
                 
                         mes = Month(Me.DTPicker6.Value)
                         Año = Year(Me.DTPicker6.Value)
                         FechaIni = CDate("1/" & Month(Me.DTPicker6.Value) & "/" & Year(Me.DTPicker6.Value))
                         FechaFin = DateSerial(Año, mes + 1, 1 - 1)

                 
                         Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
                         Me.AdoConsulta.Refresh
                         If Not Me.AdoConsulta.Recordset.EOF Then
                           Periodo = Me.AdoConsulta.Recordset("Periodo")
                            NumeroPeriodo = Me.AdoConsulta.Recordset("NPeriodo")
                            EstadoPeriodo = Me.AdoConsulta.Recordset("EstadoPeriodo")
                      

                              Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                              Me.AdoConsulta.Recordset.Update
                              NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                              
                              
 
                             Me.AdoProcesos.RecordSource = SqlString
                             Me.AdoProcesos.Refresh
                             Me.osProgress1.Visible = True
                             Me.osProgress1.Min = 0
                             Me.osProgress1.Value = 0
                             If Not Me.AdoProcesos.Recordset.EOF Then
                             Me.AdoProcesos.Recordset.MoveFirst
                             Me.osProgress1.Max = Me.AdoProcesos.Recordset.RecordCount
                             End If
                             Me.AdoProcesos.Refresh
                             
                             Beneficiario = ""
                             Ret1Porc = 0
                             Ret2Porc = 0
                             Ret3Porc = 0
                             Ret4Porc = 0
                             
                             
                             
                        Do While Not Me.AdoProcesos.Recordset.EOF
                           
                           Select Case TipoFactura
                              Case "Planilla"
                              
                                MontoBanco = 0
                                FechaFactura = Me.AdoProcesos.Recordset("Fecha_Recibo")
                                FechaVence = Me.AdoProcesos.Recordset("Fecha_Recibo")
                                NumeroFactura = Me.AdoProcesos.Recordset("CodReciboPago")
                                CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                
                                Ret1Porc = Mid(Me.AdoProcesos.Recordset("Retencion1"), 1, Len(Me.AdoProcesos.Recordset("Retencion1")) - 1)
                                Ret2Porc = Mid(Me.AdoProcesos.Recordset("Retencion2"), 1, Len(Me.AdoProcesos.Recordset("Retencion2")) - 1)
                                Ret3Porc = Mid(Me.AdoProcesos.Recordset("Retencion3"), 1, Len(Me.AdoProcesos.Recordset("Retencion3")) - 1)
                                Ret4Porc = Mid(Me.AdoProcesos.Recordset("Retencion4"), 1, Len(Me.AdoProcesos.Recordset("Retencion4")) - 1)
                                
                                If CodigoCuentaCliente = "" Then
                                  MsgBox "No existe la cuenta del Proveedor", vbCritical, "Zeus contable"
                                  Exit Sub
                                End If
                                
                                    SubTotal = 0
                                    Descuento = 0
                                    Iva = 0
                                    NetoPagar = 0
                                    Pagado = 0
                    

                                Me.AdoConsultaFactura.RecordSource = "SELECT DetalleReciboPago.idDetallePago, DetalleReciboPago.CodReciboPago, DetalleReciboPago.Fecha_Recibo, DetalleReciboPago.Numero_Compra, DetalleReciboPago.MontoPagado , Detalle_MetodoPagoProveedores.NombrePago, Detalle_MetodoPagoProveedores.Monto, MetodoPago.Cod_Cuenta FROM DetalleReciboPago INNER JOIN ReciboPago ON DetalleReciboPago.CodReciboPago = ReciboPago.CodReciboPago AND DetalleReciboPago.Fecha_Recibo = ReciboPago.Fecha_Recibo INNER JOIN Detalle_MetodoPagoProveedores ON DetalleReciboPago.CodReciboPago = Detalle_MetodoPagoProveedores.CodReciboPago INNER JOIN MetodoPago ON Detalle_MetodoPagoProveedores.NombrePago = MetodoPago.NombrePago  " & _
                                                                     "WHERE (DetalleReciboPago.CodReciboPago = '" & NumeroFactura & "') AND (DetalleReciboPago.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) ORDER BY DetalleReciboPago.Fecha_Recibo"
                                Me.AdoConsultaFactura.Refresh
                                If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                     DescripcionRecibo = "Pago a Proveedores"  'Me.AdoConsultaFactura.Recordset("Descripcion")
                                  If Not IsNull(Me.AdoConsultaFactura.Recordset("MontoPagado")) Then
                                    If Me.AdoConsultaFactura.Recordset("MontoPagado") = Me.AdoProcesos.Recordset("Sub_Total") Then
                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("Sub_Total")), "##,##0.00")
                                        If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                          Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                        End If
                                        NetoPagar = Format(Val(Me.AdoProcesos.Recordset("Total")), "##,##0.00")
                                    Else
                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("Sub_Total")), "##,##0.00")
                                        If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                         Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                        End If
                                        NetoPagar = Me.AdoConsultaFactura.Recordset("MontoPagado")
                                        


                                    End If
                               
                                
                                
                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                       MonedaFactura = Me.AdoProcesos.Recordset("MonedaRecibo")
                                       TasaMovimiento = 1
    
                                        SqlString = "SELECT * From TasaCambio WHERE (FechaTasa = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102))"
                                        Me.AdoBuscaFacturacion.RecordSource = SqlString
                                        Me.AdoBuscaFacturacion.Refresh
                                        If Not Me.AdoBuscaFacturacion.Recordset.EOF Then
                                          TasaCambio = Format(Val(Me.AdoBuscaFacturacion.Recordset("MontoTasa")), "##,##0.0000")
                                        Else
                                          TasaCambio = 1
                                        End If
                                   End If
                                 End If
                                 
                                      If MonedaFactura = "Cordobas" Then
                                        TasaCambio = 1
                                        MonedaMovimiento = "Córdobas"
                                      End If
                                 
                                 
                                      If Reg = 1 Then
                                         '////////////////////////////////////////////////////////////////
                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                         '/////////////////////////////////////////////////////////////////
                                         MonedaMovimiento = "Córdobas"
                                         If Me.ChkCheques.Value = xtpChecked Then
                                           Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker6.Value, "yyyy-mm-dd"), "Movimiento de Pago a Proveedores", "CHEQUE", MonedaMovimiento)
                                         Else
                                           Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker6.Value, "yyyy-mm-dd"), "Movimiento de Pago a Proveedores", "Pago", MonedaMovimiento)
                                         End If
                                         Reg = 2
                                      End If
                                      
                                       DescripcionMovimiento = "Registrando Recibo No " & NumeroFactura & "  " & DescripcionRecibo
                                      

                                      '///////////////////////////////////////////////////////////////////////////////////////
                                      '/////////////////////////////GRABA DETALLE DE RECIBO PAGOS///////////////////////////////////
                                      '///////////////////////////////////////////////////////////////////////////////////////
                                       TotalRetencion = 0
                                       Registro = 1
                                       CodigoCuentaBanco = ""
'                                       Me.AdoConsultaFactura.RecordSource = "SELECT  DetalleReciboPago.idDetallePago, DetalleReciboPago.CodReciboPago, DetalleReciboPago.Fecha_Recibo, DetalleReciboPago.Numero_Compra, DetalleReciboPago.MontoPagado  FROM  DetalleRecibo INNER JOIN MetodoPago ON DetalleRecibo.NombrePago = MetodoPago.NombrePago  " & _
'                                                                            "WHERE  (DetalleRecibo.CodReciboPago = '" & NumeroFactura & "') AND (DetalleRecibo.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) AND (DetalleRecibo.MontoPagado < 0) ORDER BY DetalleRecibo.Fecha_Recibo"
                                       Me.AdoConsultaFactura.RecordSource = "SELECT DetalleReciboPago.idDetallePago, DetalleReciboPago.CodReciboPago, DetalleReciboPago.Fecha_Recibo, DetalleReciboPago.Numero_Compra, DetalleReciboPago.MontoPagado , Detalle_MetodoPagoProveedores.NombrePago, Detalle_MetodoPagoProveedores.Monto, MetodoPago.Cod_Cuenta FROM DetalleReciboPago INNER JOIN ReciboPago ON DetalleReciboPago.CodReciboPago = ReciboPago.CodReciboPago AND DetalleReciboPago.Fecha_Recibo = ReciboPago.Fecha_Recibo INNER JOIN Detalle_MetodoPagoProveedores ON DetalleReciboPago.CodReciboPago = Detalle_MetodoPagoProveedores.CodReciboPago INNER JOIN MetodoPago ON Detalle_MetodoPagoProveedores.NombrePago = MetodoPago.NombrePago  " & _
                                                                            "WHERE (DetalleReciboPago.CodReciboPago = '" & NumeroFactura & "') AND (DetalleReciboPago.Fecha_Recibo = CONVERT(DATETIME, '" & Format(FechaFactura, "yyyy-MM-dd") & "', 102)) ORDER BY DetalleReciboPago.Fecha_Recibo"
                                       Me.AdoConsultaFactura.Refresh
                                   Do While Not Me.AdoConsultaFactura.Recordset.EOF
                                      
                                       NumeroCompra = Me.AdoConsultaFactura.Recordset("Numero_Compra")
                                      
                                       Debito = Abs(Me.AdoConsultaFactura.Recordset("MontoPagado"))
                                       Debito = Format(Debito * TasaCambio, "##,##0.00")
                                       Credito = 0
                                       CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                        
'                                        DescripcionMovimiento = "Movimiento de Registro de Recibos No" & NumeroFactura
                                        If Me.ChkCheques.Value = xtpChecked Then
                                         Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroCompra, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                        Else
                                         Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Pago", NumeroCompra, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ReciboPago")
                                        End If
                                        

                                        CodigoCuentaBanco = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                                               
                                       
                                        Me.AdoConsultaFactura.Recordset.MoveNext
                                        TotalRetencion = Abs(Debito) + TotalRetencion
                                     Loop
                                     
                                     
                                     
                                      
                                     If Registro = 1 Then
                                                     
                                             MontoBanco = SubTotal - (SubTotal * (Ret1Porc / 100)) - (SubTotal * (Ret2Porc / 100)) - (SubTotal * (Ret3Porc / 100)) - (SubTotal * (Ret4Porc / 100))
                                            
                                             NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                             Debito = 0
                                             Credito = SubTotal
                                             Credito = Format(MontoBanco * TasaCambio, "##,##0.00")
                                              
                                             
                                            '///////////////////////////////////////////////////////////////////////////////////
                                            '///////////////////////GRABO EL MOVIMIENTO DEL BANCO ///////////////////////////
                                            '///////////////////////////////////////////////////////////////////////////////////
                                             If Me.ChkCheques.Value = xtpChecked Then
                                               Beneficiario = Me.AdoProcesos.Recordset("Nombre_Proveedor")
                                               Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             Else
                                               Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Pago", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ReciboPago")
                                             End If
                                             Registro = 2
                                       
                                       
                                             '////////////////////////////////////////////////////////////////////////////
                                             '//////////////////////////AGREGO RETENCIONES /////////////////////////////////
                                             '////////////////////////////////////////////////////////////////////////////////
                                             If Ret1Porc <> 0 Then
                                              CodigoCuentaBanco = BuscaCuentaImpuestos(Ret1Porc & "%")
                                              NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                              Debito = 0
                                             
                                              Credito = Format(SubTotal * (Ret1Porc / 100) * TasaCambio, "##,##0.00")
                                              Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                             
                                             If Ret2Porc <> 0 Then
                                               CodigoCuentaBanco = BuscaCuentaImpuestos(Ret2Porc & "%")
                                               NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                               Debito = 0
                                               Credito = Format(SubTotal * (Ret2Porc / 100) * TasaCambio, "##,##0.00")
                                               Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                             
                                            If Ret3Porc <> 0 Then
                                              CodigoCuentaBanco = BuscaCuentaImpuestos(Ret3Porc & "%")
                                              NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                              Debito = 0
                                             
                                              Credito = Format(SubTotal * (Ret3Porc / 100) * TasaCambio, "##,##0.00")
                                              Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                       
                                            If Ret4Porc <> 0 Then
                                              CodigoCuentaBanco = BuscaCuentaImpuestos(Ret4Porc & "%")
                                              NombreCuenta = BuscaCuenta(CodigoCuentaBanco)
                                              Debito = 0
                                             
                                              Credito = Format(SubTotal * (Ret4Porc / 100) * TasaCambio, "##,##0.00")
                                              Resultado = GrabaDetalleFactura(CodigoCuentaBanco, Format(Me.DTPicker6.Value, "dd/mm/yyyy"), NumeroTransaccion, NumeroPeriodo, NombreCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "CHEQUE", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "ChequePago")
                                             End If
                                       End If

                                      

                                      
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                rs.Open "UPDATE ReciboPago SET Contabilizado = 1 ,Activo = 0  WHERE (CodReciboPago = '" & NumeroFactura & "') ", ConexionFacturacion
                                      
                                      
                              
                              Case Else
                             
                             
                             
                                                FechaFactura = Me.AdoProcesos.Recordset("Fecha_Compra")
                                                If Not IsNull(Me.AdoProcesos.Recordset("Fecha_Vencimiento")) Then
                                                  FechaVence = Me.AdoProcesos.Recordset("Fecha_Vencimiento")
                                                End If
                                                NumeroFactura = Me.AdoProcesos.Recordset("Numero_Compra")
                                                MonedaFactura = Me.AdoProcesos.Recordset("MonedaCompra")
                                                If Not IsNull(Me.AdoProcesos.Recordset("Su_Referencia")) Then
                                                   NumeroReferencia = "Compra No:" & NumeroFactura & " " & "Referencia: " & Me.AdoProcesos.Recordset("Su_Referencia")
                                                Else
                                                   NumeroReferencia = "Compra No:" & NumeroFactura
                                                End If
                                                
                                                If Not IsNull(Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")) Then
                                                  CodigoCuentaCliente = Me.AdoProcesos.Recordset("Cod_Cuenta_Pagar")
                                                Else
                                                  CodigoCuentaCliente = "2121"
                                                End If
                                                
                                                Select Case TipoFactura
                                                  Case "Cuenta"
'                                                     codigocuentacliente =
                                                
                                                End Select
                                                
                                                
                                                    SubTotal = 0
                                                    Descuento = 0
                                                    Iva = 0
                                                    NetoPagar = 0
                                                    TasaCambio = 1
                                                    TasaMovimiento = 1
                                                    
                                                    
                                                    SqlString = "SELECT * From Detalle_Compras WHERE (Numero_Compra = '" & NumeroFactura & "') AND (Tipo_Compra = '" & TipoFactura & "')"
                                                    Me.AdoBuscaFacturacion.RecordSource = SqlString
                                                    Me.AdoBuscaFacturacion.Refresh
                                                    
                                                    If MonedaFactura = "Dolares" Then
                                                         TasaCambio = Format(Val(Me.AdoBuscaFacturacion.Recordset("TasaCambio")), "##,##0.000000")
                                                    Else
                                                         TasaCambio = 1
                                                    End If
                                                
                
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////////////////BUSCO LOS DETALLES DE LACOMPRA ////////////////////////////////////////////////////
                                                '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                Me.AdoConsultaFactura.RecordSource = "SELECT SUM(Detalle_Compras.Cantidad) AS Cantidad, SUM(Detalle_Compras.Precio_Unitario) AS Precio_Unitario, SUM(Detalle_Compras.Descuento) AS Descuento, SUM(Detalle_Compras.Precio_Neto) AS Precio_Neto, SUM(Detalle_Compras.Importe) AS Importe FROM  Compras INNER JOIN Detalle_Compras ON Compras.Numero_Compra = Detalle_Compras.Numero_Compra AND Compras.Fecha_Compra = Detalle_Compras.Fecha_Compra And Compras.Tipo_Compra = Detalle_Compras.Tipo_Compra  " & _
                                                                                      "WHERE (Compras.Numero_Compra = '" & NumeroFactura & "') AND (Compras.Nombre_Proveedor <> N'******CANCELADO') AND (Compras.Tipo_Compra = '" & TipoFactura & "')"
                                                Me.AdoConsultaFactura.Refresh
                                                If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                                  If Not IsNull(Me.AdoConsultaFactura.Recordset("Importe")) Then
                                                    If Format(CDbl(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00") = Format(CDbl(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00") Then
                                                        SubTotal = Format(Val(Me.AdoProcesos.Recordset("SubTotal")), "##,##0.00")
                                                        SubTotal = Format(SubTotal * TasaCambio, "##,##0.00")
                                                        If Not IsNull(Me.AdoProcesos.Recordset("Descuento")) Then
                                                         If IsNumeric(Me.AdoProcesos.Recordset("Descuento")) Then
                                                          Descuento = Format(Val(Me.AdoProcesos.Recordset("Descuento")), "##,##0.00")
                                                          Descuento = Format(Descuento * TasaCambio, "##,##0.00")
                                                         End If
                                                        End If
                                                        Iva = Format(Val(Me.AdoProcesos.Recordset("IVA")), "##,##0.00")
                                                        Iva = Format(Iva * TasaCambio, "##,##0.00")
                                                        NetoPagar = Format(Val(Me.AdoProcesos.Recordset("NetoPagar")), "##,##0.00")
                                                        NetoPagar = Format(NetoPagar * TasaCambio, "##,##0.00")
                                                        Pagado = Format(Val(Me.AdoProcesos.Recordset("Pagado")), "##,##0.00")
                                                        Pagado = Format(Pagado * TasaCambio, "##,##0.00")
                                                    Else
                                                        SubTotal = Format(Val(Me.AdoConsultaFactura.Recordset("Importe")), "##,##0.00")
                                                        SubTotal = Format(SubTotal * TasaCambio, "##,##0.00")
                                                        Descuento = 0
                                                        Iva = 0
                                                        NetoPagar = 0
                                                        Pagado = 0
                                                    End If
                                               
                                                
                                                
                                                       NombreCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                       MonedaFactura = Me.AdoProcesos.Recordset("MonedaCompra")
                                                       
                
                
                    
                
                                                
                                                      If Reg = 1 Then
                                                         '////////////////////////////////////////////////////////////////
                                                         '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                                                         '/////////////////////////////////////////////////////////////////
                                                         MonedaMovimiento = "Córdobas"
                                                         Resultado = GrabaEncabezado(NumeroPeriodo, NumeroTransaccion, Format(Me.DTPicker6.Value, "yyyy-mm-dd"), "Movimiento de Compras", "Comp", MonedaMovimiento)
                                                         Reg = 2
                                                      End If
                                                           
                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                        '////////////AGREGO LA CUENTA DEL PROVEEDOR//////////////////////////////////////
                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                           
                                                           Credito = 0
                                                           Debito = 0
                                                          If NetoPagar = 0 Then
                                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                                                                                      
                                                            Me.AdoConsultaFactura.RecordSource = "SELECT Detalle_MetodoCompras.Numero_Compra, Detalle_MetodoCompras.Fecha_Compra, Detalle_MetodoCompras.Tipo_Compra, Detalle_MetodoCompras.NombrePago, Detalle_MetodoCompras.Monto, Detalle_MetodoCompras.NumeroTarjeta, Detalle_MetodoCompras.FechaVence , MetodoPago.TipoPago, MetodoPago.Cod_Cuenta, MetodoPago.Moneda FROM Detalle_MetodoCompras INNER JOIN MetodoPago ON Detalle_MetodoCompras.NombrePago = MetodoPago.NombrePago  " & _
                                                                                                 "WHERE (Detalle_MetodoCompras.Numero_Compra = '" & NumeroFactura & "')"
                                                            Me.AdoConsultaFactura.Refresh
                                                            If Not Me.AdoConsultaFactura.Recordset.EOF Then
                                                                CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                                
                                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                                NetoPagar = SubTotal + Iva - Descuento
                                                                Select Case TipoFactura
                                                                  Case "Transferencia Recibida"
                                                                       DescripcionMovimiento = "Registro de Transferencia Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                                                                  
                                                                  
                '                                                       DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Recepcion"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                                                                  
                                                                  
                '                                                       DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de Devolucion Numero " & NumeroFactura
                                                                       Debito = Format(NetoPagar, "##,##0.00")
                                                                       Credito = 0
                                                                End Select
                                                            Else
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                                NetoPagar = SubTotal + Iva - Descuento
                                                                Debito = Format(NetoPagar, "##,##0.00")
                                                            End If
                                                           ElseIf NetoPagar < 0 Then
                                                              '/////////////////AGREGO LA DIFERENICIA A OTROS INGRESOS /
                                                               Me.AdoConsulta.RecordSource = "SELECT * From Cuentas WHERE (TipoCuenta = 'Capital')"
                                                               Me.AdoConsulta.Refresh
                                                               If Not Me.AdoConsulta.Recordset.EOF Then
                                                                CodigoCuentaOtros = Me.AdoConsulta.Recordset("CodCuentas")
                                                                DescripcionCuenta = BuscaCuenta(CodigoCuentaOtros)
                                                               End If
                                                               
                                                               
                                                               Select Case TipoFactura
                                                                  Case "Transferencia Recibida"
                                                                       DescripcionMovimiento = "Registro de AJUSTE Transferencia Recibida Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                       
                                                                  Case "Recepcion"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Compra Numero " & NumeroFactura
                                                                       Credito = Format(NetoPagar, "##,##0.00")
                                                                       Debito = 0
                                                                  Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                       DescripcionMovimiento = "Registro de AJUSTE Devolucion Compra Numero " & NumeroFactura
                                                                       Debito = Format(NetoPagar, "##,##0.00")
                                                                       Credito = 0
                                                                End Select
                                                               
                                                               
                                                               Resultado = GrabaDetalleFactura(CodigoCuentaOtros, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, 0, Abs(NetoPagar), "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                               
                                                            '////////////////////////////////SIGNIFICA QUE EL PAGO DE REALIZAO DE CONTADO ////
                                                            Me.AdoConsultaFactura.RecordSource = "SELECT  * FROM Detalle_MetodoFacturas INNER JOIN MetodoPago ON Detalle_MetodoFacturas.NombrePago = MetodoPago.NombrePago  " & _
                                                                                                 "WHERE (Detalle_MetodoFacturas.Numero_Factura = '" & NumeroFactura & "')"
                                                            Me.AdoConsultaFactura.Refresh
                '                                              CodigoCuentaCliente = Me.AdoConsultaFactura.Recordset("Cod_Cuenta")
                                                              NetoPagar = SubTotal + Iva - Descuento + Abs(NetoPagar)
                                                              
                                                              Debito = Format(NetoPagar, "##,##0.00")
                                                              DescripcionMovimiento = "Registro de Facturacion Factura Numero " & NumeroFactura
                                                              DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                              
                                                           ElseIf NetoPagar > 0 Then
                                                           
                                              
                                                            DescripcionCuenta = BuscaCuenta(CodigoCuentaCliente)
                                                           
                                                           Select Case TipoFactura
                                                            Case "Transferencia Recibida"
                                                                   DescripcionMovimiento = "Registro de Transferencia Recibida Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = 0
                                                                   Credito = Format(NetoPagar, "##,##0.00")
                
                                                                  '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             Case "Recepcion"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Compra Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = 0
                                                                   Credito = Format(NetoPagar, "##,##0.00")
                                                                   
                                                                  DescripcionMovimiento = DescripcionMovimiento & "," & Me.AdoProcesos.Recordset("Observaciones")
                                                                  '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             
                                                              Case "Cuenta"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Compra Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = 0
                                                                   Credito = Format(NetoPagar, "##,##0.00")
                                                                   
                                                                  DescripcionMovimiento = DescripcionMovimiento & "," & Me.AdoProcesos.Recordset("Observaciones")
                                                                  '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             
                                                             Case "Devolucion de Compra"
                                                                        '////////////////////////////////////////////////////////////////////////////////////
                                                                        '////////////AGREGO LA CUENTA DEL CLIENTE//////////////////////////////////////
                                                                        '///////////////////////////////////////////////////////////////////////////////////////
                                                                           
                                                                        DescripcionMovimiento = ""
                                                                        If Me.ChkDescripcionCompra.Value = 1 Then
                                                                          Me.AdoBuscaFacturacion.RecordSource = "SELECT  * FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "')"
                                                                          Me.AdoBuscaFacturacion.Refresh
                                                                          Do While Not Me.AdoBuscaFacturacion.Recordset.EOF
                                                                            DescripcionMovimiento = DescripcionMovimiento & Me.AdoBuscaFacturacion.Recordset("Cantidad") & " " & Me.AdoBuscaFacturacion.Recordset("Unidad_Medida") & " " & Me.AdoBuscaFacturacion.Recordset("Descripcion_Producto") & ", "
                                                                          
                                                                            Me.AdoBuscaFacturacion.Recordset.MoveNext
                                                                          Loop
                                                                        
                                                                        Else
                                                                          DescripcionMovimiento = "Registro de Compra Numero " & NumeroFactura
                                                                        End If
                '                                                   DescripcionMovimiento = "Registro de Facturacion Devolucion Numero " & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                   Debito = Format(NetoPagar, "##,##0.00")
                                                                   Credito = 0
                                                        '/////////////////////GRABO LA CUENA DEL CLIENTE////////////////////////////////////////////
                                                                  Resultado = GrabaDetalleFactura(CodigoCuentaCliente, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                             End Select
                                                                
                
                                                           End If
                 
'++++++++++++++++++++++++++++++++++++++++++ DETALLE DE MOVIMIENTOS ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                                    '//////////////////////////////////////////////////////////////////////////////////////////
                                                    '/////////////////////CARGO LOS DETALLE DE LAS FACTURAS////////////////////////////////////
                                                    '//////////////////////////////////////////////////////////////////////////////////////////
                                                     
                                                     If TipoFactura = "Cuenta" Then
                                                        SqlString = "SELECT  Numero_Compra, Fecha_Compra, Tipo_Compra, Cod_Producto, Cantidad, Precio_Unitario, Descuento, Precio_Neto, Importe, TasaCambio From Detalle_Compras WHERE (Numero_Compra = '" & NumeroFactura & "') AND (Tipo_Compra LIKE '%" & TipoFactura & "%')"
                                                     
                                                     Else
                                                         SqlString = "SELECT  Detalle_Compras.Numero_Compra, Detalle_Compras.Fecha_Compra, Detalle_Compras.Tipo_Compra, Detalle_Compras.Cod_Producto, Detalle_Compras.Cantidad, Detalle_Compras.Precio_Unitario, Detalle_Compras.Descuento, Detalle_Compras.Precio_Neto, Detalle_Compras.Importe, Productos.Descripcion_Producto, Productos.Cod_Cuenta_Inventario, Productos.Cod_Cuenta_Costo, Productos.Cod_Cuenta_Ventas, Productos.Cod_Cuenta_GastoAjuste , Productos.Cod_Cuenta_IngresoAjuste, Detalle_Compras.TasaCambio, Productos.Costo_Promedio, Productos.Costo_Promedio_Dolar  FROM  Detalle_Compras INNER JOIN Productos ON Detalle_Compras.Cod_Producto = Productos.Cod_Productos  " & _
                                                                     "WHERE (Detalle_Compras.Numero_Compra = '" & NumeroFactura & "') AND (Detalle_Compras.Tipo_Compra Like '%" & TipoFactura & "%')"
                                                     End If
                                                     Me.AdoProcesosFacturacion.RecordSource = SqlString
                                                     Me.AdoProcesosFacturacion.Refresh
                                                     Do While Not Me.AdoProcesosFacturacion.Recordset.EOF
                                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                CodigoProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                If TipoFactura = "Cuenta" Then
                                                                  CodigoCuentaProducto = Me.AdoProcesosFacturacion.Recordset("Cod_Producto")
                                                                Else
                                                                  CodigoCuentaProducto = BuscaCodigoProducto(CodigoProducto)
                                                                End If
                                                                If MonedaFactura = "Dolares" Then
                '                                                 TasaCambio = BuscaTasaCambio(Fecha)
                                                                   TasaCambio = Format(Val(Me.AdoProcesosFacturacion.Recordset("TasaCambio")), "##,##0.0000")
                                                                Else
                                                                    TasaCambio = 1
                                                               End If
                                                                
                                                                
                                                                If MonedaFactura = "Dolares" Then
                                                                  CostoProducto = Me.AdoProcesosFacturacion.Recordset("Importe")
                '                                                   CostoProducto = Format(Me.AdoProcesosFacturacion.Recordset("Importe") * TasaCambio, "##,##0.00")
                                                                Else
                                                                   CostoProducto = Format(Me.AdoProcesosFacturacion.Recordset("Importe"), "##,##0.00")
                                                                End If
                                                            
                                                               '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                               '///////////////////////////CALCULO EL IVA DE CADA PRODUCTO//////////////////////////////////////////////
                                                               '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                
                                                                If Iva <> 0 Then
                                                                 TasaIva = BuscaTasaIva(CodigoProducto)
                                                                 DescripcionCuenta = BuscaCuenta(CodigoCuentaIva)
                                                                 
                                                                 QUIEN = ""
                                                                 
                                                                 If DescripcionCuenta <> "Nulo" Then
                                                                  
                                                                  Select Case TipoFactura
                                                                    Case "Transferencia Recibida"
                                                                         DescripcionMovimiento = "IVA Ventas Compra No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                         Credito = 0
                                                                         Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                            If Debito <> 0 Then
                                                                                Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                            End If
                                                                    Case "Recepcion"
                                                                         DescripcionMovimiento = "IVA Ventas Compra No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                         Credito = 0
                                                                         Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                            If Debito <> 0 Then
                                                                                Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                            End If
                                                                            
                                                                     Case "Cuenta"
                                                                         DescripcionMovimiento = "IVA Ventas Compra No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                         Credito = 0
                                                                         Debito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                         Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                            If Debito <> 0 Then
                                                                                Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroFactura, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                            End If
                                                                    Case "Devolucion de Compra"
                                                                      DescripcionMovimiento = "IVA Ventas Devolucion No " & NumeroFactura & " Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                      Debito = 0
                                                                      Credito = Format(Val(Me.AdoProcesosFacturacion.Recordset("Importe")) * TasaIva, "##,##0.00")
                                                                      Credito = Format(Credito * TasaCambio, "##,##0.00")
                                                                     If Credito <> 0 Then
                                                                         Resultado = GrabaDetalleFactura(CodigoCuentaIva, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                     End If
                                                                    End Select
                
                
                                                                   End If
                                                                 End If
                                                                
                                                              QUIEN = ""
                                                                
                                                                 
                                                                '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '///////////////////////////////////////BUSCO LA CUENTA DE INVENTARIO///////////////////////////////////////////
                                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                If TipoFactura <> "Cuenta" Then
                                                                    If Not IsNull(Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")) Then
                                                                      CodigoCuentaInventario = Me.AdoProcesosFacturacion.Recordset("Cod_Cuenta_Inventario")
                                                                    End If
                                                                Else
                                                                    CodigoCuentaInventario = CodigoCuentaProducto
                                                                End If
                                                                
                                                                 
                                                                 DescripcionCuenta = BuscaCuenta(CodigoCuentaProducto)
                                                                   Select Case TipoFactura
                                                                    Case "Transferencia Recibida"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                          Credito = 0
                                                                         'If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         'End If
                                                                    Case "Recepcion"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                          Credito = 0
                                                                         'If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         'End If
                                                                
                                                                    Case "Cuenta"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Debito = Format(Debito * TasaCambio, "##,##0.00")
                                                                          Credito = 0
                                                                         'If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         'End If
                                                                    Case "Devolucion de Compra"
                                                                          DescripcionMovimiento = "Costo Producto " & CodigoProducto & "   Referencia " & NumeroReferencia
                                                                          Debito = 0
                                                                          Credito = Format(Val(CostoProducto), "##,##0.00")
                                                                          Credito = Format(Val(CostoProducto) * TasaCambio, "##,##0.00")
                                                                         If DescripcionCuenta <> "Nulo" Then
                                                                          Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaMovimiento, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                         End If
                                                                    End Select
                            
                            
                                                        Me.AdoProcesosFacturacion.Recordset.MoveNext
                                                     Loop
                                                     
                                                     
                                                                 '////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '///////////////////////////////////////BUSCO LA CUENTA DE INVENTARIO///////////////////////////////////////////
                                                                '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                If Pagado <> 0 Then
                                                                  SqlString = "SELECT Detalle_MetodoCompras.Numero_Compra, Detalle_MetodoCompras.Fecha_Compra, Detalle_MetodoCompras.NombrePago, Detalle_MetodoCompras.Monto,Detalle_MetodoCompras.NumeroTarjeta , MetodoPago.Cod_Cuenta, MetodoPago.Moneda FROM  Detalle_MetodoCompras INNER JOIN MetodoPago ON Detalle_MetodoCompras.NombrePago = MetodoPago.NombrePago " & _
                                                                              "WHERE (Detalle_MetodoCompras.Numero_Compra = '" & NumeroFactura & "') "
                                                                  Me.AdoConsultaFacturacion.RecordSource = SqlString
                                                                  Me.AdoConsultaFacturacion.Refresh
                                                                  Do While Not Me.AdoConsultaFacturacion.Recordset.EOF
                                                                     CodigoCuentaMetodo = Me.AdoConsultaFacturacion.Recordset("Cod_Cuenta")
                                                                     
                                                                     DescripcionCuenta = BuscaCuenta(CodigoCuentaMetodo)
                                                                    Select Case TipoFactura
                                                                        Case "Recepcion"
                                                                         DescripcionMovimiento = "PAGO DE COMPRA" & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                         Debito = 0
                                                                         Credito = Me.AdoConsultaFacturacion.Recordset("Monto")
                                                                             If DescripcionCuenta <> "Nulo" Then
                                                                              Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                             End If
                                                                        Case "Cuenta"
                                                                         DescripcionMovimiento = "PAGO DE COMPRA" & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                         Debito = 0
                                                                         Credito = Me.AdoConsultaFacturacion.Recordset("Monto")
                                                                             If DescripcionCuenta <> "Nulo" Then
                                                                              Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                             End If
                                                                        Case "Devolucion de Compra"
                                                                         DescripcionMovimiento = "PAGO DE DEVOLUCION" & NumeroFactura & "   Referencia " & NumeroReferencia
                                                                         Debito = Me.AdoConsultaFacturacion.Recordset("Monto")
                                                                         Credito = 0
                                                                             If DescripcionCuenta <> "Nulo" Then
                                                                              Resultado = GrabaDetalleFactura(CodigoCuentaInventario, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Credito", TasaCambio, Debito, Credito, "Comp", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                             End If
                                                                    End Select
                                                                    
                                                                    
                                                                
                
                                                                     If Debito <> 0 Then
                                                                         Resultado = GrabaDetalleFactura(CodigoCuentaMetodo, Me.DTPicker6.Value, NumeroTransaccion, NumeroPeriodo, DescripcionCuenta, DescripcionMovimiento, "Debito", TasaCambio, Debito, Credito, "VTAS", NumeroReferencia, FechaFactura, Descuento, FechaVence, CodigoCuentaCliente, "FacturaVenta")
                                                                     End If
                                                                 
                                                                    Me.AdoConsultaFacturacion.Recordset.MoveNext
                                                                  Loop
                                                                
                                                                End If
                                                                
                                                                
                                                                
                                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                '/////////////////////////////////////////////ACTUALIZO LA FACTURA COMO CONTABILIZADO //////////////////////////////////////////
                                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                rs.Open "UPDATE Compras SET Contabilizado = 1 ,Activo = 0  WHERE (Numero_Compra = '" & NumeroFactura & "') ", ConexionFacturacion
                                                     
                                                     
                                                     
                                                  End If
                                                
                                                End If
                
                                End Select
                                                                                       
                                Me.osProgress1.Value = Me.osProgress1.Value + 1
                                Me.AdoProcesos.Recordset.MoveNext
                            Loop
                             
'

                         End If

                                    
End Sub

Private Sub PushButton3_Click()
Dim SqlString As String, FechaInicio As String, FechaFin As String
Dim TipoFactura As String

    If Me.OptRecepcion.Value = True Then
     TipoFactura = "Recepcion"
    ElseIf Me.OptPlanilla.Value = True Then
     TipoFactura = "Pago Proveedor"
    ElseIf Me.OptPlanillaTransportista.Value = True Then
     TipoFactura = "Pago Transportista"
    ElseIf Me.OptLiquidacion.Value = True Then
     TipoFactura = "LiquidacionLeche"
    End If
    
    Me.DTPicker10.Value = Me.DTPicker12.Value
    
   Select Case TipoFactura
       
            Case "Recepcion"
        
            Me.TDGridPlanillaLeche.Columns(0).DataField = "Fecha_Compra"
            Me.TDGridPlanillaLeche.Columns(1).DataField = "Numero_Compra"
            Me.TDGridPlanillaLeche.Columns(2).DataField = "Nombre_Proveedor"
            Me.TDGridPlanillaLeche.Columns(3).DataField = "SubTotal"
            Me.TDGridPlanillaLeche.Columns(4).DataField = "Descuento"
            Me.TDGridPlanillaLeche.Columns(5).DataField = "IVA"
            Me.TDGridPlanillaLeche.Columns(6).DataField = "NetoPagar"
            Me.TDGridPlanillaLeche.Columns(7).DataField = "Marca"
            Me.TDGridPlanillaLeche.Splits(0).Caption = "Listado de Recepcion"
            
            
            FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
            SqlString = "SELECT  Fecha_Compra, Numero_Compra, Nombre_Proveedor, SubTotal, Descuento,IVA, NetoPagar,Marca From Compras  " & _
                         "WHERE   (Contabilizado = 0) AND (Tipo_Compra = '" & TipoFactura & "') AND (Fecha_Compra BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102))"

      Case "Pago Proveedor"
        
            Me.TDGridPlanillaLeche.Columns(0).DataField = "NumPlanilla"
            Me.TDGridPlanillaLeche.Columns(1).DataField = "CodTipoNomina"
            Me.TDGridPlanillaLeche.Columns(2).DataField = "FechaInicial"
            Me.TDGridPlanillaLeche.Columns(3).DataField = "FechaFinal"
            Me.TDGridPlanillaLeche.Columns(4).DataField = "Año"
            Me.TDGridPlanillaLeche.Columns(5).DataField = "mes"
            Me.TDGridPlanillaLeche.Columns(6).DataField = "Periodo"
            Me.TDGridPlanillaLeche.Columns(7).DataField = "Marca"
            Me.TDGridPlanillaLeche.Splits(0).Caption = "Listado de Planillas"
            
            
            FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
            SqlString = "SELECT   NumPlanilla, CodTipoNomina, FechaInicial, FechaFinal, Año, mes, Periodo, Marca From Nomina Where (Contabilizado = 0)"

            Me.TDGridPlanillaLeche.Columns(0).Caption = "NumPlanilla"
            Me.TDGridPlanillaLeche.Columns(1).Caption = "CodTipoNomina"
            Me.TDGridPlanillaLeche.Columns(2).Caption = "Fecha Inicial"
            Me.TDGridPlanillaLeche.Columns(3).Caption = "Fecha Final"
            Me.TDGridPlanillaLeche.Columns(4).Caption = "Año"
            Me.TDGridPlanillaLeche.Columns(5).Caption = "Mes"
            Me.TDGridPlanillaLeche.Columns(6).Caption = "Periodo"
            Me.TDGridPlanillaLeche.Columns(7).Caption = "Marca"
            
            
      Case "Pago Transportista"
        
            Me.TDGridPlanillaLeche.Columns(0).DataField = "NumPlanilla"
            Me.TDGridPlanillaLeche.Columns(1).DataField = "CodTipoNomina"
            Me.TDGridPlanillaLeche.Columns(2).DataField = "FechaInicial"
            Me.TDGridPlanillaLeche.Columns(3).DataField = "FechaFinal"
            Me.TDGridPlanillaLeche.Columns(4).DataField = "Año"
            Me.TDGridPlanillaLeche.Columns(5).DataField = "mes"
            Me.TDGridPlanillaLeche.Columns(6).DataField = "Periodo"
            Me.TDGridPlanillaLeche.Columns(7).DataField = "Marca"
            Me.TDGridPlanillaLeche.Splits(0).Caption = "Listado de Planillas Transportista"
            
            
            FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
            SqlString = "SELECT   NumPlanilla, CodTipoNomina, FechaInicial, FechaFinal, Año, mes, Periodo, Marca From NominaTransportista Where (Contabilizado = 0)"

            Me.TDGridPlanillaLeche.Columns(0).Caption = "NumPlanilla"
            Me.TDGridPlanillaLeche.Columns(1).Caption = "CodTipoNomina"
            Me.TDGridPlanillaLeche.Columns(2).Caption = "Fecha Inicial"
            Me.TDGridPlanillaLeche.Columns(3).Caption = "Fecha Final"
            Me.TDGridPlanillaLeche.Columns(4).Caption = "Año"
            Me.TDGridPlanillaLeche.Columns(5).Caption = "Mes"
            Me.TDGridPlanillaLeche.Columns(6).Caption = "Periodo"
            Me.TDGridPlanillaLeche.Columns(7).Caption = "Marca"
                        
                        
      Case "LiquidacionLeche"
        
            Me.TDGridPlanillaLeche.Columns(0).DataField = "NumeroLiquidacion"
            Me.TDGridPlanillaLeche.Columns(1).DataField = "FechaInicio"
            Me.TDGridPlanillaLeche.Columns(2).DataField = "FechaFin"
            Me.TDGridPlanillaLeche.Columns(3).DataField = "Nombre_Cliente"
            Me.TDGridPlanillaLeche.Columns(4).DataField = "Apellido_Cliente"
            Me.TDGridPlanillaLeche.Columns(5).DataField = "PrecioUnitario"
            Me.TDGridPlanillaLeche.Columns(6).DataField = "PorcientoIR"
            Me.TDGridPlanillaLeche.Columns(7).DataField = "Marca"
            Me.TDGridPlanillaLeche.Splits(0).Caption = "Listado de Planillas Transportista"
            
            
            FechaInicio = Format(Me.DTPicker11.Value, "yyyy-mm-dd")
            FechaFin = Format(Me.DTPicker12.Value, "yyyy-mm-dd")
            SqlString = "SELECT  LiquidacionLeche.NumeroLiquidacion, LiquidacionLeche.FechaInicio, LiquidacionLeche.FechaFin, Clientes.Nombre_Cliente, Clientes.Apellido_Cliente, LiquidacionLeche.Marca , LiquidacionLeche.Cod_Bodega, LiquidacionLeche.PorcientoIR, LiquidacionLeche.PrecioUnitario, LiquidacionLeche.Contabilizado FROM LiquidacionLeche INNER JOIN Clientes ON LiquidacionLeche.Cod_Proveedor = Clientes.Cod_Cliente Where (LiquidacionLeche.Contabilizado = 0)"

            Me.TDGridPlanillaLeche.Columns(0).Caption = "NumeroLiquidacion"
            Me.TDGridPlanillaLeche.Columns(1).Caption = "Fecha Inicio"
            Me.TDGridPlanillaLeche.Columns(2).Caption = "Fecha Fin"
            Me.TDGridPlanillaLeche.Columns(3).Caption = "Nombre Cliente"
            Me.TDGridPlanillaLeche.Columns(4).Caption = "Apellido Cliente"
            Me.TDGridPlanillaLeche.Columns(5).Caption = "PrecioUnitario"
            Me.TDGridPlanillaLeche.Columns(6).Caption = "PorcientoIR"
            Me.TDGridPlanillaLeche.Columns(7).Caption = "Marca"
      
      
      End Select
      
      
 Me.AdoRecepcion.RecordSource = SqlString
 Me.AdoRecepcion.Refresh
 TDGridPlanillaLeche.DataSource = Me.AdoRecepcion
 
 If Not Me.AdoRecepcion.Recordset.EOF Then
    Me.CmdContabilizarPlanilla.Enabled = True
    Me.CmdRecepcion.Enabled = True
    Me.LblFechaRecepcion.Visible = True
    Me.DTPicker10.Visible = True
    Me.CheckBox1.Visible = True
 Else
    Me.CmdContabilizarPlanilla.Enabled = False
    Me.LblFechaRecepcion.Visible = False
    Me.DTPicker10.Visible = False
    Me.CmdRecepcion.Enabled = False
    Me.CheckBox1.Visible = False
 End If

End Sub

Private Sub RadioButton1_Click()
'  If Me.RadioButton1.Value = True Then
'    Me.CmdRecepcion.Visible = True
'    Me.CmdContabilizarCompras.Visible = False
'  Else
'    Me.CmdContabilizarCompras.Visible = True
'    Me.CmdRecepcion.Visible = False
'  End If
End Sub

Private Sub RadioButton4_Click()
  If Me.RadioButton4.Value = True Then
    Me.ChkCheques.Visible = True
  Else
    Me.ChkCheques.Visible = False
  End If
End Sub

Private Sub RadioButton5_Click()

End Sub

