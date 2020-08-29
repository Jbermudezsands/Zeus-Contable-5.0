VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmExportacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportacion de Transacciones"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton CmdExportarAM 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar AM"
      UseVisualStyle  =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoMovimientos 
      Height          =   375
      Left            =   1320
      Top             =   8640
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoMovimientos"
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
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdConsultar 
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   1440
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Consultar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C1A1&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   13335
      TabIndex        =   8
      Top             =   -120
      Width           =   13335
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "FrmExportar.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   13320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Exportacion de Transacciones"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   4065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de Registros"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12975
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   495
         Left            =   7800
         TabIndex        =   13
         Top             =   120
         Width           =   4935
         _Version        =   786432
         _ExtentX        =   8705
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton OptPacioli 
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   150
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Pacioli 3000"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptZeus 
            Height          =   255
            Left            =   1800
            TabIndex        =   15
            Top             =   150
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Zeus Contabilidad"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptAM 
            Height          =   255
            Left            =   3840
            TabIndex        =   20
            Top             =   150
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "AM"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExportar.frx":1B42
         TabIndex        =   11
         Top             =   300
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   79626241
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   330
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   79626241
         CurrentDate     =   40457
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "FrmExportar.frx":1BB8
         TabIndex        =   12
         Top             =   300
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox TxtDebito 
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "##,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtCredito 
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "##,##0.00"
      PromptChar      =   "_"
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   11520
      OleObjectBlob   =   "FrmExportar.frx":1C28
      TabIndex        =   5
      Top             =   6960
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   9840
      OleObjectBlob   =   "FrmExportar.frx":1C94
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   6480
      Visible         =   0   'False
      Width           =   7935
      _Version        =   786432
      _ExtentX        =   13996
      _ExtentY        =   873
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmExportar.frx":1CFE
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CodCuentas"
      Columns(0).DataField=   "CodCuentas"
      Columns(0).DataWidth=   50
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NombreCuenta"
      Columns(1).DataField=   "NombreCuenta"
      Columns(1).DataWidth=   255
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NumeroMovimiento"
      Columns(2).DataField=   "NumeroMovimiento"
      Columns(2).DataWidth=   11
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "VoucherNo"
      Columns(3).DataField=   "VoucherNo"
      Columns(3).DataWidth=   50
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DescripcionMovimiento"
      Columns(4).DataField=   "DescripcionMovimiento"
      Columns(4).DataWidth=   255
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Factura No."
      Columns(5).DataField=   "FacturaNo"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Cheque No."
      Columns(6).DataField=   "ChequeNo"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Clave"
      Columns(7).DataField=   "Clave"
      Columns(7).DataWidth=   10
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "TCambio"
      Columns(8).DataField=   "TCambio"
      Columns(8).DataWidth=   23
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Debito"
      Columns(9).DataField=   "Debito"
      Columns(9).DataWidth=   22
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Credito"
      Columns(10).DataField=   "Credito"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "FechaTransaccion"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "NPeriodo"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "FechaDescuento"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "DescuentoDisponible"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "FechaVence"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Beneficiario"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "CodCuentaProveedor"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "TipoFactura"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "NTransaccion"
      Columns(19).DataField=   "Credito"
      Columns(19).DataWidth=   22
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   20
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Movimientos de Indices"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=20"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3254"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131588"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3069"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2990"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131588"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(2)._AlignLeft=0"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=3254"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=131588"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=3704"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3625"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=131588"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=131588"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=131588"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=1667"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1588"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=131588"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(42)=   "Column(8).Width=3254"
      Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(45)=   "Column(8)._ColStyle=131588"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(8)._AlignLeft=0"
      Splits(0)._ColumnProps(48)=   "Column(9).Width=3254"
      Splits(0)._ColumnProps(49)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(9)._WidthInPix=3175"
      Splits(0)._ColumnProps(51)=   "Column(9)._ColStyle=131588"
      Splits(0)._ColumnProps(52)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(53)=   "Column(9)._AlignLeft=0"
      Splits(0)._ColumnProps(54)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=131588"
      Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(59)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(11)._ColStyle=131588"
      Splits(0)._ColumnProps(63)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(64)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(65)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(12)._ColStyle=131588"
      Splits(0)._ColumnProps(69)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(71)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(72)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(74)=   "Column(13)._ColStyle=131588"
      Splits(0)._ColumnProps(75)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(76)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(77)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(78)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(80)=   "Column(14)._ColStyle=131588"
      Splits(0)._ColumnProps(81)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(83)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(15)._ColStyle=131588"
      Splits(0)._ColumnProps(87)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(89)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(90)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(92)=   "Column(16)._ColStyle=131588"
      Splits(0)._ColumnProps(93)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(95)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(96)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(98)=   "Column(17)._ColStyle=131588"
      Splits(0)._ColumnProps(99)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(100)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(101)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(102)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(104)=   "Column(18)._ColStyle=131588"
      Splits(0)._ColumnProps(105)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(107)=   "Column(19).Width=3254"
      Splits(0)._ColumnProps(108)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(19)._WidthInPix=3175"
      Splits(0)._ColumnProps(110)=   "Column(19)._ColStyle=131588"
      Splits(0)._ColumnProps(111)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(112)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(113)=   "Column(19)._AlignLeft=0"
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
      PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureCurrentRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
      PictureCurrentRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAP//"
      PictureCurrentRow(2)=   "/////////////////////////////////////////8bHxgAAxsfGAAAAhIaExsfGxsfGxsfGxsfG"
      PictureCurrentRow(3)=   "xsfGxsfGxsfGxsfGxsfG////xsfGAADGx8YAAACEhoTGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      PictureCurrentRow(4)=   "x8b////Gx8YAAMbHxgAAAISGhMbHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxv///8bHxgAAxsfG"
      PictureCurrentRow(5)=   "AAAAhIaExsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfG////xsfGAADGx8YAAACEhoTGx8bGx8bG"
      PictureCurrentRow(6)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8b////Gx8YAAMbHxgAAAISGhISGhISGhISGhISGhISGhISGhISG"
      PictureCurrentRow(7)=   "hISGhISGhP///8bHxgAAxsfGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxsfG"
      PictureCurrentRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFFAEFF&,.fgcolor=&H800080&"
      _StyleDefs(20)  =   ":id=22,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(21)  =   ":id=22,.fontname=Lucida Calligraphy"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&HECB877&"
      _StyleDefs(23)  =   ":id=14,.fgcolor=&H800000&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(24)  =   ":id=14,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=2,.bgcolor=&HFF0000&"
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
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(72)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(73)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(75)  =   "Splits(0).Columns(10).Style:id=86,.parent=13"
      _StyleDefs(76)  =   "Splits(0).Columns(10).HeadingStyle:id=83,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(10).FooterStyle:id=84,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(10).EditorStyle:id=85,.parent=17"
      _StyleDefs(79)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(80)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(81)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(83)  =   "Splits(0).Columns(12).Style:id=90,.parent=13"
      _StyleDefs(84)  =   "Splits(0).Columns(12).HeadingStyle:id=87,.parent=14"
      _StyleDefs(85)  =   "Splits(0).Columns(12).FooterStyle:id=88,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(12).EditorStyle:id=89,.parent=17"
      _StyleDefs(87)  =   "Splits(0).Columns(13).Style:id=94,.parent=13"
      _StyleDefs(88)  =   "Splits(0).Columns(13).HeadingStyle:id=91,.parent=14"
      _StyleDefs(89)  =   "Splits(0).Columns(13).FooterStyle:id=92,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(13).EditorStyle:id=93,.parent=17"
      _StyleDefs(91)  =   "Splits(0).Columns(14).Style:id=98,.parent=13"
      _StyleDefs(92)  =   "Splits(0).Columns(14).HeadingStyle:id=95,.parent=14"
      _StyleDefs(93)  =   "Splits(0).Columns(14).FooterStyle:id=96,.parent=15"
      _StyleDefs(94)  =   "Splits(0).Columns(14).EditorStyle:id=97,.parent=17"
      _StyleDefs(95)  =   "Splits(0).Columns(15).Style:id=102,.parent=13"
      _StyleDefs(96)  =   "Splits(0).Columns(15).HeadingStyle:id=99,.parent=14"
      _StyleDefs(97)  =   "Splits(0).Columns(15).FooterStyle:id=100,.parent=15"
      _StyleDefs(98)  =   "Splits(0).Columns(15).EditorStyle:id=101,.parent=17"
      _StyleDefs(99)  =   "Splits(0).Columns(16).Style:id=106,.parent=13"
      _StyleDefs(100) =   "Splits(0).Columns(16).HeadingStyle:id=103,.parent=14"
      _StyleDefs(101) =   "Splits(0).Columns(16).FooterStyle:id=104,.parent=15"
      _StyleDefs(102) =   "Splits(0).Columns(16).EditorStyle:id=105,.parent=17"
      _StyleDefs(103) =   "Splits(0).Columns(17).Style:id=110,.parent=13"
      _StyleDefs(104) =   "Splits(0).Columns(17).HeadingStyle:id=107,.parent=14"
      _StyleDefs(105) =   "Splits(0).Columns(17).FooterStyle:id=108,.parent=15"
      _StyleDefs(106) =   "Splits(0).Columns(17).EditorStyle:id=109,.parent=17"
      _StyleDefs(107) =   "Splits(0).Columns(18).Style:id=114,.parent=13"
      _StyleDefs(108) =   "Splits(0).Columns(18).HeadingStyle:id=111,.parent=14"
      _StyleDefs(109) =   "Splits(0).Columns(18).FooterStyle:id=112,.parent=15"
      _StyleDefs(110) =   "Splits(0).Columns(18).EditorStyle:id=113,.parent=17"
      _StyleDefs(111) =   "Splits(0).Columns(19).Style:id=78,.parent=13"
      _StyleDefs(112) =   "Splits(0).Columns(19).HeadingStyle:id=75,.parent=14"
      _StyleDefs(113) =   "Splits(0).Columns(19).FooterStyle:id=76,.parent=15"
      _StyleDefs(114) =   "Splits(0).Columns(19).EditorStyle:id=77,.parent=17"
      _StyleDefs(115) =   "Named:id=33:Normal"
      _StyleDefs(116) =   ":id=33,.parent=0"
      _StyleDefs(117) =   "Named:id=34:Heading"
      _StyleDefs(118) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(119) =   ":id=34,.wraptext=-1"
      _StyleDefs(120) =   "Named:id=35:Footing"
      _StyleDefs(121) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(122) =   "Named:id=36:Selected"
      _StyleDefs(123) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(124) =   "Named:id=37:Caption"
      _StyleDefs(125) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(126) =   "Named:id=38:HighlightRow"
      _StyleDefs(127) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(128) =   "Named:id=39:EvenRow"
      _StyleDefs(129) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(130) =   "Named:id=40:OddRow"
      _StyleDefs(131) =   ":id=40,.parent=33"
      _StyleDefs(132) =   "Named:id=41:RecordSelector"
      _StyleDefs(133) =   ":id=41,.parent=34"
      _StyleDefs(134) =   "Named:id=42:FilterBar"
      _StyleDefs(135) =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   1320
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
   Begin MSAdodcLib.Adodc AdoTransacciones 
      Height          =   375
      Left            =   4680
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
      RecordSource    =   ""
      Caption         =   "AdoTransacciones"
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
   Begin MSAdodcLib.Adodc AdoRegistros 
      Height          =   375
      Left            =   8160
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
      RecordSource    =   ""
      Caption         =   "AdoRegistros"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "Cns"
      FileName        =   "*.Cns"
      Filter          =   "Cns"
   End
   Begin XtremeSuiteControls.PushButton CmdExportarPacioli 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExportarZeus 
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   7080
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar Zeus"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "FrmExportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 

End Sub

Private Sub CmdConsultar_Click()
Dim R As Variant
Dim CodigoCuenta As String, TotalDebio As Double, TotalCredito As Double, Debito As Double, Credito As Double
Dim Registros As Double, rs As New ADODB.Recordset

'////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////BUSCO EL PERIODO DE LA TRANSACCION ////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
 mes = Month(Me.DTPFechaFin.Value)
 Año = Year(Me.DTPFechaFin.Value)
 FechaIni = CDate("1/" & Month(Me.DTPFechaFin.Value) & "/" & Year(Me.DTPFechaFin.Value))
 FechaFin = DateSerial(Año, mes + 1, 1 - 1)
 NumFecha1 = FechaIni
 NumFecha2 = FechaFin
 Fechas1 = Format(FechaIni, "yyyy/mm/dd")
 Fechas2 = Format(FechaFin, "yyyy/mm/dd")
 
 DoEvents
 
''  Me.AdoConsulta.RecordSource = "SELECT  * From Transacciones WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
'  Me.AdoConsulta.RecordSource = "SELECT  CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, ROUND(TCambio, 2)* Debito AS Debito, ROUND(TCambio, 2)*Credito AS Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada , FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura  From Transacciones  " & _
'                                "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
'  Me.AdoConsulta.Refresh
'
'
'' Me.AdoTransacciones.RecordSource = "SELECT * From Transacciones WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
' Me.AdoTransacciones.RecordSource = "SELECT  CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, ROUND(TCambio, 2)*Debito AS Debito, ROUND(TCambio, 2)*Credito AS Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada , FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura  From Transacciones  " & _
'                                    "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
' Me.AdoTransacciones.Refresh
'
'
' Me.AdoConsulta.RecordSource = "SELECT MAX(CodCuentas) AS CodCuentas, SUM(TCambio) AS TCambio, SUM(Debito * ROUND(TCambio, 2)) AS Debito, SUM(Credito * ROUND(TCambio, 2)) AS Credito From Transacciones  " & _
'                               "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102))"
' Me.AdoConsulta.Refresh

 '**************************************************************************************************************************************************************************************************************************************************************************************************************************************************
 '//////////////////////////////////////////BORRO LOS REGISTROS /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 '**************************************************************************************************************************************************************************************************************************************************************************************************************************************************
 rs.Open "DELETE FROM [Registros]", Conexion
 
 Me.AdoRegistros.RecordSource = "SELECT  * From Registros"
 Me.AdoRegistros.Refresh
 
 
 Me.AdoMovimientos.RecordSource = "SELECT  CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, TCambio*Debito AS Debito, TCambio*Credito AS Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada , FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura  From Transacciones  " & _
                                    "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
 Me.AdoMovimientos.Refresh
 If Not Me.AdoMovimientos.Recordset.EOF Then
   Me.AdoMovimientos.Recordset.MoveLast
 End If
 
 Registros = Me.AdoMovimientos.Recordset.RecordCount
 Barra.Visible = True
 Barra.Min = 0
 Barra.Value = 0
 Barra.Max = Registros

 If Not Me.AdoMovimientos.Recordset.BOF Then
   Me.AdoMovimientos.Recordset.MoveFirst
 End If
 TotalDebito = 0
 TotalCredito = 0
 Do While Not Me.AdoMovimientos.Recordset.EOF
    '******************************************************************************************************************************************************************************************************************************************
    '///////////////////////////////////////////////////SUMO CADA MOVIMIENTO Y TRUNCO LOS DECIMALES //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '******************************************************************************************************************************************************************************************************************************************
  
  DoEvents
      CodigoCuenta = Me.AdoMovimientos.Recordset("CodCuentas")
      
      Debito = Me.AdoMovimientos.Recordset("Debito")
      Debito = TRUNC(Debito, 3)
      Debito = Format(Debito, "##,##0.00")
      TotalDebito = Debito + TotalDebito

      Credito = Me.AdoMovimientos.Recordset("Credito")
      Credito = TRUNC(Credito, 3)
      Credito = Format(Credito, "##,##0.00")
      TotalCredito = Credito + TotalCredito

                      
                      
                      
    '******************************************************************************************************************************************************************************************************************************************
    '///////////////////////////////////////////////////AGRO EL TOTAL DE CADA MOVIMIENTO //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '******************************************************************************************************************************************************************************************************************************************
      Me.AdoRegistros.Recordset.AddNew
      Me.AdoRegistros.Recordset("IdRegistros") = 1
      Me.AdoRegistros.Recordset("Fecha") = Me.AdoMovimientos.Recordset("FechaTransaccion")
      Me.AdoRegistros.Recordset("NTransaccion") = Me.AdoMovimientos.Recordset("NumeroMovimiento")
      Me.AdoRegistros.Recordset("Fuente") = Me.AdoMovimientos.Recordset("Fuente")
      Me.AdoRegistros.Recordset("CodCuenta") = Me.AdoMovimientos.Recordset("CodCuentas")
      Me.AdoRegistros.Recordset("FacturaNumero") = Me.AdoMovimientos.Recordset("FacturaNo")
      Me.AdoRegistros.Recordset("TipoMovimiento") = Me.AdoMovimientos.Recordset("Clave")
      Me.AdoRegistros.Recordset("RefCheque") = Me.AdoMovimientos.Recordset("ChequeNo")
      Me.AdoRegistros.Recordset("Descripcion") = Me.AdoMovimientos.Recordset("DescripcionMovimiento")
      Me.AdoRegistros.Recordset("VoucherNo") = Me.AdoMovimientos.Recordset("VoucherNo")
      Me.AdoRegistros.Recordset("TCambio") = Me.AdoMovimientos.Recordset("TCambio")
      Me.AdoRegistros.Recordset("ImporteTransaccionDebito") = Debito
      Me.AdoRegistros.Recordset("ImporteTransaccionCredito") = Credito
      Me.AdoRegistros.Recordset.Update
                      
                      
      Me.AdoMovimientos.Recordset.MoveNext
      Barra.Value = Barra.Value + 1
  Loop


   Me.AdoTransacciones.RecordSource = "SELECT  Cuentas.CodCuentas, Cuentas.DescripcionCuentas AS NombreCuenta, Registros.NTransaccion AS NumeroMovimiento, Registros.Descripcion AS DescripcionMovimiento, Registros.FacturaNumero AS FacturaNo, Registros.RefCheque AS ChequeNo, Registros.TipoMovimiento AS Clave, Registros.ImporteTransaccionDebito AS Debito, Registros.ImporteTransaccionCredito AS Credito, Registros.VoucherNo , Registros.TCambio FROM  Registros INNER JOIN  Cuentas ON Registros.CodCuenta = Cuentas.CodCuentas ORDER BY NumeroMovimiento"
   Me.AdoTransacciones.Refresh
   
   Me.TxtDebito.Text = Format(TotalDebito, "##,##0.00")
   Me.TxtCredito.Text = Format(TotalCredito, "##,##0.00")
 

End Sub

Private Sub CmdExportarAM_Click()
Dim V As Double, H As Double, i As Double, Directorio As String
Dim Fecha As Date, NumeroTransaccion As Double, Fuente As String
Dim oExcel As Object, NumeroAnterior As Double, Cadena As String
Dim oBook As Object, objExcel As Object
'Dim oSheet As Object

Me.CommonDialog1.Filter = "Archivos de Excel|*.xls"
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName

 Exportar = True
 
'   Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
   'Inicio el Nuevo LibroExcel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set objExcel = oBook.Worksheets(1)
        V = 1
        H = 0
        i = 1
        NumeroAnterior = 0
        NumeroTransaccion = 0
        Cadena = "FIN_PARTIDAS"
     Me.AdoRegistros.RecordSource = "SELECT  Cuentas.CodCuentas, Cuentas.DescripcionCuentas AS NombreCuenta, Registros.NTransaccion AS NumeroMovimiento, Registros.Descripcion AS DescripcionMovimiento, Registros.FacturaNumero AS FacturaNo, Registros.RefCheque AS ChequeNo, Registros.TipoMovimiento AS Clave, Registros.ImporteTransaccionDebito AS Debito, Registros.ImporteTransaccionCredito AS Credito, Registros.VoucherNo , Registros.TCambio,Registros.Fecha AS Fecha,Registros.Fuente AS Fuente FROM  Registros INNER JOIN  Cuentas ON Registros.CodCuenta = Cuentas.CodCuentas ORDER BY NumeroMovimiento"
     Me.AdoRegistros.Refresh
     Do While Not Me.AdoRegistros.Recordset.EOF
     
            Fecha = Me.AdoRegistros.Recordset("Fecha")
            
           
            NumeroTransaccion = Me.AdoRegistros.Recordset("NumeroMovimiento")
            Fuente = Me.AdoRegistros.Recordset("Fuente")
            If Fuente = "CHEQUE" Then
             Fuente = "Eg"
            Else
             Fuente = "Otros"
            End If
            
            If NumeroAnterior <> NumeroTransaccion Then
                '---------------------------BUSCO EL ENCABEZADO DE LA TRANSACCION --------------------------------
                Me.AdoConsulta.RecordSource = "SELECT  * From IndiceTransaccion WHERE (FechaTransaccion = CONVERT(DATETIME, '" & Format(Fecha, "yyyy-mm-dd") & "', 102)) AND (NumeroMovimiento = " & NumeroTransaccion & ")"
                Me.AdoConsulta.Refresh
                If Not Me.AdoConsulta.Recordset.EOF Then
                  objExcel.Cells(V, H + 1) = Fuente
                  objExcel.Cells(V, H + 2) = NumeroTransaccion
                  objExcel.Cells(V, H + 3) = Me.AdoRegistros.Recordset("DescripcionMovimiento")
                  objExcel.Cells(V, H + 4) = Day(Fecha)
                 V = V + 1
                 i = i + 1
                End If
            End If
            
            objExcel.Cells(V, H + 2) = Me.AdoRegistros.Recordset("CodCuentas")
            objExcel.Cells(V, H + 3) = 0
            objExcel.Cells(V, H + 4) = Me.AdoRegistros.Recordset("DescripcionMovimiento")
            objExcel.Cells(V, H + 5) = 1
            objExcel.Cells(V, H + 6) = Me.AdoRegistros.Recordset("Debito")
            objExcel.Cells(V, H + 7) = Me.AdoRegistros.Recordset("Credito")
            

            V = V + 1
            i = i + 1
            
            NumeroAnterior = NumeroTransaccion
            
        Me.AdoRegistros.Recordset.MoveNext
         If Not Me.AdoRegistros.Recordset.EOF Then
           If NumeroAnterior <> Me.AdoRegistros.Recordset("NumeroMovimiento") Then
              objExcel.Cells(V, H + 2) = "FIN_PARTIDAS"
              V = V + 1
              i = i + 1
           End If
         Else
              objExcel.Cells(V, H + 2) = "FIN_PARTIDAS"
              V = V + 1
              i = i + 1
         End If
     Loop

   'Salvar Excel
   oBook.SaveAs Directorio
   oExcel.Quit
   MsgBox "Proceso Terminado!!!", vbInformation, "Zeus Contable"
End Sub

Private Sub CmdExportarPacioli_Click()
'On Error GoTo TipoErrs
Dim SQLExporta As String, Longitud As Integer, Respuesta As Integer
Dim Cadena As String, mes As String, Dia As String, ano As String
Dim TextoMonto As String, TipoMovimiento As String, J As Integer
Dim Consecutivo As Double, FechaDescuento As String, FechaVencimiento As String


 Me.AdoRegistros.RecordSource = "SELECT CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, ROUND(TCambio, 2) * Debito AS Debito, ROUND(TCambio, 2) * Credito AS Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada , FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura From Transacciones  " & _
                                "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
 Me.AdoRegistros.Refresh
Salir = False
Barra.Visible = True
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName
AdoRegistros.Recordset.MoveLast
Maximo = AdoRegistros.Recordset.RecordCount
If (Dir(Directorio) <> "") Then
  Respuesta = MsgBox("Reescribir el Archivo?", vbYesNo, "Zeus Contabilidad")
  If Respuesta = 6 Then
               
               Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoRegistros.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   J = 0
                 Do While Not AdoRegistros.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                 If Me.AdoRegistros.Recordset("Clave") = "Debito" Then
                 TipoMovimiento = "01"
                 
                 Else
                 TipoMovimiento = "05"
                 End If
                 
                 Fuente = Mid(AdoRegistros.Recordset("Fuente"), 1, 4)
                 For i = 1 To 4 - Len(Fuente)
                  Fuente = Fuente & " "
                 Next i
                 
                 
                 Consecutivo = Cosecutivo + 1
                 Fecha = AdoRegistros.Recordset("FechaTransaccion")
                 NTransaccion = Format(AdoRegistros.Recordset("NumeroMovimiento"), "0000000#")
                 
                 '////////////////////////////////////////////////////////////////////////////////////
                 '///////////////////////////BUSCO SI EXITE CUENTA DE EXPORTACION//////////////////////
                 '///////////////////////////////////////////////////////////////////////////////////
                 CodCuenta = AdoRegistros.Recordset("CodCuentas")
                 Me.AdoConsulta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo, UbicacionReporte, SubDivicion, CausaIva , CausaRetencion, DescRetencion, Nombre1, Nombre2, Apellido1, Apellido2, Cedula, RUC, Telefono, Direccion, CodCuentaImporta From Cuentas " & _
                                               "WHERE     (CodCuentas = '" & CodCuenta & "')  "
                 Me.AdoConsulta.Refresh
                 If Not Me.AdoConsulta.Recordset.EOF Then
                   If Not IsNull(Me.AdoConsulta.Recordset("CodCuentaImporta")) Then
                     CodCuenta = AdoConsulta.Recordset("CodCuentaImporta")
                   End If
                 End If
                                               
                 
                 CodDepartamento = "   "
                 CodAcciones = "                 "
                 ClaveProyecto = "               "
                 
                 If AdoRegistros.Recordset("FacturaNo") <> "-" Then
                    NFactura = AdoRegistros.Recordset("FacturaNo")
                 Else
                    NFactura = ""
                 End If
                 
                 For i = 1 To 8 - Len(NFactura)
                  NFactura = NFactura & " "
                 Next i
                  NFactura = Mid(NFactura, 1, 8)
                 
                 ReferenciaCh = Mid(AdoRegistros.Recordset("ChequeNo"), 1, 6)
                 For i = 1 To 6 - Len(ReferenciaCh)
                  ReferenciaCh = ReferenciaCh & " "
                 Next i
                  ReferenciaCh = Mid(ReferenciaCh, 1, 6)
                 
                 Descripcion = Mid(AdoRegistros.Recordset("DescripcionMovimiento"), 1, 35)
                 For i = 1 To 35 - Len(Descripcion)
                  Descripcion = Descripcion & " "
                 Next i
                  Descripcion = Mid(Descripcion, 1, 35)
                 
                 FechaDescuento = Format(AdoRegistros.Recordset("FechaDescuento"), "MMDDYYYY")
                 FechaVencimiento = Format(AdoRegistros.Recordset("FechaVence"), "MMDDYYYY")
                 ImporteDescuento = Format(AdoRegistros.Recordset("DescuentoDisponible"), "####0.00")
                 ValorUnit = 0
                 TipoTransaccion = "00"
                           
                 '/////////Verifico el tipo de movimiento//////////////
                   If TipoMovimiento = "01" Or TipoMovimiento = "02" Or TipoMovimiento = "03" Or TipoMovimiento = "04" Or TipoMovimiento = "10" Or TipoMovimiento = "11" Or TipoMovimiento = "12" Or TipoMovimiento = "13" Or TipoMovimiento = "14" Or TipoMovimiento = "20" Or TipoMovimiento = "21" Or TipoMovimiento = "27" Or TipoMovimiento = "31" Then
                     TextoMonto = Format(AdoRegistros.Recordset("Debito"), "####0.0000")
                   ElseIf TipoMovimiento = "05" Or TipoMovimiento = "06" Or TipoMovimiento = "07" Or TipoMovimiento = "08" Or TipoMovimiento = "09" Or TipoMovimiento = "15" Or TipoMovimiento = "16" Or TipoMovimiento = "17" Or TipoMovimiento = "18" Or TipoMovimiento = "19" Or TipoMovimiento = "22" Or TipoMovimiento = "28" Or TipoMovimiento = "29" Or TipoMovimiento = "30" Then
                      TextoMonto = Format(AdoRegistros.Recordset("Credito"), "####0.0000")
                    End If
                    
                    For i = 1 To 18 - Len(TextoMonto)
                       TextoMonto = " " & TextoMonto
                    Next i
                    
                    
                    mes = Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(mes)
                    If Longitud = 1 Then
                     mes = "0" & Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    
                    Dia = Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(Dia)
                    If Longitud = 1 Then
                     Dia = "0" & Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    ano = Cadena & Trim(Str(Year(AdoRegistros.Recordset("FechaTransaccion"))))
                    Cadena = mes & Dia & ano
                    Cadena = Cadena & NTransaccion
                    Cadena = Cadena & Fuente
                    Cadena = Cadena & Trim(Str(CodCuenta))
                    For i = 1 To 36 - Len(Cadena)
                    Cadena = Cadena & " "
                    Next i
                    Cadena = Cadena & CodDepartamento & CodAcciones & ClaveProyecto & NFactura + TipoMovimiento + ReferenciaCh
                    Cadena = Cadena & Descripcion & FechaDescuento & FechaVencimiento
                    Cadena = Cadena & TextoMonto
                    
                    For i = 1 To 17 - Len(ImporteDescuento)
                       ImporteDescuento = " " & ImporteDescuento
                    Next i
                    For i = 1 To 17 - Len(ValorUnit)
                       ValorUnit = " " & ValorUnit
                    Next i
                    Cadena = Cadena & ImporteDescuento & ValorUnit
                    Cadena = Cadena & TipoTransaccion
                    Print #1, Cadena
                                    
                    
                    
                  AdoRegistros.Recordset.MoveNext
                  J = J + 1
                  Me.Caption = "Procesando:  " & J & " de " & Maximo & " Registros "
                  DoEvents
                  .Value = J
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Enlace"
                Salir = True
  End If
Else '//////En caso que no exista el Archivo///////////
                
                Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoRegistros.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   J = 0
                 Do While Not AdoRegistros.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                 
                 If Me.AdoRegistros.Recordset("Clave") = "Debito" Then
                    TipoMovimiento = "01"
                 Else
                    TipoMovimiento = "05"
                 End If
                 
                 Fuente = Mid(AdoRegistros.Recordset("Fuente"), 1, 4)
                 For i = 1 To 4 - Len(Fuente)
                  Fuente = Fuente & " "
                 Next i
                 
                 
                 Consecutivo = Cosecutivo + 1
                 Fecha = AdoRegistros.Recordset("FechaTransaccion")
                 NTransaccion = Format(AdoRegistros.Recordset("NumeroMovimiento"), "0000000#")
                 
                 '////////////////////////////////////////////////////////////////////////////////////
                 '///////////////////////////BUSCO SI EXITE CUENTA DE EXPORTACION//////////////////////
                 '///////////////////////////////////////////////////////////////////////////////////
                 CodCuenta = AdoRegistros.Recordset("CodCuentas")
                 Me.AdoConsulta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo, UbicacionReporte, SubDivicion, CausaIva , CausaRetencion, DescRetencion, Nombre1, Nombre2, Apellido1, Apellido2, Cedula, RUC, Telefono, Direccion, CodCuentaImporta From Cuentas " & _
                                               "WHERE     (CodCuentas = '" & CodCuenta & "')  "
                 Me.AdoConsulta.Refresh
                 If Not Me.AdoConsulta.Recordset.EOF Then
                   If Not IsNull(Me.AdoConsulta.Recordset("CodCuentaImporta")) Then
                     CodCuenta = AdoConsulta.Recordset("CodCuentaImporta")
                   End If
                 End If
                 
                 CodDepartamento = "   "
                 CodAcciones = "                 "
                 ClaveProyecto = "               "
                 
                 If AdoRegistros.Recordset("FacturaNo") <> "-" Then
                    NFactura = AdoRegistros.Recordset("FacturaNo")
                 Else
                    NFactura = ""
                 End If
                 
                 NFactura = Mid(NFactura, 1, 8)
                 For i = 1 To 8 - Len(NFactura)
                  NFactura = NFactura & " "
                 Next i
                 
                 
                 If Not IsNull(AdoRegistros.Recordset("ChequeNo")) Then
                    ReferenciaCh = AdoRegistros.Recordset("ChequeNo")
                    For i = 1 To 6 - Len(ReferenciaCh)
                     ReferenciaCh = ReferenciaCh & " "
                    Next i
                    ReferenciaCh = Mid(ReferenciaCh, 1, 6)
                 End If
                 
                 Descripcion = Mid(AdoRegistros.Recordset("DescripcionMovimiento"), 1, 35)
                 For i = 1 To 35 - Len(Descripcion)
                  Descripcion = Descripcion & " "
                 Next i
                 Descripcion = Mid(Descripcion, 1, 35)
                 
                 FechaDescuento = Format(AdoRegistros.Recordset("FechaDescuento"), "MMDDYYYY")
                 FechaVencimiento = Format(AdoRegistros.Recordset("FechaVence"), "MMDDYYYY")
                 ImporteDescuento = Format(AdoRegistros.Recordset("DescuentoDisponible"), "####0.00")
                 ValorUnit = 0
                 TipoTransaccion = "00"
                           
                 '/////////Verifico el tipo de movimiento//////////////
                   If TipoMovimiento = "01" Or TipoMovimiento = "02" Or TipoMovimiento = "03" Or TipoMovimiento = "04" Or TipoMovimiento = "10" Or TipoMovimiento = "11" Or TipoMovimiento = "12" Or TipoMovimiento = "13" Or TipoMovimiento = "14" Or TipoMovimiento = "20" Or TipoMovimiento = "21" Or TipoMovimiento = "27" Or TipoMovimiento = "31" Then
                     TextoMonto = Format(AdoRegistros.Recordset("Debito"), "####0.0000")
                   ElseIf TipoMovimiento = "05" Or TipoMovimiento = "06" Or TipoMovimiento = "07" Or TipoMovimiento = "08" Or TipoMovimiento = "09" Or TipoMovimiento = "15" Or TipoMovimiento = "16" Or TipoMovimiento = "17" Or TipoMovimiento = "18" Or TipoMovimiento = "19" Or TipoMovimiento = "22" Or TipoMovimiento = "28" Or TipoMovimiento = "29" Or TipoMovimiento = "30" Then
                      TextoMonto = Format(AdoRegistros.Recordset("Credito"), "####0.0000")
                    End If
                    
                    For i = 1 To 18 - Len(TextoMonto)
                       TextoMonto = " " & TextoMonto
                    Next i
                    
                    
                    mes = Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(mes)
                    If Longitud = 1 Then
                     mes = "0" & Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    
                    Dia = Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(Dia)
                    If Longitud = 1 Then
                     Dia = "0" & Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    ano = Cadena & Trim(Str(Year(AdoRegistros.Recordset("FechaTransaccion"))))
                    Cadena = mes & Dia & ano
                    Cadena = Cadena & NTransaccion
                    Cadena = Cadena & Fuente
                    Cadena = Cadena & Trim(CodCuenta)
                    For i = 1 To 36 - Len(Cadena)
                    Cadena = Cadena & " "
                    Next i
                    Cadena = Cadena & CodDepartamento & CodAcciones & ClaveProyecto & NFactura & TipoMovimiento & ReferenciaCh
                    Cadena = Cadena & Descripcion & FechaDescuento & FechaVencimiento
                    Cadena = Cadena & TextoMonto
                    
                    For i = 1 To 17 - Len(ImporteDescuento)
                       ImporteDescuento = " " & ImporteDescuento
                    Next i
                    For i = 1 To 17 - Len(ValorUnit)
                       ValorUnit = " " & ValorUnit
                    Next i
                    Cadena = Cadena & ImporteDescuento & ValorUnit
                    Cadena = Cadena & TipoTransaccion
                    Print #1, Cadena
                                    
                    
                    
                  AdoRegistros.Recordset.MoveNext
                  J = J + 1
                  .Value = J
                  Me.Caption = "Procesando:  " & J & " de " & Maximo & " Registros "
                  DoEvents
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Zeus Facturacion"



Salir = True
End If
Exit Sub
TipoErrs:
MsgBox err.Description
Salir = True
End Sub

Private Sub CmdExportarZeus_Click()
'On Error GoTo TipoErrs
Dim SQLExporta As String, Longitud As Integer, Respuesta As Integer
Dim Cadena As String, mes As String, Dia As String, ano As String
Dim TextoMonto As String, TipoMovimiento As String, J As Integer
Dim Consecutivo As Double, FechaDescuento As String, FechaVencimiento As String, VoucherNo As String


' Me.AdoRegistros.RecordSource = "SELECT  CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, ROUND(Debito * TCambio, 2) AS Debito, ROUND(Credito * TCambio, 2) AS Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada , FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura  From Transacciones  " & _
'                                "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
 
' Me.AdoRegistros.RecordSource = "SELECT CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, ROUND(TCambio, 2) * Debito AS Debito, ROUND(TCambio, 2) * Credito AS Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada , FechaDescuento, DescuentoDisponible, FechaVence, Beneficiario, CodCuentaProveedor, TipoFactura From Transacciones  " & _
'                                "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Format(Me.DTPFechaIni.Value, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTPFechaFin.Value, "yyyy-MM-dd") & "', 102)) ORDER BY NTransaccion"
 
   
 Me.AdoRegistros.RecordSource = "SELECT  Cuentas.CodCuentas, Cuentas.DescripcionCuentas AS NombreCuenta, Registros.NTransaccion, Registros.Descripcion AS DescripcionMovimiento, Registros.FacturaNumero AS FacturaNo, Registros.RefCheque AS ChequeNo, Registros.TipoMovimiento AS Clave, Registros.ImporteTransaccionDebito AS Debito, Registros.ImporteTransaccionCredito AS Credito, Registros.VoucherNo , Registros.TCambio, Registros.Fuente, Registros.Fecha AS FechaTransaccion, Registros.NTransaccion AS NumeroMovimiento, Registros.FechaDescuento, Registros.FechaVencimiento As FechaVence, Registros.ImporteDescuento As DescuentoDisponible, Registros.ValorUnitario, Registros.TipoTransaccion, Registros.CreditoDolar, Registros.DebitoDolar, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto FROM  Registros INNER JOIN  Cuentas ON Registros.CodCuenta = Cuentas.CodCuentas ORDER BY NumeroMovimiento  "
 Me.AdoRegistros.Refresh
Salir = False
Barra.Visible = True
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName
AdoRegistros.Recordset.MoveLast
Maximo = AdoRegistros.Recordset.RecordCount
If (Dir(Directorio) <> "") Then
  Respuesta = MsgBox("Reescribir el Archivo?", vbYesNo, "Zeus Contabilidad")
  If Respuesta = 6 Then
               
               Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoRegistros.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   J = 0
                 Do While Not AdoRegistros.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                 If Me.AdoRegistros.Recordset("Clave") = "Debito" Then
                 TipoMovimiento = "01"
                 
                 Else
                 TipoMovimiento = "05"
                 End If
                 
                 Fuente = Mid(AdoRegistros.Recordset("Fuente"), 1, 4)
                 For i = 1 To 4 - Len(Fuente)
                  Fuente = Fuente & " "
                 Next i
                 
                 
                 Consecutivo = Cosecutivo + 1
                 Fecha = AdoRegistros.Recordset("FechaTransaccion")
                 NTransaccion = Format(AdoRegistros.Recordset("NumeroMovimiento"), "0000000#")

                 '////////////////////////////////////////////////////////////////////////////////////
                 '///////////////////////////BUSCO SI EXITE CUENTA DE EXPORTACION//////////////////////
                 '///////////////////////////////////////////////////////////////////////////////////
                 CodCuenta = AdoRegistros.Recordset("CodCuentas")


                 Me.AdoConsulta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo, UbicacionReporte, SubDivicion, CausaIva , CausaRetencion, DescRetencion, Nombre1, Nombre2, Apellido1, Apellido2, Cedula, RUC, Telefono, Direccion, CodCuentaImporta From Cuentas " & _
                                               "WHERE     (CodCuentas = '" & CodCuenta & "')  "
                 Me.AdoConsulta.Refresh
                 If Not Me.AdoConsulta.Recordset.EOF Then
                   If Not IsNull(Me.AdoConsulta.Recordset("CodCuentaImporta")) Then
                     CodCuenta = AdoConsulta.Recordset("CodCuentaImporta")
                   End If
                 End If
                 
                 For i = 1 To 19 - Len(CodCuenta)
                    CodCuenta = CodCuenta & " "
                 Next i
                                               
                 VoucherNo = AdoRegistros.Recordset("VoucherNo")
                 For i = 1 To 17 - Len(VoucherNo)
                  VoucherNo = VoucherNo & " "
                 Next i
                  VoucherNo = Mid(VoucherNo, 1, 17)
                 
'                 CodDepartamento = "   "
'                 CodAcciones = "                 "
                 CodAcciones = VoucherNo
                 ClaveProyecto = "               "
                 
                 If AdoRegistros.Recordset("FacturaNo") <> "-" Then
                    NFactura = AdoRegistros.Recordset("FacturaNo")
                 Else
                    NFactura = ""
                 End If
                 
                 For i = 1 To 8 - Len(NFactura)
                  NFactura = NFactura & " "
                 Next i
                  NFactura = Mid(NFactura, 1, 8)
                 
                 ReferenciaCh = Mid(AdoRegistros.Recordset("ChequeNo"), 1, 6)
                 For i = 1 To 6 - Len(ReferenciaCh)
                  ReferenciaCh = ReferenciaCh & " "
                 Next i
                  ReferenciaCh = Mid(ReferenciaCh, 1, 6)
                 
                 Descripcion = Mid(AdoRegistros.Recordset("DescripcionMovimiento"), 1, 35)
                 For i = 1 To 35 - Len(Descripcion)
                  Descripcion = Descripcion & " "
                 Next i
                  Descripcion = Mid(Descripcion, 1, 35)
                 
                 FechaDescuento = Format(AdoRegistros.Recordset("FechaDescuento"), "MMDDYYYY")
                 If FechaDescuento = "" Then
                  FechaDescuento = "01011900"
                 End If
                 FechaVencimiento = Format(AdoRegistros.Recordset("FechaVence"), "MMDDYYYY")
                 If FechaVencimiento = "" Then
                   FechaVencimiento = "01011900"
                 End If
                 ImporteDescuento = Format(AdoRegistros.Recordset("DescuentoDisponible"), "####0.00")
                 ValorUnit = 0
                 TipoTransaccion = "00"
                           
                 '/////////Verifico el tipo de movimiento//////////////
                   If TipoMovimiento = "01" Or TipoMovimiento = "02" Or TipoMovimiento = "03" Or TipoMovimiento = "04" Or TipoMovimiento = "10" Or TipoMovimiento = "11" Or TipoMovimiento = "12" Or TipoMovimiento = "13" Or TipoMovimiento = "14" Or TipoMovimiento = "20" Or TipoMovimiento = "21" Or TipoMovimiento = "27" Or TipoMovimiento = "31" Then
                     TextoMonto = Format(AdoRegistros.Recordset("Debito"), "####0.0000")
                   ElseIf TipoMovimiento = "05" Or TipoMovimiento = "06" Or TipoMovimiento = "07" Or TipoMovimiento = "08" Or TipoMovimiento = "09" Or TipoMovimiento = "15" Or TipoMovimiento = "16" Or TipoMovimiento = "17" Or TipoMovimiento = "18" Or TipoMovimiento = "19" Or TipoMovimiento = "22" Or TipoMovimiento = "28" Or TipoMovimiento = "29" Or TipoMovimiento = "30" Then
                      TextoMonto = Format(AdoRegistros.Recordset("Credito"), "####0.0000")
                    End If
                    
                    For i = 1 To 18 - Len(TextoMonto)
                       TextoMonto = " " & TextoMonto
                    Next i
                    
                    
                    mes = Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(mes)
                    If Longitud = 1 Then
                     mes = "0" & Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    
                    Dia = Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(Dia)
                    If Longitud = 1 Then
                     Dia = "0" & Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    ano = Cadena & Trim(Str(Year(AdoRegistros.Recordset("FechaTransaccion"))))
                    Cadena = mes & Dia & ano
                    Cadena = Cadena & NTransaccion
                    Cadena = Cadena & Fuente
                    Cadena = Cadena & CodCuenta
'                    Cadena = Cadena & Trim(Str(CodCuenta))
'                    For i = 1 To 36 - Len(Cadena)
'                    Cadena = Cadena & " "
'                    Next i
                    Cadena = Cadena & CodDepartamento & CodAcciones & ClaveProyecto & NFactura + TipoMovimiento + ReferenciaCh
                    Cadena = Cadena & Descripcion & FechaDescuento & FechaVencimiento
                    Cadena = Cadena & TextoMonto
                    
                    For i = 1 To 17 - Len(ImporteDescuento)
                       ImporteDescuento = " " & ImporteDescuento
                    Next i
                    For i = 1 To 17 - Len(ValorUnit)
                       ValorUnit = " " & ValorUnit
                    Next i
                    Cadena = Cadena & ImporteDescuento & ValorUnit
                    Cadena = Cadena & TipoTransaccion
                    Print #1, Cadena
                                    
                    
                    
                  AdoRegistros.Recordset.MoveNext
                  J = J + 1
                  Me.Caption = "Procesando:  " & J & " de " & Maximo & " Registros "
                  DoEvents
                  .Value = J
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Enlace"
                Salir = True
  End If
Else '//////En caso que no exista el Archivo///////////
                
                Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoRegistros.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   J = 0
                 Do While Not AdoRegistros.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                 
                 If Me.AdoRegistros.Recordset("Clave") = "Debito" Then
                    TipoMovimiento = "01"
                 Else
                    TipoMovimiento = "05"
                 End If
                 
                 Fuente = Mid(AdoRegistros.Recordset("Fuente"), 1, 4)
                 For i = 1 To 4 - Len(Fuente)
                  Fuente = Fuente & " "
                 Next i
                 
                 
                 Consecutivo = Cosecutivo + 1
                 Fecha = AdoRegistros.Recordset("FechaTransaccion")
                 NTransaccion = Format(AdoRegistros.Recordset("NumeroMovimiento"), "0000000#")
                 
                 
                 '////////////////////////////////////////////////////////////////////////////////////
                 '///////////////////////////BUSCO SI EXITE CUENTA DE EXPORTACION//////////////////////
                 '///////////////////////////////////////////////////////////////////////////////////
                 CodCuenta = AdoRegistros.Recordset("CodCuentas")

                 Me.AdoConsulta.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo, UbicacionReporte, SubDivicion, CausaIva , CausaRetencion, DescRetencion, Nombre1, Nombre2, Apellido1, Apellido2, Cedula, RUC, Telefono, Direccion, CodCuentaImporta From Cuentas " & _
                                               "WHERE     (CodCuentas = '" & CodCuenta & "')  "
                 Me.AdoConsulta.Refresh
                 If Not Me.AdoConsulta.Recordset.EOF Then
                   If Not IsNull(Me.AdoConsulta.Recordset("CodCuentaImporta")) Then
                     CodCuenta = AdoConsulta.Recordset("CodCuentaImporta")
                   End If
                 End If
                 
                 
                 
                 For i = 1 To 19 - Len(CodCuenta)
                      CodCuenta = CodCuenta & " "
                 Next i
                 
                 If Not IsNull(AdoRegistros.Recordset("VoucherNo")) Then
                 VoucherNo = AdoRegistros.Recordset("VoucherNo")
                 Else
                 VoucherNo = " "
                 End If
                 For i = 1 To 17 - Len(VoucherNo)
                  VoucherNo = VoucherNo & " "
                 Next i
                  VoucherNo = Mid(VoucherNo, 1, 17)
                 
'                 CodDepartamento = "   "
'                 CodAcciones = "                 "
                 CodAcciones = VoucherNo
                 ClaveProyecto = "               "
                 
                 If AdoRegistros.Recordset("FacturaNo") <> "-" Then
                    NFactura = AdoRegistros.Recordset("FacturaNo")
                 Else
                    NFactura = ""
                 End If
                 
                 NFactura = Mid(NFactura, 1, 8)
                 For i = 1 To 8 - Len(NFactura)
                  NFactura = NFactura & " "
                 Next i
                 
                 
                 If Not IsNull(AdoRegistros.Recordset("ChequeNo")) Then
                    ReferenciaCh = AdoRegistros.Recordset("ChequeNo")
                    For i = 1 To 6 - Len(ReferenciaCh)
                     ReferenciaCh = ReferenciaCh & " "
                    Next i
                    ReferenciaCh = Mid(ReferenciaCh, 1, 6)
                 End If
                 
                 Descripcion = Mid(AdoRegistros.Recordset("DescripcionMovimiento"), 1, 25)
                 For i = 1 To 35 - Len(Descripcion)
                  Descripcion = Descripcion & " "
                 Next i
'                 Descripcion = Mid(Descripcion, 1, 35)
                 
                 FechaDescuento = Format(AdoRegistros.Recordset("FechaDescuento"), "MMDDYYYY")
                 If FechaDescuento = "" Then
                  FechaDescuento = "01011900"
                 End If
                 FechaVencimiento = Format(AdoRegistros.Recordset("FechaVence"), "MMDDYYYY")
                 If FechaVencimiento = "" Then
                   FechaVencimiento = "01011900"
                 End If
                 ImporteDescuento = Format(AdoRegistros.Recordset("DescuentoDisponible"), "####0.00")
                 ValorUnit = 0
                 TipoTransaccion = "00"
                           
                 '/////////Verifico el tipo de movimiento//////////////
                   If TipoMovimiento = "01" Or TipoMovimiento = "02" Or TipoMovimiento = "03" Or TipoMovimiento = "04" Or TipoMovimiento = "10" Or TipoMovimiento = "11" Or TipoMovimiento = "12" Or TipoMovimiento = "13" Or TipoMovimiento = "14" Or TipoMovimiento = "20" Or TipoMovimiento = "21" Or TipoMovimiento = "27" Or TipoMovimiento = "31" Then
                     TextoMonto = Format(AdoRegistros.Recordset("Debito"), "####0.0000")
                   ElseIf TipoMovimiento = "05" Or TipoMovimiento = "06" Or TipoMovimiento = "07" Or TipoMovimiento = "08" Or TipoMovimiento = "09" Or TipoMovimiento = "15" Or TipoMovimiento = "16" Or TipoMovimiento = "17" Or TipoMovimiento = "18" Or TipoMovimiento = "19" Or TipoMovimiento = "22" Or TipoMovimiento = "28" Or TipoMovimiento = "29" Or TipoMovimiento = "30" Then
                      TextoMonto = Format(AdoRegistros.Recordset("Credito"), "####0.0000")
                    End If
                    
                    For i = 1 To 18 - Len(TextoMonto)
                       TextoMonto = " " & TextoMonto
                    Next i
                    
                    
                    mes = Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(mes)
                    If Longitud = 1 Then
                     mes = "0" & Trim(Str(Month(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    
                    Dia = Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    Longitud = Len(Dia)
                    If Longitud = 1 Then
                     Dia = "0" & Cadena & Trim(Str(Day(AdoRegistros.Recordset("FechaTransaccion"))))
                    End If
                    ano = Cadena & Trim(Str(Year(AdoRegistros.Recordset("FechaTransaccion"))))
                    Cadena = mes & Dia & ano
                    Cadena = Cadena & NTransaccion
                    Cadena = Cadena & Fuente
'                    Cadena = Cadena & Trim(CodCuenta)
                    Cadena = Cadena & CodCuenta

                    Cadena = Cadena & CodDepartamento & CodAcciones & ClaveProyecto & NFactura & TipoMovimiento & ReferenciaCh
                    Cadena = Cadena & Descripcion & FechaDescuento & FechaVencimiento
                    Cadena = Cadena & TextoMonto
                    
                    For i = 1 To 17 - Len(ImporteDescuento)
                       ImporteDescuento = " " & ImporteDescuento
                    Next i
                    For i = 1 To 17 - Len(ValorUnit)
                       ValorUnit = " " & ValorUnit
                    Next i
''                    Cadena = Cadena & ImporteDescuento & ValorUnit
''                    Cadena = Cadena & TipoTransaccion
                    Print #1, Cadena
                                    
                    
                    
                  AdoRegistros.Recordset.MoveNext
                  J = J + 1
                  .Value = J
                  Me.Caption = "Procesando:  " & J & " de " & Maximo & " Registros "
                  DoEvents
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Zeus Facturacion"



Salir = True
End If
Exit Sub
TipoErrs:
MsgBox err.Description
Salir = True
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()


MDIPrimero.Skin1.ApplySkin hWnd


Me.DTPFechaFin.Value = Format(Now, "dd/mm/yyyy")
Me.DTPFechaIni.Value = Format(Now, "dd/mm/yyyy")

 Me.DBGTransacciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGTransacciones.OddRowStyle.BackColor = &H80000005
 Me.DBGTransacciones.AlternatingRowStyle = True
 

With Me.AdoTransacciones
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.AdoRegistros
   .ConnectionString = Conexion
End With


With Me.AdoMovimientos
   .ConnectionString = Conexion
End With
End Sub

Private Sub PushButton1_Click()

End Sub

Private Sub OptAM_Click()
  If Me.OptAM.Value = True Then
         Me.CmdExportarPacioli.Visible = False
         Me.CmdExportarZeus.Visible = False
         Me.CmdExportarAM.Visible = True
  End If
End Sub

Private Sub OptPacioli_Click()
  If Me.OptPacioli.Value = True Then
     Me.CmdExportarPacioli.Visible = True
     Me.CmdExportarZeus.Visible = False
     Me.CmdExportarAM.Visible = False
  End If
End Sub

Private Sub OptZeus_Click()
  If Me.OptZeus.Value = True Then
         Me.CmdExportarPacioli.Visible = False
         Me.CmdExportarAM.Visible = False
         Me.CmdExportarZeus.Visible = True
  End If

End Sub

Private Sub RadioButton1_Click()

End Sub
