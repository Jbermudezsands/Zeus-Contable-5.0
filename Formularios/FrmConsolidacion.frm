VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmConsolidacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consolidacion de Compañias"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   12015
   Begin MSAdodcLib.Adodc AdoTransacciones 
      Height          =   375
      Left            =   4680
      Top             =   8400
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
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   495
      Left            =   6600
      Top             =   8520
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "AdoCuenta"
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
   Begin MSAdodcLib.Adodc AdoIndiceTransaccion 
      Height          =   375
      Left            =   5520
      Top             =   9120
      Width           =   4215
      _ExtentX        =   7435
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
      Caption         =   "AdoIndiceTransaccion"
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
   Begin MSAdodcLib.Adodc AdoPeriodos 
      Height          =   375
      Left            =   5160
      Top             =   8880
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
   Begin MSAdodcLib.Adodc AdoAnexar 
      Height          =   375
      Left            =   5400
      Top             =   9120
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
      Caption         =   "AdoAnexar"
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
      Left            =   6120
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
   Begin XtremeSuiteControls.Resizer Resizer1 
      Height          =   30
      Left            =   7080
      TabIndex        =   32
      Top             =   4920
      Width           =   30
      _Version        =   786432
      _ExtentX        =   53
      _ExtentY        =   53
      _StockProps     =   1
   End
   Begin TrueOleDBGrid80.TDBGrid DBRegistro 
      Bindings        =   "FrmConsolidacion.frx":0000
      Height          =   3015
      Left            =   120
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo Cuenta"
      Columns(0).DataField=   "CodCuenta"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripcion Movimiento"
      Columns(1).DataField=   "Descripcion"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Transaccion No."
      Columns(2).DataField=   "NTransaccion"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Fecha Transaccion"
      Columns(3).DataField=   "Fecha"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Debito"
      Columns(4).DataField=   "ImporteTransaccionDebito"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Credito"
      Columns(5).DataField=   "ImporteTransaccionCredito"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   1085
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   16315377
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   16315377
      RowDividerColor =   16315377
      RowSubDividerColor=   16315377
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=184,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Named:id=33:Normal"
      _StyleDefs(55)  =   ":id=33,.parent=0"
      _StyleDefs(56)  =   "Named:id=34:Heading"
      _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(58)  =   ":id=34,.wraptext=-1"
      _StyleDefs(59)  =   "Named:id=35:Footing"
      _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   "Named:id=36:Selected"
      _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=37:Caption"
      _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(65)  =   "Named:id=38:HighlightRow"
      _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=39:EvenRow"
      _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=40:OddRow"
      _StyleDefs(70)  =   ":id=40,.parent=33"
      _StyleDefs(71)  =   "Named:id=41:RecordSelector"
      _StyleDefs(72)  =   ":id=41,.parent=34"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc AdoConsecutivo 
      Height          =   375
      Left            =   -240
      Top             =   8760
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "AdoConsecutivo"
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
   Begin MSAdodcLib.Adodc AdoSuma 
      Height          =   375
      Left            =   2520
      Top             =   8760
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "AdoSuma"
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
   Begin MSAdodcLib.Adodc AdoImporta 
      Height          =   375
      Left            =   240
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
      Caption         =   "AdoImporte"
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
      Left            =   240
      Top             =   8880
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
   Begin MSAdodcLib.Adodc AdoMovimientos 
      Height          =   375
      Left            =   -1080
      Top             =   8760
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   1320
      Top             =   9120
      Visible         =   0   'False
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
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   7680
      TabIndex        =   20
      Top             =   3360
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Consultar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame Frame4 
      Caption         =   "Informacion de Periodo"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   4575
      Begin VB.ComboBox CmbFin 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmConsolidacion.frx":0019
         Left            =   1320
         List            =   "FrmConsolidacion.frx":0041
         TabIndex        =   12
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox CmbIni 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmConsolidacion.frx":006C
         Left            =   1320
         List            =   "FrmConsolidacion.frx":0094
         TabIndex        =   11
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   2040
         TabIndex        =   7
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton Option8 
            Caption         =   "2000"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option7 
            Caption         =   "2000"
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option6 
            Caption         =   "2000"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo Desde:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo Hasta:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   4800
      TabIndex        =   5
      Top             =   1200
      Width           =   7095
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1200
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   5295
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
         Left            =   6480
         Picture         =   "FrmConsolidacion.frx":00BF
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TxtNomibreArchivo 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   345
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   609
         _Version        =   393216
         Format          =   79364097
         CurrentDate     =   40603
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre del Archivo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
      Begin XtremeSuiteControls.RadioButton OptImportar 
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Importar"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptExportar 
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   12135
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "FrmConsolidacion.frx":020D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   12120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Consolidacion de Compañias"
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
         Left            =   4080
         TabIndex        =   1
         Top             =   360
         Width           =   4065
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   375
      Left            =   9120
      TabIndex        =   21
      Top             =   3360
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Procesar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   375
      Left            =   10560
      TabIndex        =   22
      Top             =   3360
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmConsolidacion.frx":1D4F
      Height          =   3015
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5318
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
      Columns(2).Caption=   "DescripcionMovimiento"
      Columns(2).DataField=   "DescripcionMovimiento"
      Columns(2).DataWidth=   255
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TCambio"
      Columns(3).DataField=   "TCambio"
      Columns(3).DataWidth=   23
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Debito"
      Columns(4).DataField=   "MDebito"
      Columns(4).DataWidth=   22
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Credito"
      Columns(5).DataField=   "MCredito"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Movimientos de Indices"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3704"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3625"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131588"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3254"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=131588"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(3)._AlignLeft=0"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=3254"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=131586"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(4)._AlignLeft=0"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=131586"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
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
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=86,.parent=13,.alignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
      _StyleDefs(59)  =   "Named:id=33:Normal"
      _StyleDefs(60)  =   ":id=33,.parent=0"
      _StyleDefs(61)  =   "Named:id=34:Heading"
      _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   ":id=34,.wraptext=-1"
      _StyleDefs(64)  =   "Named:id=35:Footing"
      _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   "Named:id=36:Selected"
      _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(68)  =   "Named:id=37:Caption"
      _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(70)  =   "Named:id=38:HighlightRow"
      _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(72)  =   "Named:id=39:EvenRow"
      _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(74)  =   "Named:id=40:OddRow"
      _StyleDefs(75)  =   ":id=40,.parent=33"
      _StyleDefs(76)  =   "Named:id=41:RecordSelector"
      _StyleDefs(77)  =   ":id=41,.parent=34"
      _StyleDefs(78)  =   "Named:id=42:FilterBar"
      _StyleDefs(79)  =   ":id=42,.parent=33"
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   7455
      _Version        =   786432
      _ExtentX        =   13150
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSMask.MaskEdBox TxtDebito 
      Height          =   375
      Left            =   8640
      TabIndex        =   27
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "##,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtCredito 
      Height          =   375
      Left            =   10320
      TabIndex        =   28
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "##,##0.00"
      PromptChar      =   "_"
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "FrmConsolidacion.frx":1D6C
      TabIndex        =   29
      Top             =   7440
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "FrmConsolidacion.frx":1DD8
      TabIndex        =   30
      Top             =   7440
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblRegistros 
      Height          =   495
      Left            =   1680
      OleObjectBlob   =   "FrmConsolidacion.frx":1E42
      TabIndex        =   33
      Top             =   7080
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "FrmConsolidacion.frx":1EA0
      TabIndex        =   34
      Top             =   7080
      Width           =   975
   End
End
Attribute VB_Name = "FrmConsolidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscaCuenta_Click()

Me.CommonDialog1.Filter = "Archivos Zeus|*.Cns"

If Me.OptExportar.Value = True Then
  Me.CommonDialog1.ShowSave
Else
  Me.CommonDialog1.ShowOpen
End If
Directorio = Me.CommonDialog1.FileName
If Directorio = "" Then
 Exit Sub
Else
  Me.TxtNomibreArchivo.Text = Directorio
End If

 
End Sub

Private Sub Form_Load()
With Me.DtaConsulta
   .ConnectionString = Conexion
End With

With Me.AdoMovimientos
   .ConnectionString = Conexion
End With

With Me.AdoAnexar
   .ConnectionString = Conexion
End With

With Me.AdoTransacciones
   .ConnectionString = Conexion
End With

With Me.AdoIndiceTransaccion
   .ConnectionString = Conexion
End With

With Me.AdoCuenta
   .ConnectionString = Conexion
End With

With Me.AdoPeriodos
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.AdoRegistros
   .ConnectionString = Conexion
End With

With Me.AdoImporta
   .ConnectionString = Conexion
End With
With Me.AdoSuma
   .ConnectionString = Conexion
End With

With Me.AdoConsecutivo
   .ConnectionString = Conexion
   .RecordSource = "NConsecutivos"
   .Refresh
End With

 Me.DBGTransacciones.EvenRowStyle.BackColor = &H80FFFF
 Me.DBGTransacciones.OddRowStyle.BackColor = &HC0FFFF
 Me.DBGTransacciones.AlternatingRowStyle = True
 
   Me.DBGTransacciones.Columns(4).NumberFormat = "##,##0.00"
   Me.DBGTransacciones.Columns(5).NumberFormat = "##,##0.00"
   
   
    Me.DBRegistro.EvenRowStyle.BackColor = RGB(216, 228, 248)
    Me.DBRegistro.OddRowStyle.BackColor = &H80000005
    Me.DBRegistro.AlternatingRowStyle = True
    Me.DBRegistro.Columns(4).NumberFormat = "##,##0.00"
    Me.DBRegistro.Columns(5).NumberFormat = "##,##0.00"



            Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
            Me.DtaConsulta.Refresh
            Do While Not DtaConsulta.Recordset.EOF
             If AÑO1 = "" Then
              AÑO1 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option8.Caption = AÑO1
             ElseIf AÑO2 = "" Then
              AÑO2 = Year(DtaConsulta.Recordset("FechaPeriodo"))
              Me.Option7.Caption = AÑO2
             Else
               AÑO3 = Year(DtaConsulta.Recordset("FechaPeriodo"))
               Me.Option6.Caption = AÑO3
             End If
              
              Me.DtaConsulta.Recordset.MoveNext
            Loop
            
End Sub

Private Sub OptExportar_Click()
  Dim NombreArchivo As String
  
   NombreArchivo = App.Path & "\" & MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa") & ".Cns"
   Me.Label7.Enabled = True
   Me.Label8.Enabled = True
   Me.CmbIni.Enabled = True
   Me.CmbFin.Enabled = True
   Me.Option6.Enabled = True
   Me.Option7.Enabled = True
   Me.Option8.Enabled = True
   Me.dtfecha.Enabled = False
   Me.TxtDescripcion.Visible = True
   Me.Label3.Visible = True
   
   Me.TxtNomibreArchivo.Text = NombreArchivo
   Me.TxtDescripcion.Text = MDIPrimero.AdoConfiguracion.Recordset("NombreEmpresa") & " Consolidacion"
 

End Sub

Private Sub OptImportar_Click()
   Me.Label7.Enabled = False
   Me.Label8.Enabled = False
   Me.CmbIni.Enabled = False
   Me.CmbFin.Enabled = False
   Me.Option6.Enabled = False
   Me.Option7.Enabled = False
   Me.Option8.Enabled = False
   Me.dtfecha.Enabled = True
   Me.TxtDescripcion.Visible = False
   Me.Label3.Visible = False
   
   Me.TxtNomibreArchivo.Text = ""
   Me.TxtDescripcion.Text = ""
 
End Sub

Private Sub PushButton1_Click()
Dim NumeroTabla As Double, NumeroPeriodo1 As Double, NumeroPeriodo2 As Double, i As Double
Dim FechaIni As String, FechaFin As String, Fecha1 As String, Fecha2 As String, NumFecha1 As Date
Dim NumFecha2 As Date, Contador As Integer, Campo As String, Fechas As String, FechaT1 As String, FechaT2 As String, FechaT3 As String
Dim TotalDebito As Double, TotalCredito As Double, Cadena As String, Consecutivo As Integer, NTransaccion As Double, Anulado As Integer, CodigoArchivo As String, CodigoCuenta As String
Dim Fuente As String, CodCuenta As String, CodDepartamento As String, CodAcciones As String, ClaveProyecto As String, NFactura As String, Ttransaccion As String, ReferenciaCh As String, Descripcion As String, FechaDescuento As String, FechaVencimiento As String, ImporteDescuento As String, ImporteTransaccion As Double, ValorUnit As String
Dim TipoTransaccion As String, Longitud As Double, Ultimo As Boolean, Encontrado As Boolean, Conexion As String, Directorio As String, Buscado As Boolean, TipoMovimiento As String, Salir As Boolean, Continuar As Boolean


Me.Height = 8295

Me.DBGTransacciones.Visible = True
Me.DBRegistro.Visible = False

If Me.OptExportar.Value = True Then


        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////BUSCO LA FECHA /////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
                    NumeroPeriodo1 = Me.CmbIni.Text
                    NumeroPeriodo2 = Me.CmbFin.Text
                    
                    If Me.Option8 = True Then
                     NumeroTabla = 1
                    ElseIf Me.Option7 = True Then
                      NumeroTabla = 2
                    ElseIf Me.Option6 = True Then
                      NumeroTabla = 3
                    End If
                    
                      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
                      Me.DtaConsulta.Refresh
                      If Me.DtaConsulta.Recordset.RecordCount = 0 Then
                        MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                        Exit Sub
                      End If
                       Me.DtaConsulta.Recordset.MoveLast
                       i = Me.DtaConsulta.Recordset.RecordCount
                       Me.DtaConsulta.Recordset.MoveFirst
                      Do While Not DtaConsulta.Recordset.EOF
                
                
                        If i = 1 Then
                          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                          NumFecha1 = FechaIni
                          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                        Else
                
                         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
                          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                          NumFecha1 = FechaIni
                         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
                          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                         End If
                        End If
                        Me.DtaConsulta.Recordset.MoveNext
                      Loop
                      
                      Fecha1 = Format(FechaIni, "yyyy-mm-dd")
                      Fecha2 = Format(FechaFin, "yyyy-mm-dd")
                      
                      
             '/////////////////////////////////////////////////////////////////////////////////////////////
             '////////////////////////////////////LLENO LA CONSULA //////////////////////////////////////////
             '///////////////////////////////////////////////////////////////////////////////////////////////
                         
             Me.AdoMovimientos.RecordSource = "SELECT CodCuentas, NombreCuenta, MAX(DescripcionMovimiento) AS DescripcionMovimiento, AVG(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(ROUND(Debito * TCambio, 2)) AS MDebito, SUM(ROUND(Credito * TCambio, 2)) AS MCredito From Transacciones " & _
                                              "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME,  '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY CodCuentas, NombreCuenta ORDER BY CodCuentas"
             Me.AdoMovimientos.Refresh
             
             Me.DBGTransacciones.Columns(0).Width = 1500
             Me.DBGTransacciones.Columns(1).Width = 2500
             Me.DBGTransacciones.Columns(2).Width = 2500
             Me.DBGTransacciones.Columns(3).Width = 1000
             
            Me.AdoSuma.RecordSource = "SELECT  MAX(CodCuentas) AS Expr3, MAX(NombreCuenta) AS Expr2, MAX(DescripcionMovimiento) AS Expr1, AVG(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(ROUND(Debito * TCambio, 2)) AS MDebito, SUM(ROUND(Credito * TCambio, 2)) AS MCredito From Transacciones " & _
                                      "WHERE (FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) ORDER BY MAX(CodCuentas)"
            Me.AdoSuma.Refresh
            
            Me.TxtDebito.Text = Format(Me.AdoSuma.Recordset("MDebito"), "##,##0.00")
            Me.TxtCredito.Text = Format(Me.AdoSuma.Recordset("MCredito"), "##,##0.00")
            
 Else

            
            '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            '//////////////////////////////////////////////////////IMPORTE EL ARCHIVO A LOS REGISTROS //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
            Directorio = Me.TxtNomibreArchivo.Text
            If Directorio = "" Then
             Exit Sub
            End If
            
            Me.AdoRegistros.RecordSource = "SELECT * From Registros"
            Me.AdoRegistros.Refresh
            Do While Not Me.AdoRegistros.Recordset.EOF
              Me.AdoRegistros.Recordset.Delete
             Me.AdoRegistros.Recordset.MoveNext
            Loop
            
            Consecutivo = 0
            crlf$ = Chr(13) & Chr(10)
            Contador = 0
            Open Directorio For Input As #1
            
            Me.AdoImporta.RecordSource = "SELECT Registros.Control, Registros.IdRegistros, Registros.Fecha, Registros.NTransaccion, Registros.Fuente, Registros.CodCuenta, Registros.Descripcion, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto, Registros.FacturaNumero, Registros.TipoMovimiento, Registros.RefCheque, Registros.FechaDescuento, Registros.FechaVencimiento, Registros.ImporteTransaccionDebito, Registros.ImporteTransaccionCredito, Registros.ImporteDescuento, Registros.ValorUnitario, Registros.TipoTransaccion From Registros Where (((Registros.IdRegistros) = " & Consecutivo & ")) "
            Me.AdoImporta.Refresh
            
            If Consecutivo = 0 Then
              Me.AdoConsecutivo.Recordset("ConsecutivoImporta") = AdoConsecutivo.Recordset("ConsecutivoImporta") + 1
             Me.AdoConsecutivo.Recordset.Update
             If Not IsNull(AdoConsecutivo.Recordset("ConsecutivoImporta")) Then
             Consecutivo = AdoConsecutivo.Recordset("ConsecutivoImporta")
             Else
              Consecutivo = 1
             End If
            End If
            Anulado = 0
            While Not EOF(1)
            Salir = False
            Line Input #1, Cadena
            
             Fechas = Mid(Cadena, 1, 8)
             FechaT2 = Mid(Fechas, 1, 2)
             FechaT1 = Mid(Fechas, 3, 2)
             FechaT3 = Mid(Fechas, 5, 4)
             Fecha = FechaT1 & "/" & FechaT2 + "/" + FechaT3
            
             NTransaccion = Mid(Cadena, 9, 8)
             Fuente = Mid(Cadena, 17, 4)
             CodCuenta = Mid(Cadena, 21, 16)
             CodDepartamento = Mid(Cadena, 37, 3)
             CodAcciones = Mid(Cadena, 40, 17)
             ClaveProyecto = Mid(Cadena, 57, 15)
             NFactura = Mid(Cadena, 72, 8)
             Ttransaccion = Mid(Cadena, 80, 2)
             ReferenciaCh = Mid(Cadena, 82, 6)
             Descripcion = Mid(Cadena, 88, 35)
             
             
             '/////////////QUITOS LOS ESPACIOS DEL CODIGO CUENTA///////////////
              Longitud = Len(CodCuenta)
              CodigoArchivo = ""
              CodigoCuenta = ""
             For i = 1 To Longitud
                 CodigoArchivo = Mid(CodCuenta, i, 1)
                 If CodigoArchivo <> " " Then
                   CodigoCuenta = CodigoCuenta & CodigoArchivo
            
                 End If
               
               Next
            
               CodCuenta = CodigoCuenta
            
            
             '///////Convierto el formato de fecha////////////////
             
             
             ImporteTransaccion = Mid(Cadena, 140, 17)
             ImporteDescuento = Mid(Cadena, 157, 17)
             ValorUnit = Mid(Cadena, 174, 17)
             TipoTransaccion = Mid(Cadena, 191, 2)
            
            If Not CodCuenta = "                " Then
             If ImporteTransaccion <> 0 Then
              Me.AdoRegistros.Recordset.AddNew
              AdoRegistros.Recordset("IdRegistros") = Consecutivo
              AdoRegistros.Recordset("Fecha") = Fecha
              AdoRegistros.Recordset("NTransaccion") = NTransaccion
              AdoRegistros.Recordset("Fuente") = Fuente
              AdoRegistros.Recordset("CodCuenta") = CodCuenta
              AdoRegistros.Recordset("CodDepartamento") = CodDepartamento
              AdoRegistros.Recordset("CodAcciones") = CodAcciones
              AdoRegistros.Recordset("ClaveProyecto") = ClaveProyecto
              AdoRegistros.Recordset("FacturaNumero") = NFactura
              AdoRegistros.Recordset("TipoMovimiento") = Ttransaccion
              AdoRegistros.Recordset("RefCheque") = ReferenciaCh
              AdoRegistros.Recordset("Descripcion") = Descripcion
              If Not FechaDescuento = "" Then
            '  AdoRegistros.Recordset("FechaDescuento") = FechaDescuento
              End If
            '  AdoRegistros.Recordset("FechaVencimiento") = FechaVencimiento
              
              If Ttransaccion = "01" Or Ttransaccion = "02" Or Ttransaccion = "03" Or Ttransaccion = "04" Or Ttransaccion = "10" Or Ttransaccion = "11" Or Ttransaccion = "12" Or Ttransaccion = "13" Or Ttransaccion = "14" Or Ttransaccion = "20" Or Ttransaccion = "21" Or Ttransaccion = "23" Or Ttransaccion = "27" Or Ttransaccion = "31" Then
                AdoRegistros.Recordset("ImporteTransaccionDebito") = ImporteTransaccion
              ElseIf Ttransaccion = "05" Or Ttransaccion = "06" Or Ttransaccion = "07" Or Ttransaccion = "08" Or Ttransaccion = "09" Or Ttransaccion = "15" Or Ttransaccion = "16" Or Ttransaccion = "17" Or Ttransaccion = "18" Or Ttransaccion = "19" Or Ttransaccion = "22" Or Ttransaccion = "24" Or Ttransaccion = "25" Or Ttransaccion = "26" Or Ttransaccion = "28" Or Ttransaccion = "29" Or Ttransaccion = "30" Then
                AdoRegistros.Recordset("ImporteTransaccionCredito") = ImporteTransaccion
              Else
                  Cadena = "Este Archivo Contienen un Codigo Nuevo" & vbLf
                  Cadena = Cadena & "Llame a su Soporte Tecnico para Anexarlo" & vbLf
                  Cadena = Cadena & "Al listado Permitido, para Importar"
                MsgBox Cadena
                Exit Sub
              End If
            
              If ImporteDescuento <> "                 " Then
                AdoRegistros.Recordset("ImporteDescuento") = ImporteDescuento
              End If
              AdoRegistros.Recordset("ValorUnitario") = ValorUnit
              AdoRegistros.Recordset("TipoTransaccion") = TipoTransaccion
              
              
              Me.AdoRegistros.Recordset.Update
              Contador = Contador + 1
              Me.LblRegistros.Caption = Contador
              DoEvents
             Else
              Anulado = Anulado + 1
             End If
             Else
              Anulado = Anulado + 1
             End If
            Wend
            Close #1
            Me.AdoImporta.RecordSource = "SELECT Registros.Control, Registros.IdRegistros, Registros.Fecha, Registros.NTransaccion, Registros.Fuente, Registros.CodCuenta, Registros.Descripcion, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto, Registros.FacturaNumero, Registros.TipoMovimiento, Registros.RefCheque, Registros.FechaDescuento, Registros.FechaVencimiento, Registros.ImporteTransaccionDebito, Registros.ImporteTransaccionCredito, Registros.ImporteDescuento, Registros.ValorUnitario, Registros.TipoTransaccion From Registros Where (((Registros.IdRegistros) = " & Consecutivo & ")) "
            Me.AdoImporta.Refresh
            
            Me.DBGTransacciones.Visible = False
            Me.DBRegistro.Visible = True
            Me.DBRegistro.Columns(1).Width = 3000
           
            MsgBox "Proceso Terminado", vbExclamation, "Sistema de Enlace"
            If Not Anulado = 0 Then
             MsgBox "Se quitaron " & Anulado & " Transacciones anuladas"
            End If
            Me.AdoSuma.RecordSource = "SELECT Sum(Registros.ImporteTransaccionDebito) AS TotalDebito, Sum(Registros.ImporteTransaccionCredito) AS TotalCredito From Registros"
            Me.AdoSuma.Refresh
            TotalCredito = Format(AdoSuma.Recordset("TotalCredito"), "##,##0.00")
            Me.TxtCredito.Text = TotalCredito
            TotalDebito = Format(AdoSuma.Recordset("TotalDebito"), "##,##0.00")
            Me.TxtDebito.Text = TotalDebito
            Salir = True
 
 End If
End Sub

Private Sub PushButton2_Click()
'On Error GoTo TipoErrs
Dim SQLExporta As String, Longitud As Integer, Respuesta As Integer
Dim Cadena As String, mes As String, Dia As String, ano As String
Dim TextoMonto As String, TipoMovimiento As String, J As Integer
Dim Consecutivo As Double, FechaDescuento As String, FechaVencimiento As String
Dim NumeroTabla As Double, NumeroPeriodo1 As Double, NumeroPeriodo2 As Double, i As Double
Dim FechaIni As String, FechaFin As String, Fecha1 As String, Fecha2 As String, NumFecha1 As Date
Dim NumFecha2 As Date
Dim TipoCuenta As String


Me.Height = 4290
  
If Me.OptExportar.Value = True Then


        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////BUSCO LA FECHA /////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////
                    NumeroPeriodo1 = Me.CmbIni.Text
                    NumeroPeriodo2 = Me.CmbFin.Text
                    
                    If Me.Option8 = True Then
                     NumeroTabla = 1
                    ElseIf Me.Option7 = True Then
                      NumeroTabla = 2
                    ElseIf Me.Option6 = True Then
                      NumeroTabla = 3
                    End If
                    
                      Me.DtaConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
                      Me.DtaConsulta.Refresh
                      If Me.DtaConsulta.Recordset.RecordCount = 0 Then
                        MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                        Exit Sub
                      End If
                       Me.DtaConsulta.Recordset.MoveLast
                       i = Me.DtaConsulta.Recordset.RecordCount
                       Me.DtaConsulta.Recordset.MoveFirst
                      Do While Not DtaConsulta.Recordset.EOF
                
                
                        If i = 1 Then
                          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                          NumFecha1 = FechaIni
                          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                        Else
                
                         If NumeroPeriodo1 = Me.DtaConsulta.Recordset("Periodo") Then
                          FechaIni = "01/" & Month(Me.DtaConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.DtaConsulta.Recordset("FechaPeriodo"))
                          NumFecha1 = FechaIni
                         ElseIf NumeroPeriodo2 = Me.DtaConsulta.Recordset("Periodo") Then
                          FechaFin = Me.DtaConsulta.Recordset("FechaPeriodo")
                          NumFecha2 = Me.DtaConsulta.Recordset("FechaPeriodo")
                         End If
                        End If
                        Me.DtaConsulta.Recordset.MoveNext
                      Loop
                      
                      Fecha1 = Format(FechaIni, "yyyy-mm-dd")
                      Fecha2 = Format(FechaFin, "yyyy-mm-dd")

 


       '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
       '///////////////////////////////////EXPORTAR LOS REGISTROS DE LA CONSULTA //////////////////////////////////////////////////////
       '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                                              
       
        
        Me.AdoRegistros.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.NombreCuenta, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, AVG(Transacciones.TCambio) AS TCambio, SUM(Transacciones.Debito) AS Debito, SUM(Transacciones.Credito) AS Credito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio, 2)) AS MDebito, SUM(ROUND(Transacciones.Credito * Transacciones.TCambio, 2)) AS MCredito, Cuentas.TipoCuenta FROM  Transacciones INNER JOIN Cuentas ON Transacciones.CodCuentas = Cuentas.CodCuentas " & _
                                       "WHERE (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Transacciones.CodCuentas, Transacciones.NombreCuenta, Cuentas.TipoCuenta ORDER BY Transacciones.CodCuentas"
        Me.AdoRegistros.Refresh
        
        Salir = False
        Barra.Visible = True
'        Me.CommonDialog1.ShowSave
        Directorio = Me.TxtNomibreArchivo.Text
        
       
'        Directorio = Me.CommonDialog1.FileName
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
                           .MAX = Maximo
                           J = 0
                         Do While Not AdoRegistros.Recordset.EOF
                         '////////Inicialiso las variables/////////////////
                         
'                         If Me.AdoRegistros.Recordset("Clave") = "Debito" Then
'                            TipoMovimiento = "01"
'                         Else
'                            TipoMovimiento = "05"
'                         End If

                          TipoCuenta = Me.AdoRegistros.Recordset("TipoCuenta")
                         
'                         Fuente = Mid(AdoRegistros.Recordset("Fuente"), 1, 4)
                         Fuente = "CONS"
                         For i = 1 To 4 - Len(Fuente)
                          Fuente = Fuente & " "
                         Next i
                         
                         
                         Consecutivo = Cosecutivo + 1
'                         Fecha = AdoRegistros.Recordset("FechaTransaccion")
                          Fecha = Now
'                         NTransaccion = Format(AdoRegistros.Recordset("NumeroMovimiento"), "0000000#")
                          NTransaccion = Format(1, "0000000#")
                         CodCuenta = AdoRegistros.Recordset("CodCuentas")
                         
                         Me.AdoConsulta.RecordSource = "SELECT  Cuentas.* From Cuentas WHERE  (CodCuentas = '" & CodCuenta & "')"
                         Me.AdoConsulta.Refresh
                         If Not Me.AdoConsulta.Recordset.EOF Then
                            If Not IsNull(Me.AdoConsulta.Recordset("CodCuentaImporta")) Then
                              CodCuenta = Me.AdoConsulta.Recordset("CodCuentaImporta")
                            End If
                         End If
                         
                         CodDepartamento = "   "
                         CodAcciones = "                 "
                         ClaveProyecto = "               "
                         
'                         If AdoRegistros.Recordset("FacturaNo") <> "-" Then
'                            NFactura = AdoRegistros.Recordset("FacturaNo")
'                         Else
'                            NFactura = ""
'                         End If

                         NFactura = ""
                         
                         NFactura = Mid(NFactura, 1, 8)
                         For i = 1 To 8 - Len(NFactura)
                          NFactura = NFactura & " "
                         Next i
                         
                         
                         
'                         ReferenciaCh = AdoRegistros.Recordset("ChequeNo")
                         ReferenciaCh = ""
                         For i = 1 To 6 - Len(ReferenciaCh)
                          ReferenciaCh = ReferenciaCh & " "
                         Next i
                         ReferenciaCh = Mid(ReferenciaCh, 1, 6)
                         
                         If Me.TxtDescripcion.Text = "" Then
                            Descripcion = "Consolidacion de Compañia"
                         Else
                            Descripcion = Mid(Me.TxtDescripcion.Text, 1, 35)
                         End If
                         For i = 1 To 35 - Len(Descripcion)
                          Descripcion = Descripcion & " "
                         Next i
                         Descripcion = Mid(Descripcion, 1, 35)
                         
                         FechaDescuento = Format(Now, "MMDDYYYY")
                         FechaVencimiento = Format(Now, "MMDDYYYY")
                         ImporteDescuento = Format(0, "####0.00")
                         ValorUnit = 0
                         TipoTransaccion = "00"
                                   
                         '/////////Verifico el tipo de movimiento//////////////
'                           If TipoMovimiento = "01" Or TipoMovimiento = "02" Or TipoMovimiento = "03" Or TipoMovimiento = "04" Or TipoMovimiento = "10" Or TipoMovimiento = "11" Or TipoMovimiento = "12" Or TipoMovimiento = "13" Or TipoMovimiento = "14" Or TipoMovimiento = "20" Or TipoMovimiento = "21" Or TipoMovimiento = "27" Or TipoMovimiento = "31" Then
'                             TextoMonto = Format(AdoRegistros.Recordset("Debito"), "####0.0000")
'                           ElseIf TipoMovimiento = "05" Or TipoMovimiento = "06" Or TipoMovimiento = "07" Or TipoMovimiento = "08" Or TipoMovimiento = "09" Or TipoMovimiento = "15" Or TipoMovimiento = "16" Or TipoMovimiento = "17" Or TipoMovimiento = "18" Or TipoMovimiento = "19" Or TipoMovimiento = "22" Or TipoMovimiento = "28" Or TipoMovimiento = "29" Or TipoMovimiento = "30" Then
'                              TextoMonto = Format(AdoRegistros.Recordset("Credito"), "####0.0000")
'                            End If

                                Debito = AdoRegistros.Recordset("MDebito")
                                Credito = AdoRegistros.Recordset("MCredito")

                            If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                                
                                If Debito > Credito Then
                                 TextoMonto = Format(Debito - Credito, "####0.0000")
                                 TipoMovimiento = "01"
                                Else
                                 TextoMonto = Format(Credito - Debito, "####0.0000")
                                 TipoMovimiento = "05"
                                End If
                             Else
                               If Credito > Debito Then
                                  TextoMonto = Format(Credito - Debito, "####0.0000")
                                  TipoMovimiento = "05"
                               Else
                                  TextoMonto = Format(Debito - Credito, "####0.0000")
                                  TipoMovimiento = "01"
                               End If
                            End If
                            
                            For i = 1 To 18 - Len(TextoMonto)
                               TextoMonto = " " & TextoMonto
                            Next i
                            
                            
                            mes = Trim(Str(Month(Now)))
                            Longitud = Len(mes)
                            If Longitud = 1 Then
                             mes = "0" & Trim(Str(Month(Now)))
                            End If
                            
                            Dia = Cadena & Trim(Str(Day(Now)))
                            Longitud = Len(Dia)
                            If Longitud = 1 Then
                             Dia = "0" & Cadena & Trim(Str(Day(Now)))
                            End If
                            ano = Cadena & Trim(Str(Year(Now)))
                            Cadena = mes & Dia & ano
                            Cadena = Cadena & NTransaccion
                            Cadena = Cadena & Fuente
                            Cadena = Cadena & Trim(AdoRegistros.Recordset("CodCuentas"))
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
          End If
        Else '//////En caso que no exista el Archivo///////////
                        
                        Open Directorio For Output As #1
                        'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                        
                        AdoRegistros.Recordset.MoveFirst
                        With Barra
                           .Min = 0
                           .Value = 0
                           .MAX = Maximo
                           J = 0
                         Do While Not AdoRegistros.Recordset.EOF
                         '////////Inicialiso las variables/////////////////
                         
'                         If Me.AdoRegistros.Recordset("Clave") = "Debito" Then
'                            TipoMovimiento = "01"
'                         Else
'                            TipoMovimiento = "05"
'                         End If

                          TipoCuenta = Me.AdoRegistros.Recordset("TipoCuenta")
                         
'                         Fuente = Mid(AdoRegistros.Recordset("Fuente"), 1, 4)
                         Fuente = "CONS"
                         For i = 1 To 4 - Len(Fuente)
                          Fuente = Fuente & " "
                         Next i
                         
                         
                         Consecutivo = Cosecutivo + 1
'                         Fecha = AdoRegistros.Recordset("FechaTransaccion")
                          Fecha = Now
'                         NTransaccion = Format(AdoRegistros.Recordset("NumeroMovimiento"), "0000000#")
                          NTransaccion = Format(1, "0000000#")
                         CodCuenta = AdoRegistros.Recordset("CodCuentas")
                         CodDepartamento = "   "
                         CodAcciones = "                 "
                         ClaveProyecto = "               "
                         
'                         If AdoRegistros.Recordset("FacturaNo") <> "-" Then
'                            NFactura = AdoRegistros.Recordset("FacturaNo")
'                         Else
'                            NFactura = ""
'                         End If

                         NFactura = ""
                         
                         NFactura = Mid(NFactura, 1, 8)
                         For i = 1 To 8 - Len(NFactura)
                          NFactura = NFactura & " "
                         Next i
                         
                         
                         
'                         ReferenciaCh = AdoRegistros.Recordset("ChequeNo")
                         ReferenciaCh = ""
                         For i = 1 To 6 - Len(ReferenciaCh)
                          ReferenciaCh = ReferenciaCh & " "
                         Next i
                         ReferenciaCh = Mid(ReferenciaCh, 1, 6)
                         
                         If Me.TxtDescripcion.Text = "" Then
                            Descripcion = "Consolidacion de Compañia"
                         Else
                            Descripcion = Mid(Me.TxtDescripcion.Text, 1, 35)
                         End If
                         For i = 1 To 35 - Len(Descripcion)
                          Descripcion = Descripcion & " "
                         Next i
                         Descripcion = Mid(Descripcion, 1, 35)
                         
                         FechaDescuento = Format(Now, "MMDDYYYY")
                         FechaVencimiento = Format(Now, "MMDDYYYY")
                         ImporteDescuento = Format(0, "####0.00")
                         ValorUnit = 0
                         TipoTransaccion = "00"
                                   
                         '/////////Verifico el tipo de movimiento//////////////
'                           If TipoMovimiento = "01" Or TipoMovimiento = "02" Or TipoMovimiento = "03" Or TipoMovimiento = "04" Or TipoMovimiento = "10" Or TipoMovimiento = "11" Or TipoMovimiento = "12" Or TipoMovimiento = "13" Or TipoMovimiento = "14" Or TipoMovimiento = "20" Or TipoMovimiento = "21" Or TipoMovimiento = "27" Or TipoMovimiento = "31" Then
'                             TextoMonto = Format(AdoRegistros.Recordset("Debito"), "####0.0000")
'                           ElseIf TipoMovimiento = "05" Or TipoMovimiento = "06" Or TipoMovimiento = "07" Or TipoMovimiento = "08" Or TipoMovimiento = "09" Or TipoMovimiento = "15" Or TipoMovimiento = "16" Or TipoMovimiento = "17" Or TipoMovimiento = "18" Or TipoMovimiento = "19" Or TipoMovimiento = "22" Or TipoMovimiento = "28" Or TipoMovimiento = "29" Or TipoMovimiento = "30" Then
'                              TextoMonto = Format(AdoRegistros.Recordset("Credito"), "####0.0000")
'                            End If

                                Debito = AdoRegistros.Recordset("MDebito")
                                Credito = AdoRegistros.Recordset("MCredito")

                            If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
                                
                                If Debito > Credito Then
                                 TextoMonto = Format(Debito - Credito, "####0.0000")
                                 TipoMovimiento = "01"
                                Else
                                 TextoMonto = Format(Credito - Debito, "####0.0000")
                                 TipoMovimiento = "05"
                                End If
                             Else
                               If Credito > Debito Then
                                  TextoMonto = Format(Credito - Debito, "####0.0000")
                                  TipoMovimiento = "05"
                               Else
                                  TextoMonto = Format(Debito - Credito, "####0.0000")
                                  TipoMovimiento = "01"
                               End If
                            End If
                            
                            For i = 1 To 18 - Len(TextoMonto)
                               TextoMonto = " " & TextoMonto
                            Next i
                            
                            
                            mes = Trim(Str(Month(Now)))
                            Longitud = Len(mes)
                            If Longitud = 1 Then
                             mes = "0" & Trim(Str(Month(Now)))
                            End If
                            
                            Dia = Cadena & Trim(Str(Day(Now)))
                            Longitud = Len(Dia)
                            If Longitud = 1 Then
                             Dia = "0" & Cadena & Trim(Str(Day(Now)))
                            End If
                            ano = Cadena & Trim(Str(Year(Now)))
                            Cadena = mes & Dia & ano
                            Cadena = Cadena & NTransaccion
                            Cadena = Cadena & Fuente
                            Cadena = Cadena & Trim(AdoRegistros.Recordset("CodCuentas"))
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

       End If


Else
                    Dim FechaArchivo As Date, NumeroPeriodo As Integer, NumeroMovimiento As Integer
                    Dim Fechas1 As String, Fechas2 As String, MovimientoArchivo As Double, SQL As String
                    Dim FechaGrabada As Date, MovimientoGrabado As Double
                    Dim NombreCuenta As String, CodigoArchivo As String
                    Dim CodigoCuenta As String, CantidadRegistros As Double, ExisteCodigo As Boolean
                    Dim Abrir As String
                    
                    Directorio = Me.TxtNomibreArchivo.Text
                    
                    
                   
                     '////////////////////////////////////////////////////////////////////////
                     '//////BUSCO SI TODAS LAS CUENTAS DEL ARCHIVO EXISTEN EN LA BASE DE DATOS/
                     '/////////////////////////////////////////////////////////////////////////
                     
                    
                    Me.AdoRegistros.Recordset.MoveLast
                    CantidadRegistros = AdoRegistros.Recordset.RecordCount
                    
                    Me.Barra.Visible = True
                     
                    With Me.Barra
                     .Min = 0
                     .MAX = CantidadRegistros
                     .Value = 0
                     J = 1
                     
                     ExisteCodigo = True
                     Open Directorio For Output As #1
                         Print #1, "Zeus Contable"
                         Print #1, "Importacion de Transacciones"
                         Print #1, ""
                         Me.AdoRegistros.Recordset.MoveFirst
                         Do While Not Me.AdoRegistros.Recordset.EOF
                           Longitud = Len(Me.AdoRegistros.Recordset("CodCuenta"))
                           CodigoCuenta = ""
                           For i = 1 To Longitud
                             CodigoArchivo = Mid(Me.AdoRegistros.Recordset("CodCuenta"), i, 1)
                             If CodigoArchivo <> " " Then
                               CodigoCuenta = CodigoCuenta & Mid(Me.AdoRegistros.Recordset("CodCuenta"), i, 1)
                             
                             End If
                           
                           Next
                           
'                              Me.LblProgreso.Caption = "Extraccion de la Cuenta: " & CodigoCuenta
                              DoEvents
                               
                               SQL = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
                                     "From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
                               Me.AdoConsulta.RecordSource = SQL
                               Me.AdoConsulta.Refresh
                               If Me.AdoConsulta.Recordset.EOF Then
                                  Cadena = CodigoCuenta
                                  Print #1, Cadena
                                  ExisteCodigo = False
                               End If
                            
                              .Value = J
                                 J = J + 1
                               Me.AdoRegistros.Recordset.MoveNext
                        
                          Loop
                      Close #1
                     End With
                     
                     If ExisteCodigo = False Then
                       MsgBox "No existen Cuentas", vbCritical, "Sistema Contable"
                    
                       Abrir = "notepad.exe " & Directorio
                       Shell Abrir
                       Exit Sub
                     End If
                     
                     '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                     '****************************************************************************************************************************
                     '****************************************************************************************************************************
                     '****************************************************************************************************************************
                     '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                     '///////////////////////////////////SI EXISTE TODAS LAS CUENTAS DE AGREGAN A LAS TRANSACCIONES///////////////////////////////
                     '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////7/
                     '****************************************************************************************************************************
                     '****************************************************************************************************************************
                     '****************************************************************************************************************************
                     '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                     
                     
                      
                      SQL = "SELECT Control, IdRegistros, Fecha, NTransaccion, Fuente, CodCuenta, CodDepartamento, CodAcciones, ClaveProyecto, FacturaNumero, TipoMovimiento, " & _
                            "RefCheque, Descripcion, FechaVencimiento, ImporteTransaccionDebito, ImporteTransaccionCredito, ImporteDescuento, ValorUnitario, TipoTransaccion, " & _
                            "DebitoDolar , CreditoDolar, FechaDescuento From Registros ORDER BY Control "
                            
                     Me.AdoAnexar.RecordSource = SQL
                     Me.AdoAnexar.Refresh
                     
                      Me.AdoAnexar.Recordset.MoveLast
                      CantidadRegistros = Me.AdoAnexar.Recordset.RecordCount
                      
                      Me.AdoAnexar.Recordset.MoveFirst
                      With Me.Barra
                     .Min = 0
                     .MAX = CantidadRegistros
                     .Value = 0
                      J = 1
                    
                     Do While Not Me.AdoAnexar.Recordset.EOF
                     '////////////////////////////////////////////////////
                     '///////////////BUSCO EL CODIGO DE LA CUENTA//////////////////////////////
                     '////////////////////////////////////////////////////////////////////
                     
                    ' CodigoCuenta = ""
                    '  For I = 1 To Longitud + 1
                    '     CodigoArchivo = Mid(Me.AdoAnexar.Recordset("CodCuenta"), I, 1)
                    '     If CodigoArchivo <> " " Then
                    '       CodigoCuenta = CodigoCuenta & Mid(Me.AdoAnexar.Recordset("CodCuenta"), I, 1)
                         
                    '     End If
                    '
                    '   Next
                    '
                     
                     CodigoCuenta = Me.AdoAnexar.Recordset("CodCuenta")
                     
                     
                      '//////BUSCO EL PERIODO DEL MOVIMIENTOS PARA AGREGARLO//////////////////////////////
                    
                      Fechas1 = CDate("1/" & Month(Me.dtfecha.Value) & "/" & Year(Me.dtfecha.Value))
                      mes = Month(Fechas1)
                      Año = Year(Fechas1)
                      Fechas2 = DateSerial(Año, mes + 1, 1 - 1)
                      Me.AdoPeriodos.RecordSource = "SELECT NPeriodo, NumeroTabla, FechaPeriodo, EstadoPeriodo, NTransacciones, Periodo From Periodos WHERE (FechaPeriodo BETWEEN '" & Format(Fechas1, "yyyymmdd") & "' AND '" & Format(Fechas2, "yyyymmdd") & "')"
                      Me.AdoPeriodos.Refresh
                      If Not Me.AdoPeriodos.Recordset.EOF Then
                        NumeroPeriodo = Me.AdoPeriodos.Recordset("NPeriodo")
                        
                        If MovimientoArchivo = Me.AdoAnexar.Recordset("NTransaccion") And FechaArchivo = Me.AdoAnexar.Recordset("Fecha") Then
                         NumeroMovimiento = Me.AdoPeriodos.Recordset("NTransacciones")
                         MovimientoArchivo = Me.AdoAnexar.Recordset("NTransaccion")
                         FechaArchivo = Me.AdoAnexar.Recordset("Fecha")
                        Else
                         NumeroMovimiento = Me.AdoPeriodos.Recordset("NTransacciones") + 1
                         MovimientoArchivo = Me.AdoAnexar.Recordset("NTransaccion")
                         FechaArchivo = Me.AdoAnexar.Recordset("Fecha")
                       End If
                       
                        Fuente = Me.AdoAnexar.Recordset("Fuente")
                        '////////////////////BUSCO SI EL MOVIMIENTO ESTA CUADRADO//////////////////////////////
                         SQL = "SELECT     MAX(Control) AS Control, MAX(IdRegistros) AS IdRegistros, MAX(Fecha) AS Fecha, NTransaccion AS NTransaccion, SUM(ImporteTransaccionDebito) " & _
                             "AS Debito, SUM(ImporteTransaccionCredito) AS Credito, SUM(ImporteTransaccionDebito) - SUM(ImporteTransaccionCredito) AS Diferencia " & _
                             "From Registros GROUP BY NTransaccion " & _
                             "HAVING  (MAX(Fecha) = '" & Format(FechaArchivo, "yyyymmdd") & "') AND (NTransaccion = " & MovimientoArchivo & " )"
                         AdoSuma.RecordSource = SQL
                         AdoSuma.Refresh
                         If Not Me.AdoSuma.Recordset.EOF Then
                            If Format(Me.AdoSuma.Recordset("Diferencia"), "##,##0.00") <> 0# Then
                             MsgBox "Existen Transacciones Descuadradas", vbCritical, "Fecha: " & FechaArchivo & " Transac# " & MovimientoArchivo
                             Exit Sub
                         
                            End If
                         End If
                         
                         
                       '///////////////////////////////////////////////////////////////////////////////////
                       '///////SI NO EXISTE, PROBLEMA CON EL PERIODO O DESCUADRE AGREGO EL REGISTRO////////
                       '///////////AGREGO EL INDICE DE LA TRANSACCION/////////////////////////////////
                         
                       SQL = "SELECT  FechaTransaccion, NumeroMovimiento, DescripcionMovimiento, Nperiodo, Fuente, TipoMoneda " & _
                             "From IndiceTransaccion " & _
                             "Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroMovimiento & ") "
                         
                       Me.AdoIndiceTransaccion.RecordSource = SQL
                       Me.AdoIndiceTransaccion.Refresh
                       If Me.AdoIndiceTransaccion.Recordset.EOF Then
                    
                        NumeroMovimiento = NumeroMovimiento
                        Me.AdoIndiceTransaccion.Recordset.AddNew
                        Me.AdoIndiceTransaccion.Recordset("FechaTransaccion") = Format(Me.dtfecha.Value, "dd/mm/yyyy")
                        Me.AdoIndiceTransaccion.Recordset("NumeroMovimiento") = NumeroMovimiento
                        Me.AdoIndiceTransaccion.Recordset("DescripcionMovimiento") = "Importacion" & Me.AdoAnexar.Recordset("Descripcion")
                        Me.AdoIndiceTransaccion.Recordset("Nperiodo") = NumeroPeriodo
                        Me.AdoIndiceTransaccion.Recordset("Fuente") = Fuente
                        Me.AdoIndiceTransaccion.Recordset("TipoMoneda") = "Córdobas"
                        Me.AdoIndiceTransaccion.Recordset.Update
                        
                        
                        '/////////////////////////////////////////////////////////////////////////////////
                        '/////////////EDITO LA TABLA PERIODOS////////////////////////////////////////////
                        '////////////////////////////////////////////////////////////////////////////////
                         Me.AdoPeriodos.Recordset("NTransacciones") = Me.AdoPeriodos.Recordset("NTransacciones") + 1
                         Me.AdoPeriodos.Recordset.Update
                        
                       End If
                        
                    
                         
                         
                        '/////////////////////////////////////////////////////////////////////////////////
                        '//////////BUSCO LA CUENTA CONTABLE PARA EL NOMBRE////////
                        '/////////////////////////////////////////////////////////////
                        SQL = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
                            "From Cuentas WHERE (CodCuentas ='" & CodigoCuenta & "')"
                            
                        AdoCuenta.RecordSource = SQL
                        AdoCuenta.Refresh
                        If Not Me.AdoCuenta.Recordset.EOF Then
                         If Not IsNull(Me.AdoCuenta.Recordset("DescripcionCuentas")) Then
                          NombreCuenta = Me.AdoCuenta.Recordset("DescripcionCuentas")
                         Else
                          NombreCuenta = "CUENTA SIN DESCRIPCION?????"
                         End If
                        
                        End If
                         
                         
'                         Me.LblProgreso.Caption = "Agregando la Cuenta: " & CodigoCuenta & " " & NombreCuenta
                         DoEvents
                         
                        '//////////////////////////////////////////////////////////////////////
                        '///////BUSCO LA TRANSACCION PARA ANEXAR O AGREGAR EL REGISTRO////////
                        '///////////////////////////////////////////////////////////////////////
                        
                        SQL = "SELECT  CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, " & _
                             "Debito, Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada, FechaDescuento, DescuentoDisponible, " & _
                             "FechaVence From Transacciones " & _
                             "Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroMovimiento & ")"
                        
                        Me.AdoTransacciones.RecordSource = SQL
                        Me.AdoTransacciones.Refresh
                    
                          Me.AdoTransacciones.Recordset.AddNew
                           Me.AdoTransacciones.Recordset("CodCuentas") = CodigoCuenta
                           Me.AdoTransacciones.Recordset("FechaTransaccion") = Format(Me.dtfecha.Value, "dd/mm/yyyy")
                           Me.AdoTransacciones.Recordset("NPeriodo") = NumeroPeriodo
                           Me.AdoTransacciones.Recordset("NumeroMovimiento") = NumeroMovimiento
                           Me.AdoTransacciones.Recordset("NombreCuenta") = NombreCuenta
                           Me.AdoTransacciones.Recordset("DescripcionMovimiento") = Me.AdoAnexar.Recordset("Descripcion")
                           If Me.AdoAnexar.Recordset("ImporteTransaccionDebito") = 0 Then
                              Me.AdoTransacciones.Recordset("Clave") = "Credito"
                           Else
                              Me.AdoTransacciones.Recordset("Clave") = "Debito"
                           End If
                           Me.AdoTransacciones.Recordset("TCambio") = 1
                           'Me.AdoTransacciones.Recordset("FacturaNo") ="-"
                           Me.AdoTransacciones.Recordset("ChequeNo") = Me.AdoAnexar.Recordset("RefCheque")
                           Me.AdoTransacciones.Recordset("Fuente") = Fuente
                           Me.AdoTransacciones.Recordset("FechaTasas") = FechaArchivo
                           Me.AdoTransacciones.Recordset("Debito") = Format(Me.AdoAnexar.Recordset("ImporteTransaccionDebito"), "##,##0.00")
                           Me.AdoTransacciones.Recordset("Credito") = Format(Me.AdoAnexar.Recordset("ImporteTransaccionCredito"), "##,##0.00")
                              
                          Me.AdoTransacciones.Recordset.Update
                       
                        
                    
                        
                      Else
                       MsgBox "No Existe Periodos para este Registro", vbCritical, "Sistema Contable"
                       Exit Sub
                      End If
                     
                    
                      .Value = J
                        J = J + 1
                      Me.AdoAnexar.Recordset.MoveNext
                     Loop
                    End With
                     MsgBox "Proceso de Importacion Terminado!!!!", vbExclamation, "Sistema Contable"
  
End If
  
End Sub

Private Sub PushButton3_Click()
Unload Me
End Sub
