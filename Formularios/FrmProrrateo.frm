VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmProrrateo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prorrateo de Cuentas"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9015
   Begin MSAdodcLib.Adodc AdoDetalleTransaccion 
      Height          =   375
      Left            =   5160
      Top             =   8880
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
      Caption         =   "AdoDetalleTransaccion"
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
   Begin MSAdodcLib.Adodc AdoIndice 
      Height          =   375
      Left            =   600
      Top             =   9000
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
   Begin MSAdodcLib.Adodc AdoSuma 
      Height          =   495
      Left            =   600
      Top             =   9000
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
   Begin VB.TextBox TxtNumeroTabla 
      Height          =   285
      Left            =   2760
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtProrrateo 
      Height          =   375
      Left            =   5040
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc AdoProrrateo 
      Height          =   375
      Left            =   2520
      Top             =   9000
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
      Caption         =   "AdoProrrateo"
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
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   7080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdoDestino 
      Height          =   375
      Left            =   600
      Top             =   9000
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
      Caption         =   "AdoDestino"
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
   Begin MSAdodcLib.Adodc AdoOrigen 
      Height          =   375
      Left            =   5280
      Top             =   9120
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
      Caption         =   "AdoOrigen"
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
      Left            =   480
      Top             =   9120
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
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      Top             =   7080
      Width           =   1095
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   4920
      Width           =   8775
      _Version        =   786432
      _ExtentX        =   15478
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "DESTINO"
      UseVisualStyle  =   -1  'True
      Begin VB.CommandButton CmdBorrarLineaDestino 
         Caption         =   "Borrar Linea"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridDestino 
         Bindings        =   "FrmProrrateo.frx":0000
         Height          =   1335
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2355
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NumeroProrrateo"
         Columns(0).DataField=   "NumeroProrrateo"
         Columns(0).DataWidth=   50
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "CodCuenta"
         Columns(1).DataField=   "CodCuenta"
         Columns(1).DataWidth=   50
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "TipoProrrateo"
         Columns(2).DataField=   "TipoProrrateo"
         Columns(2).DataWidth=   50
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descripcion"
         Columns(3).DataField=   "Descripcion"
         Columns(3).DataWidth=   50
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MontoBase"
         Columns(4).DataField=   "MontoBase"
         Columns(4).DataWidth=   23
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Porciento"
         Columns(5).DataField=   "Porciento"
         Columns(5).DataWidth=   23
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Importe"
         Columns(6).DataField=   "Importe"
         Columns(6).DataWidth=   23
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1).Button=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(4)._AlignLeft=0"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(5)._AlignLeft=0"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=8196"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(6)._AlignLeft=0"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowAddNew     =   -1  'True
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Named:id=33:Normal"
         _StyleDefs(59)  =   ":id=33,.parent=0"
         _StyleDefs(60)  =   "Named:id=34:Heading"
         _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   ":id=34,.wraptext=-1"
         _StyleDefs(63)  =   "Named:id=35:Footing"
         _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=36:Selected"
         _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=37:Caption"
         _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(69)  =   "Named:id=38:HighlightRow"
         _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=39:EvenRow"
         _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(73)  =   "Named:id=40:OddRow"
         _StyleDefs(74)  =   ":id=40,.parent=33"
         _StyleDefs(75)  =   "Named:id=41:RecordSelector"
         _StyleDefs(76)  =   ":id=41,.parent=34"
         _StyleDefs(77)  =   "Named:id=42:FilterBar"
         _StyleDefs(78)  =   ":id=42,.parent=33"
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDestino 
         Height          =   375
         Left            =   5520
         OleObjectBlob   =   "FrmProrrateo.frx":0019
         TabIndex        =   33
         Top             =   1800
         Width           =   3135
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   8775
      _Version        =   786432
      _ExtentX        =   15478
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "ORIGEN"
      UseVisualStyle  =   -1  'True
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalOrigen 
         Height          =   375
         Left            =   5520
         OleObjectBlob   =   "FrmProrrateo.frx":0077
         TabIndex        =   32
         Top             =   1800
         Width           =   3135
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridOrigen 
         Bindings        =   "FrmProrrateo.frx":00D5
         Height          =   1335
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2355
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NumeroProrrateo"
         Columns(0).DataField=   "NumeroProrrateo"
         Columns(0).DataWidth=   50
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "CodCuenta"
         Columns(1).DataField=   "CodCuenta"
         Columns(1).DataWidth=   50
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "TipoProrrateo"
         Columns(2).DataField=   "TipoProrrateo"
         Columns(2).DataWidth=   50
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descripcion"
         Columns(3).DataField=   "Descripcion"
         Columns(3).DataWidth=   50
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MontoBase"
         Columns(4).DataField=   "MontoBase"
         Columns(4).DataWidth=   23
         Columns(4).NumberFormat=   "Standard"
         Columns(4).EditMask=   "##,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "% Porciento"
         Columns(5).DataField=   "Porciento"
         Columns(5).DataWidth=   23
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Importe"
         Columns(6).DataField=   "Importe"
         Columns(6).DataWidth=   23
         Columns(6).NumberFormat=   "Standard"
         Columns(6).EditMask=   "##,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1).Button=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(4)._AlignLeft=0"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(5)._AlignLeft=0"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=8196"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(6)._AlignLeft=0"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowAddNew     =   -1  'True
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Named:id=33:Normal"
         _StyleDefs(59)  =   ":id=33,.parent=0"
         _StyleDefs(60)  =   "Named:id=34:Heading"
         _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   ":id=34,.wraptext=-1"
         _StyleDefs(63)  =   "Named:id=35:Footing"
         _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=36:Selected"
         _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=37:Caption"
         _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(69)  =   "Named:id=38:HighlightRow"
         _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=39:EvenRow"
         _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(73)  =   "Named:id=40:OddRow"
         _StyleDefs(74)  =   ":id=40,.parent=33"
         _StyleDefs(75)  =   "Named:id=41:RecordSelector"
         _StyleDefs(76)  =   ":id=41,.parent=34"
         _StyleDefs(77)  =   "Named:id=42:FilterBar"
         _StyleDefs(78)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton CmdBorrarLineaOrigen 
         Caption         =   "Borrar Linea"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   7080
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9015
      TabIndex        =   15
      Top             =   0
      Width           =   9015
      Begin VB.Image Image2 
         Height          =   945
         Left            =   480
         Picture         =   "FrmProrrateo.frx":00ED
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   9000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Prorrateo de Cuentas"
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
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   2745
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8775
      _Version        =   786432
      _ExtentX        =   15478
      _ExtentY        =   2990
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   7200
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   40061
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   1200
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   40061
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   5400
         TabIndex        =   6
         Top             =   480
         Width           =   2415
         Begin VB.OptionButton Option6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2000"
            Height          =   255
            Left            =   1560
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2000"
            Height          =   255
            Left            =   840
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2000"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.ComboBox CmbIni 
         Height          =   315
         ItemData        =   "FrmProrrateo.frx":1C2F
         Left            =   2400
         List            =   "FrmProrrateo.frx":1C57
         TabIndex        =   5
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox CmbFin 
         Height          =   315
         ItemData        =   "FrmProrrateo.frx":1C82
         Left            =   4320
         List            =   "FrmProrrateo.frx":1CAA
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin TrueOleDBList80.TDBCombo TDBProrrateo 
         Bindings        =   "FrmProrrateo.frx":1CD5
         DataSource      =   "AdoProrrateo"
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
         _GAPHEIGHT      =   53
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
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   0
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   -1  'True
         ColumnFooters   =   0   'False
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         AutoSize        =   -1  'True
         ListField       =   "NumeroProrrateo"
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmProrrateo.frx":1CF0
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmProrrateo.frx":1D9A
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmProrrateo.frx":1E16
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmProrrateo.frx":1EA8
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmProrrateo.frx":1F1A
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmProrrateo.frx":1F98
         TabIndex        =   13
         Top             =   1200
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   5520
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   40061
      End
   End
End
Attribute VB_Name = "FrmProrrateo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PorAcumulado As Double


Private Sub CmdBorrar_Click()
Dim NumeroProrrateo As Double, Respuesta As Double

   Respuesta = MsgBox("Esta Seguro de borrar el Registro?", vbYesNo, "Sistema Contable")
   If Respuesta = 6 Then

       If FrmProrrateo.TxtProrrateo.Text <> "" Then
         NumeroProrrateo = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.AdoConsulta.RecordSource = "SELECT * From Prorrateo Where (NumeroProrrateo = " & NumeroProrrateo & " )"
         FrmProrrateo.AdoConsulta.Refresh
         If Not FrmProrrateo.AdoConsulta.Recordset.EOF Then
           FrmProrrateo.AdoConsulta.Recordset.Delete
         End If
         
       End If
         FrmProrrateo.DTPFecha.Enabled = True
         FrmProrrateo.CmbFin.Enabled = True
         FrmProrrateo.CmbIni.Enabled = True
         FrmProrrateo.Frame5.Enabled = True
         FrmProrrateo.TDBProrrateo.Enabled = True
    
    End If
End Sub

Private Sub CmdBorrarLineaDestino_Click()
Dim Respuesta As Double, NumeroProrrateo As Double

 On Error GoTo TipoErrs
   Respuesta = MsgBox("Esta Seguro de borrar el Registro?", vbYesNo, "Sistema Contable")
   If Respuesta = 6 Then
    Me.AdoDestino.Recordset.Delete
    Me.AdoDestino.Refresh
   End If
   
    NumeroProrrateo = Me.TxtProrrateo.Text
    Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                              "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
    Me.AdoSuma.Refresh
    If Not Me.AdoSuma.Recordset.EOF Then
           Me.LblDestino.Caption = " Total   " & Format(Me.AdoSuma.Recordset("Importe"), "##,##0.00")
     End If
     
 Exit Sub
TipoErrs:
    MsgBox err.Description
End Sub

Private Sub CmdBorrarLineaOrigen_Click()
Dim Respuesta As Double, NumeroProrrateo As Double
   Respuesta = MsgBox("Esta Seguro de borrar el Registro?", vbYesNo, "Sistema Contable")
   If Respuesta = 6 Then
    Me.AdoOrigen.Recordset.Delete
    Me.AdoOrigen.Refresh
   End If
   
    NumeroProrrateo = Me.TxtProrrateo.Text
    Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                              "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
    Me.AdoSuma.Refresh
    If Not Me.AdoSuma.Recordset.EOF Then
           Me.LblTotalOrigen.Caption = " Total   " & Format(Me.AdoSuma.Recordset("Importe"), "##,##0.00")
     End If
End Sub

Private Sub CmdGrabar_Click()
  Dim TotalOrigen As Double, TotalDestino As Double, NumeroProrrateo As Double
     
      If Me.TxtProrrateo.Text <> "" Then
          NumeroProrrateo = Me.TxtProrrateo.Text
  
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalOrigen = Me.AdoSuma.Recordset("Importe")
            End If
            
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalDestino = Me.AdoSuma.Recordset("Importe")
            End If
      End If
      
       If FrmProrrateo.TxtProrrateo.Text <> "" Then
         NumeroProrrateo = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.AdoConsulta.RecordSource = "SELECT * From Prorrateo Where (NumeroProrrateo = " & NumeroProrrateo & " )"
         FrmProrrateo.AdoConsulta.Refresh
         If Not FrmProrrateo.AdoConsulta.Recordset.EOF Then
            
            FrmProrrateo.AdoConsulta.Recordset("NumeroProrrateo") = FrmProrrateo.TxtProrrateo.Text
            FrmProrrateo.AdoConsulta.Recordset("FechaMovimiento") = Format(Me.DTPFecha.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("PeriodoIni") = Format(Me.DTPFechaIni.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("PeriodoFin") = Format(Me.DTPFechaFin.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("Periodo1") = Me.CmbIni.Text
            FrmProrrateo.AdoConsulta.Recordset("Periodo2") = Me.CmbFin.Text
            If Me.Option8.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 1
            ElseIf Me.Option7.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 2
            ElseIf Me.Option6.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 3
            End If
             FrmProrrateo.AdoConsulta.Recordset("Año") = Year(Me.DTPFechaFin.Value)
             If Me.TxtDescripcion.Text <> "" Then
             FrmProrrateo.AdoConsulta.Recordset("DescripcionMovimiento") = Me.TxtDescripcion.Text
             End If
            FrmProrrateo.AdoConsulta.Recordset.Update
         End If
     End If
            
     If Format(TotalOrigen, "##,##0.00") <> Format(TotalDestino, "##,##0.00") Then
       
       MsgBox "Debe Cuadra el Prorrateo, para Grabar", vbCritical, "Sistema Contable"
       Exit Sub
     End If
     
     
         FrmProrrateo.DTPFecha.Enabled = True
         FrmProrrateo.CmbFin.Enabled = True
         FrmProrrateo.CmbIni.Enabled = True
         FrmProrrateo.Frame5.Enabled = True
         FrmProrrateo.TDBProrrateo.Enabled = True
End Sub

Private Sub CmdNuevo_Click()
Dim Consecutivo As Double

  Dim TotalOrigen As Double, TotalDestino As Double, NumeroProrrateo As Double
     
      If Me.TxtProrrateo.Text <> "" Then
          NumeroProrrateo = Me.TxtProrrateo.Text
  
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalOrigen = Me.AdoSuma.Recordset("Importe")
            End If
            
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalDestino = Me.AdoSuma.Recordset("Importe")
            End If
      End If
            
     If Format(TotalOrigen, "##,##0.00") <> Format(TotalDestino, "##,##0.00") Then
       
       MsgBox "Debe Cuadra el Prorrateo, para Grabar", vbCritical, "Sistema Contable"
       Exit Sub
     End If




 Me.AdoProrrateo.Refresh
 If Me.AdoProrrateo.Recordset.EOF Then
  Consecutivo = 1
  Me.TxtProrrateo.Text = Consecutivo
  Me.TDBProrrateo.Text = Consecutivo
 Else
   Me.AdoProrrateo.Recordset.MoveLast
   Consecutivo = Me.AdoProrrateo.Recordset("NumeroProrrateo") + 1
     Me.TxtProrrateo.Text = Consecutivo
     Me.TDBProrrateo.Text = Consecutivo
 End If

         FrmProrrateo.DTPFecha.Enabled = True
         FrmProrrateo.CmbFin.Enabled = True
         FrmProrrateo.CmbIni.Enabled = True
         FrmProrrateo.Frame5.Enabled = True
         FrmProrrateo.TDBProrrateo.Enabled = True
End Sub

Private Sub CmdProcesar_Click()
  On Error GoTo TipoErrs
  Dim TotalOrigen As Double, TotalDestino As Double, NumeroProrrateo As Double, Respuesta As Double
  Dim Periodo As Double, NumeroPeriodo As Double, FechaIni As String, FechaFin As String, EstadoPeriodo As String, NumeroTransaccion As Double
  Dim Mes As Double, Año As Double
     
      If Me.TxtProrrateo.Text <> "" Then
          NumeroProrrateo = Me.TxtProrrateo.Text
  
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalOrigen = Me.AdoSuma.Recordset("Importe")
            End If
            
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalDestino = Me.AdoSuma.Recordset("Importe")
            End If
      End If
            
      If Format(TotalOrigen, "##,##0.00") <> Format(TotalDestino, "##,##0.00") Then
       
       MsgBox "Debe Cuadra el Prorrateo, para Procesar", vbCritical, "Sistema Contable"
       Exit Sub
     End If
     
    Respuesta = MsgBox("Esta Seguro de Procesar el Prorrateo?", vbYesNo, "Sistema Contable")
    If Respuesta = 6 Then
    

       
                 '//////////////////////////////////////////////////////////////////////////////////////////////
                 '//////////SI LA CUENTA EXISTE AGREGO LOS ENCABEZADOS///////////////////////////////////////
                 '/////////////////////////////////////////////////////////////////////////////////////////////
                 

                         Mes = Month(Me.DTPFecha.Value)
                         Año = Year(Me.DTPFecha.Value)
                         FechaIni = CDate("1/" & Month(Me.DTPFecha.Value) & "/" & Year(Me.DTPFecha.Value))
                         FechaFin = DateSerial(Año, Mes + 1, 1 - 1)

                 
                         Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between '" & Format(FechaIni, "yyyymmdd") & "' And '" & Format(FechaFin, "yyyymmdd") & "'))"
                         Me.AdoConsulta.Refresh
                         If Not Me.AdoConsulta.Recordset.EOF Then
                           Periodo = Me.AdoConsulta.Recordset("Periodo")
                            NumeroPeriodo = Me.AdoConsulta.Recordset("NPeriodo")
                            EstadoPeriodo = Me.AdoConsulta.Recordset("EstadoPeriodo")
                      


                              Me.AdoConsulta.Recordset("NTransacciones") = Me.AdoConsulta.Recordset("NTransacciones") + 1
                              Me.AdoConsulta.Recordset.Update
                              NumeroTransaccion = Me.AdoConsulta.Recordset("NTransacciones")
                              
                             '////////////////////////////////////////////////////////////////
                             '////////AGREGO LOS INDICES DE TRANSACCIONES//////
                             '/////////////////////////////////////////////////////////////////
                             
                              Me.AdoIndice.RecordSource = "SELECT  * From IndiceTransaccion"
                              Me.AdoIndice.Refresh
                             
                              Me.AdoIndice.Recordset.AddNew
                              Me.AdoIndice.Recordset("FechaTransaccion") = Me.DTPFecha.Value
                              If Me.TxtDescripcion.Text <> "" Then
                                Me.AdoIndice.Recordset("DescripcionMovimiento") = Me.TxtDescripcion.Text
                              End If
                              Me.AdoIndice.Recordset("NumeroMovimiento") = NumeroTransaccion
                              Me.AdoIndice.Recordset("Fuente") = "PRORRATEO"
                              Me.AdoIndice.Recordset("NPeriodo") = NumeroPeriodo
                              Me.AdoIndice.Recordset("TipoMoneda") = "Córdobas"
                              Me.AdoIndice.Recordset.Update
                         
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION ORIGEN////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              Me.AdoDetalleTransaccion.RecordSource = "SELECT  * From Transacciones"
                              Me.AdoDetalleTransaccion.Refresh
                              
                               Me.AdoOrigen.Refresh
                                Do While Not Me.AdoOrigen.Recordset.EOF
                                Me.AdoDetalleTransaccion.Recordset.AddNew
                                 Me.AdoDetalleTransaccion.Recordset("CodCuentas") = Me.AdoOrigen.Recordset("CodCuenta")
                                 Me.AdoDetalleTransaccion.Recordset("FechaTransaccion") = Format(Me.DTPFecha.Value, "dd/mm/yyyy")
                                 Me.AdoDetalleTransaccion.Recordset("NPeriodo") = NumeroPeriodo
                                 Me.AdoDetalleTransaccion.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 Me.AdoDetalleTransaccion.Recordset("NombreCuenta") = Me.AdoOrigen.Recordset("Descripcion")
                                 Me.AdoDetalleTransaccion.Recordset("DescripcionMovimiento") = "Procesado Prorrateo No " & Me.TDBProrrateo.Text & " Base " & Format(Me.AdoOrigen.Recordset("MontoBase"), "##,##0.00") & " Porciento " & Me.AdoOrigen.Recordset("Porciento") & "%"
                                 Me.AdoDetalleTransaccion.Recordset("Clave") = "Credito"
                                 Me.AdoDetalleTransaccion.Recordset("TCambio") = 1
                                 Me.AdoDetalleTransaccion.Recordset("Debito") = 0
                                 Me.AdoDetalleTransaccion.Recordset("Credito") = Format(Me.AdoOrigen.Recordset("Importe"), "##,##0.00")
                                 Me.AdoDetalleTransaccion.Recordset("Fuente") = "Prorrateo"
                                 Me.AdoDetalleTransaccion.Recordset("FechaTasas") = Format(Me.DTPFecha.Value, "dd/mm/yyyy")
                                Me.AdoDetalleTransaccion.Recordset.Update
                                
                                 Me.AdoOrigen.Recordset.MoveNext
                                Loop
                                
                              '///////////////////////////////////////////////////////////////////////////////////////////
                              '/////////////////////AGREGO EL DETALLE TRANSACCION DESTINO////////////////////////////////
                              '////////////////////////////////////////////////////////////////////////////////////////////
                              Me.AdoDetalleTransaccion.RecordSource = "SELECT  * From Transacciones"
                              Me.AdoDetalleTransaccion.Refresh
                              
                               Me.AdoDestino.Refresh
                                Do While Not Me.AdoDestino.Recordset.EOF
                                Me.AdoDetalleTransaccion.Recordset.AddNew
                                 Me.AdoDetalleTransaccion.Recordset("CodCuentas") = Me.AdoDestino.Recordset("CodCuenta")
                                 Me.AdoDetalleTransaccion.Recordset("FechaTransaccion") = Format(Me.DTPFecha.Value, "dd/mm/yyyy")
                                 Me.AdoDetalleTransaccion.Recordset("NPeriodo") = NumeroPeriodo
                                 Me.AdoDetalleTransaccion.Recordset("NumeroMovimiento") = NumeroTransaccion
                                 Me.AdoDetalleTransaccion.Recordset("NombreCuenta") = Me.AdoDestino.Recordset("Descripcion")
                                 Me.AdoDetalleTransaccion.Recordset("DescripcionMovimiento") = "Procesado Prorrateo No " & Me.TDBProrrateo.Text & " Base " & Me.AdoDestino.Recordset("MontoBase") & " Porciento " & Me.AdoDestino.Recordset("Porciento") & "%"
                                 Me.AdoDetalleTransaccion.Recordset("Clave") = "Debito"
                                 Me.AdoDetalleTransaccion.Recordset("TCambio") = 1
                                 Me.AdoDetalleTransaccion.Recordset("Debito") = Format(Me.AdoDestino.Recordset("Importe"), "##,##0.00")
                                 Me.AdoDetalleTransaccion.Recordset("Credito") = 0
                                 Me.AdoDetalleTransaccion.Recordset("Fuente") = "Prorrateo"
                                 Me.AdoDetalleTransaccion.Recordset("FechaTasas") = Format(Me.DTPFecha.Value, "dd/mm/yyyy")
                                Me.AdoDetalleTransaccion.Recordset.Update
                                
                                 Me.AdoDestino.Recordset.MoveNext
                                Loop
                                
                              
                         Else
                          MsgBox "No Existe Ningun Periodo para la Fecha Movimiento", vbCritical, "Sistema Contable"
                          FrmProrrateo.DTPFecha.Enabled = True
                          Exit Sub
                         End If
                       
    
    
    
    End If
    
   MsgBox "La Transaccion Fue Agregada con Exito", vbCritical, "Sistema Contable"
   
   Exit Sub
TipoErrs:
   MsgBox err.Description, vbCritical, "Sistema Contable"
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Dim AÑO1 As String, AÑO2 As String, AÑO3 As String
MDIPrimero.Skin1.ApplySkin hWnd

Me.DTPFecha.Value = Now
With Me.AdoIndice
   .ConnectionString = Conexion
End With

With Me.AdoDetalleTransaccion
   .ConnectionString = Conexion
End With

With Me.AdoOrigen
   .ConnectionString = Conexion
End With

With Me.AdoSuma
   .ConnectionString = Conexion
End With

With Me.AdoDestino
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.AdoProrrateo
   .ConnectionString = Conexion
End With

 Me.TDBGridDestino.EvenRowStyle.BackColor = &H80FFFF
 Me.TDBGridDestino.OddRowStyle.BackColor = &HC0FFFF
 Me.TDBGridDestino.AlternatingRowStyle = True
 Me.TDBGridDestino.Columns(4).NumberFormat = "##,##0.00"
 Me.TDBGridDestino.Columns(6).NumberFormat = "##,##0.00"
 
 Me.TDBGridOrigen.EvenRowStyle.BackColor = &H80FFFF
 Me.TDBGridOrigen.OddRowStyle.BackColor = &HC0FFFF
 Me.TDBGridOrigen.AlternatingRowStyle = True
 Me.TDBGridOrigen.Columns(4).NumberFormat = "##,##0.00"
 Me.TDBGridOrigen.Columns(6).NumberFormat = "##,##0.00"
 

Me.AdoProrrateo.RecordSource = "SELECT  NumeroProrrateo, FechaMovimiento, DescripcionMovimiento From Prorrateo"
Me.AdoProrrateo.Refresh


Me.AdoOrigen.RecordSource = "SELECT * From DetalleProrrateo WHERE (TipoProrrateo = 'ORIGEN')AND (NumeroProrrateo = - 1)"
Me.AdoOrigen.Refresh

'Me.TDBGridOrigen.Columns(0).Visible = False
'Me.TDBGridOrigen.Columns(2).Visible = False
Me.TDBGridOrigen.Columns(0).Button = True


Me.AdoDestino.RecordSource = "SELECT * From DetalleProrrateo WHERE (TipoProrrateo = 'DESTINO') AND (NumeroProrrateo = - 1)"
Me.AdoDestino.Refresh

Me.TDBGridDestino.Columns(0).Visible = False
Me.TDBGridDestino.Columns(2).Visible = False


      Me.AdoConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos Where (((Periodos.Periodo) = 1) And ((Periodos.NumeroTabla) = 1 Or (Periodos.NumeroTabla) = 2 Or (Periodos.NumeroTabla) = 3))"
      Me.AdoConsulta.Refresh
      Do While Not AdoConsulta.Recordset.EOF
       If AÑO1 = "" Then
        AÑO1 = Year(AdoConsulta.Recordset("FechaPeriodo"))
        Me.Option8.Caption = AÑO1
       ElseIf AÑO2 = "" Then
        AÑO2 = Year(AdoConsulta.Recordset("FechaPeriodo"))
        Me.Option7.Caption = AÑO2
       Else
         AÑO3 = Year(AdoConsulta.Recordset("FechaPeriodo"))
         Me.Option6.Caption = AÑO3
       End If
        
        Me.AdoConsulta.Recordset.MoveNext
      Loop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim TotalOrigen As Double, TotalDestino As Double, NumeroProrrateo As Double
     
      If Me.TxtProrrateo.Text <> "" Then
          NumeroProrrateo = Me.TxtProrrateo.Text
  
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalOrigen = Me.AdoSuma.Recordset("Importe")
            End If
            
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              TotalDestino = Me.AdoSuma.Recordset("Importe")
            End If
      End If
            
      If Format(TotalOrigen, "##,##0.00") = Format(TotalDestino, "##,##0.00") Then
       Cancel = 0
     Else
        
        MsgBox "Debe Cuadra el Prorrateo, para Salir", vbCritical, "Sistema Contable"
        Cancel = 1
     End If
            
End Sub

Private Sub TDBGridDestino_AfterUpdate()
       If FrmProrrateo.TxtProrrateo.Text <> "" Then
         NumeroProrrateo = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.AdoConsulta.RecordSource = "SELECT * From Prorrateo Where (NumeroProrrateo = " & NumeroProrrateo & " )"
         FrmProrrateo.AdoConsulta.Refresh
         If Not FrmProrrateo.AdoConsulta.Recordset.EOF Then
            
            FrmProrrateo.AdoConsulta.Recordset("NumeroProrrateo") = FrmProrrateo.TxtProrrateo.Text
            FrmProrrateo.AdoConsulta.Recordset("FechaMovimiento") = Format(Me.DTPFecha.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("PeriodoIni") = Format(Me.DTPFechaIni.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("PeriodoFin") = Format(Me.DTPFechaFin.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("Periodo1") = Me.CmbIni.Text
            FrmProrrateo.AdoConsulta.Recordset("Periodo2") = Me.CmbFin.Text
            If Me.Option8.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 1
            ElseIf Me.Option7.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 2
            ElseIf Me.Option6.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 3
            End If
             FrmProrrateo.AdoConsulta.Recordset("Año") = Year(Me.DTPFechaFin.Value)
             If Me.TxtDescripcion.Text <> "" Then
             FrmProrrateo.AdoConsulta.Recordset("DescripcionMovimiento") = Me.TxtDescripcion.Text
             End If
            FrmProrrateo.AdoConsulta.Recordset.Update
        
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              Me.LblTotalOrigen.Caption = " Total   " & Format(Me.AdoSuma.Recordset("Importe"), "##,##0.00")
            End If
            
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              Me.LblDestino.Caption = " Total   " & Format(Me.AdoSuma.Recordset("Importe"), "##,##0.00")
            End If
         End If
       Else

        MsgBox "Debe Crear primero el Prorrateo", vbCritical, "Sistema Contable"
        Exit Sub
       End If
       

         FrmProrrateo.DTPFecha.Enabled = False
         FrmProrrateo.CmbFin.Enabled = False
         FrmProrrateo.CmbIni.Enabled = False
         FrmProrrateo.Frame5.Enabled = False
         FrmProrrateo.TDBProrrateo.Enabled = False

End Sub

Private Sub TDBGridDestino_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  Dim Porciento As Double, Importe As Double, MontoBase As Double, NumeroProrrateo As Double
  If ColIndex = 5 Then
    If IsNumeric(Me.TDBGridDestino.Columns(ColIndex).Text) Then
            NumeroProrrateo = Me.TxtProrrateo.Text
                    Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                              "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
                    Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              MontoBase = Me.AdoSuma.Recordset("Importe")
            End If
    
       Porciento = Me.TDBGridDestino.Columns(5).Text
       Porciento = Porciento / 100
       
       Importe = Porciento * MontoBase
       Me.TDBGridDestino.Columns(6).Text = Importe
    Else
       Me.TDBGridDestino.Columns(ColIndex).Text = 0
    End If
  End If
End Sub

Private Sub TDBGridDestino_ButtonClick(ByVal ColIndex As Integer)
 If ColIndex = 1 Then
 
            NumeroPeriodo1 = Me.CmbIni.Text
            NumeroPeriodo2 = Me.CmbFin.Text
            
            If Me.Option8 = True Then
             NumeroTabla = 1
            ElseIf Me.Option7 = True Then
              NumeroTabla = 2
            ElseIf Me.Option6 = True Then
              NumeroTabla = 3
            End If
            
              Me.AdoConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              Me.AdoConsulta.Refresh
              If Me.AdoConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               Me.AdoConsulta.Recordset.MoveLast
               i = Me.AdoConsulta.Recordset.RecordCount
               Me.AdoConsulta.Recordset.MoveFirst
              Do While Not AdoConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(Me.AdoConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.AdoConsulta.Recordset("FechaPeriodo"))
                  FechaFin = Me.AdoConsulta.Recordset("FechaPeriodo")
                  
                Else
        
                 If NumeroPeriodo1 = Me.AdoConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(Me.AdoConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.AdoConsulta.Recordset("FechaPeriodo"))
                 ElseIf NumeroPeriodo2 = Me.AdoConsulta.Recordset("Periodo") Then
                  FechaFin = Me.AdoConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                Me.AdoConsulta.Recordset.MoveNext
              Loop
              
              Me.DTPFechaIni.Value = FechaIni
              Me.DTPFechaFin.Value = FechaFin
    
 
 
    QueProducto = "Prorrateo2"
    FrmConsulta.Show
 End If
End Sub

Private Sub TDBGridOrigen_AfterUpdate()
       If FrmProrrateo.TxtProrrateo.Text <> "" Then
         NumeroProrrateo = FrmProrrateo.TxtProrrateo.Text
         FrmProrrateo.AdoConsulta.RecordSource = "SELECT * From Prorrateo Where (NumeroProrrateo = " & NumeroProrrateo & " )"
         FrmProrrateo.AdoConsulta.Refresh
         If Not FrmProrrateo.AdoConsulta.Recordset.EOF Then
            
            FrmProrrateo.AdoConsulta.Recordset("NumeroProrrateo") = FrmProrrateo.TxtProrrateo.Text
            FrmProrrateo.AdoConsulta.Recordset("FechaMovimiento") = Format(Me.DTPFecha.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("PeriodoIni") = Format(Me.DTPFechaIni.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("PeriodoFin") = Format(Me.DTPFechaFin.Value, "dd/mm/yyyy")
            FrmProrrateo.AdoConsulta.Recordset("Periodo1") = Me.CmbIni.Text
            FrmProrrateo.AdoConsulta.Recordset("Periodo2") = Me.CmbFin.Text
            If Me.Option8.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 1
            ElseIf Me.Option7.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 2
            ElseIf Me.Option6.Value = True Then
              FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 3
            End If
             FrmProrrateo.AdoConsulta.Recordset("Año") = Year(Me.DTPFechaFin.Value)
             If Me.TxtDescripcion.Text <> "" Then
             FrmProrrateo.AdoConsulta.Recordset("DescripcionMovimiento") = Me.TxtDescripcion.Text
             End If
            FrmProrrateo.AdoConsulta.Recordset.Update
        
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
            Me.AdoSuma.Refresh
            If Not Me.AdoSuma.Recordset.EOF Then
              Me.LblTotalOrigen.Caption = " Total   " & Format(Me.AdoSuma.Recordset("Importe"), "##,##0.00")
            End If
            
            Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                      "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
            Me.AdoSuma.Refresh
         End If
       Else

        MsgBox "Debe Crear primero el Prorrateo", vbCritical, "Sistema Contable"
        Exit Sub
       End If
       

         FrmProrrateo.DTPFecha.Enabled = False
         FrmProrrateo.CmbFin.Enabled = False
         FrmProrrateo.CmbIni.Enabled = False
         FrmProrrateo.Frame5.Enabled = False
         FrmProrrateo.TDBProrrateo.Enabled = False

End Sub

Private Sub TDBGridOrigen_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  Dim Porciento As Double, Importe As Double, MontoBase As Double
  If ColIndex = 5 Then
    If IsNumeric(Me.TDBGridOrigen.Columns(ColIndex).Text) Then
       Porciento = Me.TDBGridOrigen.Columns(5).Text
       Porciento = Porciento / 100
       MontoBase = Me.TDBGridOrigen.Columns(4).Text
       Importe = Porciento * MontoBase
       Me.TDBGridOrigen.Columns(6).Text = Importe
    Else
       Me.TDBGridOrigen.Columns(ColIndex).Text = 0
    End If
  End If
End Sub

Private Sub TDBGridOrigen_ButtonClick(ByVal ColIndex As Integer)
Dim NumeroPeriodo1 As Double, NumeroPeriodo2 As Double, NumeroTabla As Double
Dim i As Double, FechaIni As String, FechaFin As String

 If ColIndex = 1 Then
 
            NumeroPeriodo1 = Me.CmbIni.Text
            NumeroPeriodo2 = Me.CmbFin.Text
            
            If Me.Option8 = True Then
             NumeroTabla = 1
            ElseIf Me.Option7 = True Then
              NumeroTabla = 2
            ElseIf Me.Option6 = True Then
              NumeroTabla = 3
            End If
            
              Me.AdoConsulta.RecordSource = "SELECT Periodos.Periodo, Periodos.FechaPeriodo, Periodos.NumeroTabla From Periodos WHERE (((Periodos.Periodo) = " & NumeroPeriodo1 & ") AND ((Periodos.NumeroTabla)=" & NumeroTabla & "))"
              Me.AdoConsulta.Refresh
              If Me.AdoConsulta.Recordset.RecordCount = 0 Then
                MsgBox "Hay un problema con los Periodos, por Favor Revise los Periodos para buscar alguna anomalía", vbCritical
                Exit Sub
              End If
               Me.AdoConsulta.Recordset.MoveLast
               i = Me.AdoConsulta.Recordset.RecordCount
               Me.AdoConsulta.Recordset.MoveFirst
              Do While Not AdoConsulta.Recordset.EOF
        
        
                If i = 1 Then
                  FechaIni = "01/" & Month(Me.AdoConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.AdoConsulta.Recordset("FechaPeriodo"))
                  FechaFin = Me.AdoConsulta.Recordset("FechaPeriodo")
                  
                Else
        
                 If NumeroPeriodo1 = Me.AdoConsulta.Recordset("Periodo") Then
                  FechaIni = "01/" & Month(Me.AdoConsulta.Recordset("FechaPeriodo")) & "/" & Year(Me.AdoConsulta.Recordset("FechaPeriodo"))
                 ElseIf NumeroPeriodo2 = Me.AdoConsulta.Recordset("Periodo") Then
                  FechaFin = Me.AdoConsulta.Recordset("FechaPeriodo")
                 End If
                End If
                Me.AdoConsulta.Recordset.MoveNext
              Loop
              
              Me.DTPFechaIni.Value = FechaIni
              Me.DTPFechaFin.Value = FechaFin
    
 
 
    QueProducto = "Prorrateo"
    FrmConsulta.Show
 End If
End Sub

Private Sub TDBProrrateo_Change()
Me.TxtProrrateo.Text = Me.TDBProrrateo.Text
End Sub

Private Sub TDBProrrateo_ItemChange()
Me.TxtProrrateo.Text = Me.TDBProrrateo.Text
End Sub

Private Sub TDBProrrateo_KeyPress(KeyAscii As Integer)
Me.TxtProrrateo.Text = Me.TDBProrrateo.Text
End Sub

Private Sub TxtProrrateo_Change()
       If FrmProrrateo.TxtProrrateo.Text <> "" Then
                 NumeroProrrateo = FrmProrrateo.TxtProrrateo.Text
                 FrmProrrateo.AdoConsulta.RecordSource = "SELECT * From Prorrateo Where (NumeroProrrateo = " & NumeroProrrateo & " )"
                 FrmProrrateo.AdoConsulta.Refresh
                 If Not FrmProrrateo.AdoConsulta.Recordset.EOF Then
                    
                    FrmProrrateo.TxtProrrateo.Text = FrmProrrateo.AdoConsulta.Recordset("NumeroProrrateo")
                    If Not IsNull(FrmProrrateo.AdoConsulta.Recordset("FechaMovimiento")) Then
                    Me.DTPFecha.Value = FrmProrrateo.AdoConsulta.Recordset("FechaMovimiento")
                    End If
                    If Not IsNull(FrmProrrateo.AdoConsulta.Recordset("PeriodoIni")) Then
                    Me.DTPFechaIni.Value = FrmProrrateo.AdoConsulta.Recordset("PeriodoIni")
                    End If
                    If Not IsNull(FrmProrrateo.AdoConsulta.Recordset("PeriodoFin")) Then
                    Me.DTPFechaFin.Value = FrmProrrateo.AdoConsulta.Recordset("PeriodoFin")
                    End If
                    If Not IsNull(FrmProrrateo.AdoConsulta.Recordset("Periodo1")) Then
                    Me.CmbIni.Text = FrmProrrateo.AdoConsulta.Recordset("Periodo1")
                    End If
                    If Not IsNull(FrmProrrateo.AdoConsulta.Recordset("Periodo2")) Then
                    Me.CmbFin.Text = FrmProrrateo.AdoConsulta.Recordset("Periodo2")
                    End If
                    If FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 1 Then
                      Me.Option8.Value = True
                    ElseIf FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 2 Then
                      Me.Option7.Value = True
                    ElseIf FrmProrrateo.AdoConsulta.Recordset("NumeroTabla") = 3 Then
                      Me.Option6.Value = True
                    End If
                     
                     If Not IsNull(FrmProrrateo.AdoConsulta.Recordset("DescripcionMovimiento")) Then
                      Me.TxtDescripcion.Text = FrmProrrateo.AdoConsulta.Recordset("DescripcionMovimiento")
                     End If
                    
                    Me.AdoOrigen.RecordSource = "SELECT * From DetalleProrrateo WHERE (TipoProrrateo = 'ORIGEN')AND (NumeroProrrateo = " & NumeroProrrateo & ")"
                    Me.AdoOrigen.Refresh
                    
                    Me.AdoDestino.RecordSource = "SELECT * From DetalleProrrateo WHERE (TipoProrrateo = 'DESTINO')AND (NumeroProrrateo = " & NumeroProrrateo & ")"
                    Me.AdoDestino.Refresh
                 
                          FrmProrrateo.DTPFecha.Enabled = False
                          FrmProrrateo.CmbFin.Enabled = False
                          FrmProrrateo.CmbIni.Enabled = False
                          FrmProrrateo.Frame5.Enabled = False
                          FrmProrrateo.TDBProrrateo.Enabled = False
                 
                 End If
        
        
                    Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                              "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'ORIGEN')"
                    Me.AdoSuma.Refresh
                    If Not Me.AdoSuma.Recordset.EOF Then
                      Me.LblTotalOrigen.Caption = " Total   " & Format(Me.AdoSuma.Recordset("Importe"), "##,##0.00")
                    End If
                    
                    Me.AdoSuma.RecordSource = "SELECT NumeroProrrateo, MAX(CodCuenta) AS CodCuenta, MAX(TipoProrrateo) AS TipoProrrateo, MAX(Descripcion) AS Descripcion, SUM(MontoBase) AS MontoBase, SUM(Porciento) AS Porciento, SUM(Importe) AS Importe From DetalleProrrateo GROUP BY NumeroProrrateo, TipoProrrateo " & _
                                              "HAVING (NumeroProrrateo = " & NumeroProrrateo & ") AND (TipoProrrateo = 'DESTINO')"
                    Me.AdoSuma.Refresh
                    If Not Me.AdoSuma.Recordset.EOF Then
                      Me.LblDestino.Caption = "Total     " & Format(Me.AdoSuma.Recordset("Importe"), "##,##0.00")
                    End If
       
       Else
                    Me.AdoOrigen.RecordSource = "SELECT * From DetalleProrrateo WHERE (TipoProrrateo = 'ORIGEN')AND (NumeroProrrateo = -1)"
                    Me.AdoOrigen.Refresh
                    
                    Me.AdoDestino.RecordSource = "SELECT * From DetalleProrrateo WHERE (TipoProrrateo = 'DESTINO')AND (NumeroProrrateo = -1)"
                    Me.AdoDestino.Refresh
       
                          FrmProrrateo.DTPFecha.Enabled = True
                          FrmProrrateo.CmbFin.Enabled = True
                          FrmProrrateo.CmbIni.Enabled = True
                          FrmProrrateo.Frame5.Enabled = True
                          FrmProrrateo.TDBProrrateo.Enabled = True
       
       End If
End Sub
