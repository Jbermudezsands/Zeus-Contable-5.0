VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmImporta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importacion de Archivos desde Pacioli 3000"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12450
   ForeColor       =   &H8000000D&
   Icon            =   "FrmLeer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmLeer.frx":1E72
   ScaleHeight     =   5745
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAnexarZeus 
      Caption         =   "Anexar Zeus"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   16
      Top             =   5160
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblProgreso 
      Height          =   375
      Left            =   1920
      OleObjectBlob   =   "FrmLeer.frx":2B3C
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
   End
   Begin MSAdodcLib.Adodc AdoAnexar 
      Height          =   375
      Left            =   1560
      Top             =   7320
      Width           =   5055
      _ExtentX        =   8916
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
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   495
      Left            =   1560
      Top             =   7800
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   1560
      Top             =   7920
      Width           =   4695
      _ExtentX        =   8281
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
   Begin MSAdodcLib.Adodc AdoIndiceTransaccion 
      Height          =   375
      Left            =   1560
      Top             =   7920
      Width           =   4695
      _ExtentX        =   8281
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
      Left            =   1680
      Top             =   7920
      Width           =   4695
      _ExtentX        =   8281
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
   Begin MSAdodcLib.Adodc AdoTransacciones 
      Height          =   375
      Left            =   1920
      Top             =   7920
      Width           =   4695
      _ExtentX        =   8281
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
   Begin VB.CommandButton CmdAnexar 
      Caption         =   "Anexar Mov"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   375
      Left            =   10680
      OleObjectBlob   =   "FrmLeer.frx":2B9A
      TabIndex        =   13
      Top             =   4320
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   9120
      OleObjectBlob   =   "FrmLeer.frx":2C06
      TabIndex        =   12
      Top             =   4320
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblRegistros 
      Height          =   495
      Left            =   3120
      OleObjectBlob   =   "FrmLeer.frx":2C70
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   1920
      OleObjectBlob   =   "FrmLeer.frx":2CCE
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton CmdLeer 
      Caption         =   "Leer Registros"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblFuente 
      Height          =   375
      Left            =   3600
      OleObjectBlob   =   "FrmLeer.frx":2D56
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   2760
      OleObjectBlob   =   "FrmLeer.frx":2DB4
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblTransaccion 
      Height          =   255
      Left            =   4080
      OleObjectBlob   =   "FrmLeer.frx":2E1E
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   2760
      OleObjectBlob   =   "FrmLeer.frx":2E7C
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "FrmLeer.frx":2EF6
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin TrueOleDBGrid80.TDBGrid DBRegistro 
      Bindings        =   "FrmLeer.frx":2F5E
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4683
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
      Columns(1).Caption=   "Departamento"
      Columns(1).DataField=   "CodDepartamento"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Descripcion Movimiento"
      Columns(2).DataField=   "Descripcion"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Transaccion No."
      Columns(3).DataField=   "NTransaccion"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Fecha Transaccion"
      Columns(4).DataField=   "Fecha"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Debito"
      Columns(5).DataField=   "ImporteTransaccionDebito"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Credito"
      Columns(6).DataField=   "ImporteTransaccionCredito"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   1085
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   16315377
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
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
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
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
   Begin MSAdodcLib.Adodc AdoImporta 
      Height          =   375
      Left            =   1680
      Top             =   8280
      Width           =   4815
      _ExtentX        =   8493
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
      Caption         =   "AdoImporta"
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
      Left            =   1680
      Top             =   7920
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSAdodcLib.Adodc AdoConsecutivo 
      Height          =   375
      Left            =   1680
      Top             =   8280
      Width           =   4815
      _ExtentX        =   8493
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
      Left            =   1560
      Top             =   7920
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.Cns"
      Filter          =   "Cns"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   122683393
      CurrentDate     =   37714
   End
   Begin MSMask.MaskEdBox TxtDebito 
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "##,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtCredito 
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "##,##0.00"
      PromptChar      =   "_"
   End
   Begin XtremeSuiteControls.ProgressBar osProgress 
      Height          =   495
      Left            =   1800
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   7095
      _Version        =   786432
      _ExtentX        =   12515
      _ExtentY        =   873
      _StockProps     =   93
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   495
      Left            =   6600
      TabIndex        =   18
      Top             =   120
      Width           =   3855
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   873
      _StockProps     =   79
      ForeColor       =   -2147483630
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton OptPacioli 
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   150
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pacioli 3000"
         ForeColor       =   -2147483630
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptZeus 
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   150
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Zeus Contabilidad"
         ForeColor       =   -2147483630
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblAnulados2 
      Height          =   495
      Left            =   4440
      OleObjectBlob   =   "FrmLeer.frx":2F77
      TabIndex        =   22
      Top             =   3840
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblAnulados 
      Height          =   495
      Left            =   5400
      OleObjectBlob   =   "FrmLeer.frx":2FFB
      TabIndex        =   23
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "FrmImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LogErrores As String
Private HayErrores As Boolean

Private Sub AgregarError( _
        ByVal Fila As Long, _
        ByVal Movimiento As Variant, _
        ByVal Cuenta As String, _
        ByVal CentroCosto As String, _
        ByVal CodTercero As String, _
        ByVal Mensaje As String)

    HayErrores = True

    LogErrores = LogErrores & _
        String(60, "=") & vbCrLf & _
        "Fila          : " & Fila & vbCrLf & _
        "Movimiento    : " & Movimiento & vbCrLf & _
        "Cuenta        : " & Cuenta & vbCrLf & _
        "Centro Costo  : " & CentroCosto & vbCrLf & _
        "Cod. Tercero  : " & CodTercero & vbCrLf & _
        vbCrLf & _
        "ERROR:" & vbCrLf & _
        Mensaje & vbCrLf & _
        String(60, "=") & vbCrLf & vbCrLf

End Sub

Private Sub GuardarLogErrores(ByVal ArchivoExcel As String)

    Dim Archivo As Integer
    Dim NombreLog As String

    NombreLog = App.Path & "\Errores_" & _
                Format(Now, "yyyymmdd_hhnnss") & ".txt"

    Archivo = FreeFile

    Open NombreLog For Output As #Archivo

    Print #Archivo, "==============================================="
    Print #Archivo, " REPORTE DE ERRORES DE IMPORTACION"
    Print #Archivo, "==============================================="
    Print #Archivo, ""
    Print #Archivo, "Archivo : " & ArchivoExcel
    Print #Archivo, "Fecha   : " & Format(Now, "dd/MM/yyyy hh:mm:ss")
    Print #Archivo, ""
    Print #Archivo, LogErrores

    Close #Archivo

    MsgBox "Se encontraron errores." & vbCrLf & _
           "Revise el archivo:" & vbCrLf & _
           NombreLog, vbExclamation

End Sub

Private Sub IniciarProgreso(ByVal Total As Long, ByVal Mensaje As String)

    osProgress.Visible = True
    osProgress.Min = 0
    osProgress.Max = Total
    osProgress.Value = 0

    LblProgreso.Caption = Mensaje


    DoEvents

End Sub

Private Sub ActualizarProgreso(ByVal Actual As Long, _
                               ByVal Total As Long, _
                               ByVal Mensaje As String, _
                               Optional ByVal Errores As Long = 0)

    If Actual <= osProgress.Max Then
        osProgress.Value = Actual
    End If

    LblEstado.Caption = Mensaje & "Registro " & Format$(Actual, "#,##0") & _
                         " de " & Format$(Total, "#,##0") & " " & "Errores encontrados: " & _
                         Format$(Errores, "#,##0")


    DoEvents

End Sub

Private Sub FinalizarProgreso()

    osProgress.Value = osProgress.Max

    LblProgreso.Caption = "Proceso finalizado"

    DoEvents

End Sub


Private Sub CmdAnexar_Click()
Dim FechaArchivo As Date, NumeroPeriodo As Integer, NumeroMovimiento As Integer
Dim Fechas1 As String, Fechas2 As String, MovimientoArchivo As Double, Sql As String
Dim FechaGrabada As Date, MovimientoGrabado As Double, Fuente As String
Dim NombreCuenta As String, CodigoArchivo As String, Longitud As Double, i As Integer, J As Integer
Dim CodigoCuenta As String, mes As Double, CantidadRegistros As Double, ExisteCodigo As Boolean
Dim Directorio As String, Abrir As String, cadena As String

Directorio = App.Path + "\Cuentas.txt"

Me.LblProgreso.Visible = True
Me.osProgress.Visible = True
 '////////////////////////////////////////////////////////////////////////
 '//////BUSCO SI TODAS LAS CUENTAS DEL ARCHIVO EXISTEN EN LA BASE DE DATOS/
 '/////////////////////////////////////////////////////////////////////////
 

Me.AdoRegistros.Recordset.MoveLast
CantidadRegistros = AdoRegistros.Recordset.RecordCount
 
With Me.osProgress
 .Min = 0
 .Max = CantidadRegistros
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
       
          Me.LblProgreso.Caption = "Extraccion de la Cuenta: " & CodigoCuenta
          DoEvents
           
           Sql = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
                 "From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
           Me.AdoConsulta.RecordSource = Sql
           Me.AdoConsulta.Refresh
           If Me.AdoConsulta.Recordset.EOF Then
              cadena = CodigoCuenta
              Print #1, cadena
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
 
 
  
  Sql = "SELECT Control, IdRegistros, Fecha, NTransaccion, Fuente, CodCuenta, CodDepartamento, CodAcciones, ClaveProyecto, FacturaNumero, TipoMovimiento, " & _
        "RefCheque, Descripcion, FechaVencimiento, ImporteTransaccionDebito, ImporteTransaccionCredito, ImporteDescuento, ValorUnitario, TipoTransaccion, " & _
        "DebitoDolar , CreditoDolar, FechaDescuento From Registros ORDER BY Control "
        
 Me.AdoAnexar.RecordSource = Sql
 Me.AdoAnexar.Refresh
 
  Me.AdoAnexar.Recordset.MoveLast
  CantidadRegistros = Me.AdoAnexar.Recordset.RecordCount
  
  Me.AdoAnexar.Recordset.MoveFirst
  With Me.osProgress
 .Min = 0
 .Max = CantidadRegistros
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

  Fechas1 = CDate("1/" & Month(Me.AdoAnexar.Recordset("Fecha")) & "/" & Year(Me.AdoAnexar.Recordset("Fecha")))
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
     Sql = "SELECT     MAX(Control) AS Control, MAX(IdRegistros) AS IdRegistros, MAX(Fecha) AS Fecha, NTransaccion AS NTransaccion, SUM(ImporteTransaccionDebito) " & _
         "AS Debito, SUM(ImporteTransaccionCredito) AS Credito, SUM(ImporteTransaccionDebito) - SUM(ImporteTransaccionCredito) AS Diferencia " & _
         "From Registros GROUP BY NTransaccion " & _
         "HAVING  (MAX(Fecha) = '" & Format(FechaArchivo, "yyyymmdd") & "') AND (NTransaccion = " & MovimientoArchivo & " )"
     AdoSuma.RecordSource = Sql
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
     
   Sql = "SELECT  FechaTransaccion, NumeroMovimiento, DescripcionMovimiento, Nperiodo, Fuente, TipoMoneda " & _
         "From IndiceTransaccion " & _
         "Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroMovimiento & ") "
     
   Me.AdoIndiceTransaccion.RecordSource = Sql
   Me.AdoIndiceTransaccion.Refresh
   If Me.AdoIndiceTransaccion.Recordset.EOF Then

    NumeroMovimiento = NumeroMovimiento
    Me.AdoIndiceTransaccion.Recordset.AddNew
    Me.AdoIndiceTransaccion.Recordset("FechaTransaccion") = FechaArchivo
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
    Sql = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
        "From Cuentas WHERE (CodCuentas ='" & CodigoCuenta & "')"
        
    AdoCuenta.RecordSource = Sql
    AdoCuenta.Refresh
    If Not Me.AdoCuenta.Recordset.EOF Then
     If Not IsNull(Me.AdoCuenta.Recordset("DescripcionCuentas")) Then
      NombreCuenta = Me.AdoCuenta.Recordset("DescripcionCuentas")
     Else
      NombreCuenta = "CUENTA SIN DESCRIPCION?????"
     End If
    
    End If
     
     
     Me.LblProgreso.Caption = "Agregando la Cuenta: " & CodigoCuenta & " " & NombreCuenta
     DoEvents
     
    '//////////////////////////////////////////////////////////////////////
    '///////BUSCO LA TRANSACCION PARA ANEXAR O AGREGAR EL REGISTRO////////
    '///////////////////////////////////////////////////////////////////////
    
    Sql = "SELECT  CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, " & _
         "Debito, Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada, FechaDescuento, DescuentoDisponible, " & _
         "FechaVence From Transacciones " & _
         "Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroMovimiento & ")"
    
    Me.AdoTransacciones.RecordSource = Sql
    Me.AdoTransacciones.Refresh

      Me.AdoTransacciones.Recordset.AddNew
       Me.AdoTransacciones.Recordset("CodCuentas") = CodigoCuenta
       Me.AdoTransacciones.Recordset("FechaTransaccion") = FechaArchivo
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
End Sub

Private Sub CmdAnexarZeus_Click()
Dim FechaArchivo As Date, NumeroPeriodo As Integer, NumeroMovimiento As Integer
Dim Fechas1 As String, Fechas2 As String, MovimientoArchivo As Double, Sql As String
Dim FechaGrabada As Date, MovimientoGrabado As Double, Fuente As String
Dim NombreCuenta As String, CodigoArchivo As String, Longitud As Double, i As Integer, J As Integer
Dim CodigoCuenta As String, mes As Double, CantidadRegistros As Double, ExisteCodigo As Boolean
Dim Directorio As String, Abrir As String, cadena As String, VoucherNo As String

Directorio = App.Path + "\Cuentas.txt"

Me.LblProgreso.Visible = True
Me.osProgress.Visible = True
 '////////////////////////////////////////////////////////////////////////
 '//////BUSCO SI TODAS LAS CUENTAS DEL ARCHIVO EXISTEN EN LA BASE DE DATOS/
 '/////////////////////////////////////////////////////////////////////////
 

Me.AdoRegistros.Recordset.MoveLast
CantidadRegistros = AdoRegistros.Recordset.RecordCount

'If Not ValidarPeriodos(Conexion, Me.AdoRegistros.Recordset) Then Exit Sub
 
With Me.osProgress
 .Min = 0
 .Max = CantidadRegistros
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
       
          Me.LblProgreso.Caption = "Extraccion de la Cuenta: " & CodigoCuenta
          DoEvents
           
           Sql = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
                 "From Cuentas WHERE (CodCuentas = '" & CodigoCuenta & "')"
           Me.AdoConsulta.RecordSource = Sql
           Me.AdoConsulta.Refresh
           If Me.AdoConsulta.Recordset.EOF Then
              cadena = CodigoCuenta
              Print #1, cadena
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
 
 
  
  Sql = "SELECT Control, IdRegistros, Fecha, NTransaccion, Fuente, CodCuenta, CodDepartamento, CodAcciones, ClaveProyecto, FacturaNumero, TipoMovimiento, " & _
        "RefCheque, Descripcion, FechaVencimiento, ImporteTransaccionDebito, ImporteTransaccionCredito, ImporteDescuento, ValorUnitario, TipoTransaccion, " & _
        "DebitoDolar , CreditoDolar, FechaDescuento, Cod_Tercero From Registros ORDER BY NTransaccion, Control "
        
 Me.AdoAnexar.RecordSource = Sql
 Me.AdoAnexar.Refresh
 
  Me.AdoAnexar.Recordset.MoveLast
  CantidadRegistros = Me.AdoAnexar.Recordset.RecordCount
  
  Me.AdoAnexar.Recordset.MoveFirst
  With Me.osProgress
 .Min = 0
 .Max = CantidadRegistros
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

  Fechas1 = CDate("1/" & Month(Me.AdoAnexar.Recordset("Fecha")) & "/" & Year(Me.AdoAnexar.Recordset("Fecha")))
  mes = Month(Fechas1)
  Año = Year(Fechas1)
  Fechas2 = DateSerial(Año, mes + 1, 1 - 1)
  
  
 '////////////////////CODIGO RETIRADO 30/07/2026 //////////////////////////////////////////////////
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
     Sql = "SELECT     MAX(Control) AS Control, MAX(IdRegistros) AS IdRegistros, MAX(Fecha) AS Fecha, NTransaccion AS NTransaccion, SUM(ImporteTransaccionDebito) " & _
         "AS Debito, SUM(ImporteTransaccionCredito) AS Credito, SUM(ImporteTransaccionDebito) - SUM(ImporteTransaccionCredito) AS Diferencia " & _
         "From Registros GROUP BY NTransaccion " & _
         "HAVING  (MAX(Fecha) = '" & Format(FechaArchivo, "yyyymmdd") & "') AND (NTransaccion = " & MovimientoArchivo & " )"
     AdoSuma.RecordSource = Sql
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

   Sql = "SELECT  FechaTransaccion, NumeroMovimiento, DescripcionMovimiento, Nperiodo, Fuente, TipoMoneda " & _
         "From IndiceTransaccion " & _
         "Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroMovimiento & ") "

   Me.AdoIndiceTransaccion.RecordSource = Sql
   Me.AdoIndiceTransaccion.Refresh
   If Me.AdoIndiceTransaccion.Recordset.EOF Then
   
    Me.LblProgreso.Caption = "Agregando el Periodo: " & NumeroPeriodo
    DoEvents

    NumeroMovimiento = NumeroMovimiento
    Me.AdoIndiceTransaccion.Recordset.AddNew
    Me.AdoIndiceTransaccion.Recordset("FechaTransaccion") = FechaArchivo
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
    Sql = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta, CodGrupo, SaldoActual, TipoMoneda, KeyGrupo, DescripcionGrupo " & _
        "From Cuentas WHERE (CodCuentas ='" & CodigoCuenta & "')"
        
    AdoCuenta.RecordSource = Sql
    AdoCuenta.Refresh
    If Not Me.AdoCuenta.Recordset.EOF Then
     If Not IsNull(Me.AdoCuenta.Recordset("DescripcionCuentas")) Then
      NombreCuenta = Me.AdoCuenta.Recordset("DescripcionCuentas")
     Else
      NombreCuenta = "CUENTA SIN DESCRIPCION?????"
      Me.LblProgreso.Caption = "La cuenta no tiene Descripcion: " & NombreCuenta
      DoEvents
     End If
    
    End If
     
     
     Me.LblProgreso.Caption = "Agregando la Cuenta: " & CodigoCuenta & " " & NombreCuenta
     DoEvents
     
    '//////////////////////////////////////////////////////////////////////
    '///////BUSCO LA TRANSACCION PARA ANEXAR O AGREGAR EL REGISTRO////////
    '///////////////////////////////////////////////////////////////////////
    
    Sql = "SELECT  CodCuentas, FechaTransaccion, NPeriodo, NTransaccion, NumeroMovimiento, NombreCuenta, VoucherNo, DescripcionMovimiento, Clave, TCambio, " & _
         "Debito, Credito, FacturaNo, ChequeNo, Fuente, FechaTasas, Conciliada, ConciliacionProcesada, FechaDescuento, DescuentoDisponible, " & _
         "FechaVence, Cod_Tercero, Cod_Centro_Costo From Transacciones " & _
         "Where (NPeriodo = " & NumeroPeriodo & ") And (NumeroMovimiento = " & NumeroMovimiento & ")"
    
    Me.AdoTransacciones.RecordSource = Sql
    Me.AdoTransacciones.Refresh
    
      Me.LblProgreso.Caption = "Agregando Registro con la cuenta: " & CodigoCuenta
      DoEvents

      Me.AdoTransacciones.Recordset.AddNew
       Me.AdoTransacciones.Recordset("CodCuentas") = CodigoCuenta
       Me.AdoTransacciones.Recordset("VoucherNo") = Me.AdoAnexar.Recordset("CodDepartamento")
       Me.AdoTransacciones.Recordset("FechaTransaccion") = FechaArchivo
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
       
'       If Not IsNull(Me.AdoAnexar.Recordset("CodAcciones")) Then
'        Me.AdoTransacciones.Recordset("VoucherNo") = Me.AdoAnexar.Recordset("CodAcciones")
'       End If
       If Not IsNull(Me.AdoAnexar.Recordset("FacturaNumero")) Then
        Me.AdoTransacciones.Recordset("FacturaNo") = Me.AdoAnexar.Recordset("FacturaNumero")
       End If
       Me.AdoTransacciones.Recordset("ChequeNo") = Me.AdoAnexar.Recordset("RefCheque")
       Me.AdoTransacciones.Recordset("Fuente") = Fuente
       Me.AdoTransacciones.Recordset("FechaTasas") = FechaArchivo
       Me.AdoTransacciones.Recordset("Cod_Tercero") = Me.AdoAnexar.Recordset("Cod_Tercero")
       Me.AdoTransacciones.Recordset("Cod_Centro_Costo") = Me.AdoAnexar.Recordset("CodDepartamento")
       Me.AdoTransacciones.Recordset("Debito") = Me.AdoAnexar.Recordset("ImporteTransaccionDebito")
       Me.AdoTransacciones.Recordset("Credito") = Me.AdoAnexar.Recordset("ImporteTransaccionCredito")
          
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
End Sub
Private Sub AgregarRegistro( _
        ByVal Consecutivo As Long, _
        ByVal Fecha As Date, _
        ByVal NTransaccion As Double, _
        ByVal Fuente As String, _
        ByVal CodCuenta As String, _
        ByVal CodDepartamento As String, _
        ByVal CodAcciones As String, _
        ByVal ClaveProyecto As String, _
        ByVal FacturaNumero As String, _
        ByVal TipoMovimiento As String, _
        ByVal RefCheque As String, _
        ByVal Descripcion As String, _
        ByVal ImporteTransaccion As Double, _
        ByVal CodTercero As String)
        
        
          Me.AdoRegistros.Recordset.AddNew
          AdoRegistros.Recordset("IdRegistros") = Consecutivo
          AdoRegistros.Recordset("Fecha") = Fecha
          AdoRegistros.Recordset("NTransaccion") = NTransaccion
          AdoRegistros.Recordset("Fuente") = Fuente
          AdoRegistros.Recordset("CodCuenta") = CodCuenta
          AdoRegistros.Recordset("CodDepartamento") = CodDepartamento
          AdoRegistros.Recordset("CodAcciones") = CodAcciones
          AdoRegistros.Recordset("ClaveProyecto") = ClaveProyecto
          AdoRegistros.Recordset("FacturaNumero") = FacturaNumero
          AdoRegistros.Recordset("TipoMovimiento") = TipoMovimiento
          AdoRegistros.Recordset("RefCheque") = RefCheque
          AdoRegistros.Recordset("Descripcion") = Descripcion
          AdoRegistros.Recordset("Cod_Tercero") = CodTercero
          
          If TipoMovimiento = "01" Or TipoMovimiento = "02" Or TipoMovimiento = "03" Or TipoMovimiento = "04" Or TipoMovimiento = "10" Or TipoMovimiento = "11" Or TipoMovimiento = "12" Or TipoMovimiento = "13" Or TipoMovimiento = "14" Or TipoMovimiento = "20" Or TipoMovimiento = "21" Or TipoMovimiento = "23" Or TipoMovimiento = "27" Or TipoMovimiento = "31" Then
            AdoRegistros.Recordset("ImporteTransaccionDebito") = ImporteTransaccion
          ElseIf TipoMovimiento = "05" Or TipoMovimiento = "06" Or TipoMovimiento = "07" Or TipoMovimiento = "08" Or TipoMovimiento = "09" Or TipoMovimiento = "15" Or TipoMovimiento = "16" Or TipoMovimiento = "17" Or TipoMovimiento = "18" Or TipoMovimiento = "19" Or TipoMovimiento = "22" Or TipoMovimiento = "24" Or TipoMovimiento = "25" Or TipoMovimiento = "26" Or TipoMovimiento = "28" Or TipoMovimiento = "29" Or TipoMovimiento = "30" Then
            AdoRegistros.Recordset("ImporteTransaccionCredito") = ImporteTransaccion
          Else
              cadena = "Este Archivo Contienen un Codigo Nuevo" & vbLf
              cadena = cadena & "Llame a su Soporte Tecnico para Anexarlo" & vbLf
              cadena = cadena & "Al listado Permitido, para Importar"
            MsgBox cadena
            Exit Sub
          End If

          
          
          Me.AdoRegistros.Recordset.Update

End Sub




Private Sub LeerTxt(ByVal Directorio As String)

On Error GoTo TipoErrs

    Dim Contador As Integer
    Dim Campo As String, Fechas As String, FechaT1 As String, FechaT2 As String, FechaT3 As String
    Dim TotalDebito As Double, TotalCredito As Double, cadena As String
    Dim Consecutivo As Integer, NTransaccion As Double
    Dim Anulado As Integer, CodigoArchivo As String, CodigoCuenta As String
    Dim Fuente As String
    Dim CodCuenta As String, CodDepartamento As String
    Dim CodAcciones As String, ClaveProyecto As String, NFactura As String
    Dim Ttransaccion As String, ReferenciaCh As String
    Dim Descripcion As String, FechaDescuento As String
    Dim FechaVencimiento As String, ImporteDescuento As String
    Dim ImporteTransaccion As Double, ValorUnit As String
    Dim TipoTransaccion As String, Longitud As Double
    Dim Ultimo As Boolean
    Dim Encontrado As Boolean
    Dim Conexion As String
    Dim Buscado As Boolean
    Dim NumFecha1 As Long, NumFecha2 As Long
    Dim TipoMovimiento As String
    Dim Salir As Boolean, Continuar As Boolean

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
        'Me.DBRegistro.Columns(0).Visible = False
        'Me.DBRegistro.Columns(1).Visible = False
        'Me.DBRegistro.Columns(2).Visible = False
        'Me.DBRegistro.Columns(3).Visible = False
        'Me.DBRegistro.Columns(4).Visible = False
        'Me.DBRegistro.Columns(5).Caption = "Cuenta"
        'Me.DBRegistro.Columns(7).Caption = "Departamento"
        'Me.DBRegistro.Columns(8).Visible = False
        'Me.DBRegistro.Columns(9).Visible = False
        'Me.DBRegistro.Columns(11).Visible = False
        'Me.DBRegistro.Columns(13).Visible = False
        'Me.DBRegistro.Columns(14).Visible = False
        'Me.DBRegistro.Columns(15).Caption = "Debito"
        'Me.DBRegistro.Columns(16).Caption = "Credito"
        'Me.DBRegistro.Columns(17).Visible = False
        'Me.DBRegistro.Columns(18).Visible = False
        'Me.DBRegistro.Columns(19).Visible = False
        
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
        Line Input #1, cadena
         
        
         Fechas = Mid(cadena, 1, 8)
         FechaT2 = Mid(Fechas, 1, 2)
         FechaT1 = Mid(Fechas, 3, 2)
         FechaT3 = Mid(Fechas, 5, 4)
         Fecha = FechaT1 & "/" & FechaT2 + "/" + FechaT3
         
         
        
         NTransaccion = Mid(cadena, 9, 8)
        
         
         Fuente = Mid(cadena, 17, 4)
         CodCuenta = Mid(cadena, 21, 16)
         Me.Caption = "Procesando Cuenta " & CodCuenta
         CodDepartamento = Mid(cadena, 37, 3)
         CodAcciones = Mid(cadena, 40, 17)
         ClaveProyecto = Mid(cadena, 57, 15)
         NFactura = Mid(cadena, 72, 8)
         Ttransaccion = Mid(cadena, 80, 2)
         ReferenciaCh = Mid(cadena, 82, 6)
         Descripcion = Mid(cadena, 88, 35)
         
         
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
           If CodigoCuenta = "113-5-1-103-001-" Then
             Cod = 1
           End If
        
        
         '///////Convierto el formato de fecha////////////////
         
         
        ' FechaDescuento = Mid(Cadena, 124, 8)
        ' FechaVencimiento = Mid(Cadena, 132, 8)
        
        
         ImporteTransaccion = Mid(cadena, 140, 17)
        ' ImporteDescuento = Mid(cadena, 157, 17)
        ' ValorUnit = Mid(cadena, 174, 17)
        ' TipoTransaccion = Mid(cadena, 191, 2)
        
        If Not CodCuenta = "                " Then
         If ImporteTransaccion <> 0 Then


                AgregarRegistro _
                    Consecutivo, _
                    Fecha, _
                    NTransaccion, _
                    Fuente, _
                    CodCuenta, _
                    CodDepartamento, _
                    CodAcciones, _
                    ClaveProyecto, _
                    NFactura, _
                    Ttransaccion, _
                    ReferenciaCh, _
                    Descripcion, _
                    ImporteTransaccion, _
                    ""


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
        Me.AdoImporta.RecordSource = "SELECT Registros.Control, Registros.IdRegistros, Registros.Fecha, Registros.NTransaccion, Registros.Fuente, Registros.CodCuenta, Registros.Descripcion, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto, Registros.FacturaNumero, Registros.TipoMovimiento, Registros.RefCheque, Registros.FechaDescuento, Registros.FechaVencimiento, Registros.ImporteTransaccionDebito, Registros.ImporteTransaccionCredito, Registros.ImporteDescuento, Registros.ValorUnitario, Registros.TipoTransaccion From Registros Where (((Registros.IdRegistros) = " & Consecutivo & ")) ORDER BY NTransaccion, Control"
        Me.AdoImporta.Refresh
        ' Me.DBRegistro.Columns(0).Visible = False
        'Me.DBRegistro.Columns(1).Visible = False
        'Me.DBRegistro.Columns(2).Visible = False
        'Me.DBRegistro.Columns(3).Visible = False
        'Me.DBRegistro.Columns(4).Visible = False
        'Me.DBRegistro.Columns(5).Caption = "Cuenta"
        'Me.DBRegistro.Columns(7).Caption = "Departamento"
        'Me.DBRegistro.Columns(8).Visible = False
        'Me.DBRegistro.Columns(9).Visible = False
        'Me.DBRegistro.Columns(11).Visible = False
        'Me.DBRegistro.Columns(13).Visible = False
        'Me.DBRegistro.Columns(14).Visible = False
        'Me.DBRegistro.Columns(15).Caption = "Debito"
        'Me.DBRegistro.Columns(16).Caption = "Credito"
        'Me.DBRegistro.Columns(17).Visible = False
        'Me.DBRegistro.Columns(18).Visible = False
        'Me.DBRegistro.Columns(19).Visible = False
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
        Exit Sub
TipoErrs:
        MsgBox err.Description
        Salir = True
End Sub

Private Sub LeerExcel(ByVal Directorio As String)

    On Error GoTo ErrExcel

    Dim xlApp As Object
    Dim xlLibro As Object
    Dim xlHoja As Object

    Set xlApp = CreateObject("Excel.Application")
    Set xlLibro = xlApp.Workbooks.Open(Directorio)
    Set xlHoja = xlLibro.Worksheets(1)
    
    '=========================================
    ' Cargar catálogos en memoria
    '=========================================
      CargarCatalogos
      
    Dim Fila As Long
    Dim Fecha As Variant
    Dim Movimiento As Variant
    Dim Fuente As String
    Dim Cuenta As String
    Dim CentroCosto As String
    Dim Factura As String
    Dim Tipo As String
    Dim Referencia As String
    Dim Descripcion As String
    Dim Importe As Double
    Dim CodTercero As String
    Dim NombreTercero As String
    
    Dim Consecutivo As Integer
    Dim Contador As Integer
    Dim TotalDebito As Double
    Dim TotalCredito As Double
    Dim Anulado As Integer
    
    Me.AdoRegistros.Refresh

    Do While Not Me.AdoRegistros.Recordset.EOF
        Me.AdoRegistros.Recordset.Delete
        Me.AdoRegistros.Recordset.MoveNext
    Loop
    
    Consecutivo = 0
    Contador = 0
    
    Me.AdoImporta.RecordSource = "SELECT * FROM Registros WHERE IdRegistros = " & Consecutivo
    Me.AdoImporta.Refresh
    
    If Consecutivo = 0 Then
        Me.AdoConsecutivo.Recordset("ConsecutivoImporta") = _
            Me.AdoConsecutivo.Recordset("ConsecutivoImporta") + 1
    
        Me.AdoConsecutivo.Recordset.Update
    
        If Not IsNull(Me.AdoConsecutivo.Recordset("ConsecutivoImporta")) Then
            Consecutivo = Me.AdoConsecutivo.Recordset("ConsecutivoImporta")
        Else
            Consecutivo = 1
        End If
    End If
    
    
    Dim UltimaFila As Long
    
    UltimaFila = xlHoja.Cells(xlHoja.Rows.Count, 1).End(-4162).Row
    
    Fila = 2
    
    osProgress.Visible = True
    osProgress.Min = 0
    osProgress.Max = UltimaFila - 1    'Porque comienzas en la fila 2
    osProgress.Value = 0
    Me.LblProgreso.Visible = True
    Me.LblProgreso.Caption = "Importando transacciones..." & "Errores: 0"
'    LblDetalle.Caption = ""
'    LblErrores.Caption = "Errores: 0"

    DoEvents
    
    Do While Fila <= UltimaFila
    
        Fecha = xlHoja.Cells(Fila, 1).Value
'        Movimiento = xlHoja.Cells(Fila, 2).Value
        Movimiento = 1
        Fuente = Trim(CStr(xlHoja.Cells(Fila, 3).Value))
        Cuenta = Trim(CStr(xlHoja.Cells(Fila, 4).Value))
        Cuenta = Replace(Cuenta, " ", "")
        CentroCosto = Trim(CStr(xlHoja.Cells(Fila, 5).Value))
        Factura = Trim(CStr(xlHoja.Cells(Fila, 6).Value))
        Tipo = Trim(CStr(xlHoja.Cells(Fila, 7).Value))
        Referencia = Trim(CStr(xlHoja.Cells(Fila, 8).Value))
        Descripcion = Trim(CStr(xlHoja.Cells(Fila, 9).Value))
        
        '/////VALIDO QUE LA CUENTA NO ESTE EN BLANCO
   
        If Trim(CStr(xlHoja.Cells(Fila, 10).Value)) <> "" Then
            Importe = CDbl(xlHoja.Cells(Fila, 10).Value)
        Else
            Importe = 0
        End If
    
        CodTercero = Trim(CStr(xlHoja.Cells(Fila, 11).Value))
        CodTercero = Replace(CodTercero, " ", "")
'        NombreTercero = Trim(CStr(xlHoja.Cells(Fila, 12).Value))

If CodTercero = "5353" Then
 CodTercero = "5353"
End If
       

        If ValidarRegistroExcel(Fila, Fecha, Movimiento, Cuenta, CentroCosto, Tipo, Importe, CodTercero) Then


             AgregarRegistro _
                Consecutivo, _
                CDate(Fecha), _
                CDbl(Movimiento), _
                Fuente, _
                Cuenta, _
                CentroCosto, _
                "", _
                "", _
                Factura, _
                Tipo, _
                Referencia, _
                Descripcion, _
                Importe, _
                CodTercero
        
            Contador = Contador + 1
            Me.LblRegistros.Caption = Contador
            DoEvents
        Else
            Anulado = Anulado + 1
        End If
        
        If (Fila Mod 5) = 0 Then
        
            osProgress.Value = Fila - 1
        
            Me.LblRegistros.Caption = Contador
            Me.LblAnulados.Caption = Anulado
            Me.LblProgreso.Caption = "Validando registros..."
        
            DoEvents
        
        End If
            
        Fila = Fila + 1
    
   Loop
   
    If HayErrores Then
    
        Me.AdoRegistros.Refresh
    
        Do While Not Me.AdoRegistros.Recordset.EOF
            Me.AdoRegistros.Recordset.Delete
            Me.AdoRegistros.Recordset.MoveNext
        Loop
    
        GuardarLogErrores Directorio
    
        GoTo Salir
    
    End If
    
MsgBox "Se importaron " & Contador & " registros."

Salir:

    On Error Resume Next

    xlLibro.Close False
    xlApp.Quit

    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing
    
        Me.AdoImporta.RecordSource = "SELECT Registros.Control, Registros.IdRegistros, Registros.Fecha, Registros.NTransaccion, Registros.Fuente, Registros.CodCuenta, Registros.Descripcion, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto, Registros.FacturaNumero, Registros.TipoMovimiento, Registros.RefCheque, Registros.FechaDescuento, Registros.FechaVencimiento, Registros.ImporteTransaccionDebito, Registros.ImporteTransaccionCredito, Registros.ImporteDescuento, Registros.ValorUnitario, Registros.TipoTransaccion From Registros Where (((Registros.IdRegistros)=" & Consecutivo & ")) ORDER BY NTransaccion, Control"
        Me.AdoImporta.Refresh
    
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

    Exit Sub

ErrExcel:

    MsgBox err.Number & " - " & err.Description, vbCritical

    Resume Salir

End Sub
Public Sub CargarCatalogos()

    InicializarCatalogos

    CargarCatalogoCuentas
    CargarCatalogoCentrosCosto
    CargarCatalogoTerceros

End Sub


Private Function ValidarRegistroExcel( _
        ByVal Fila As Long, _
        ByVal Fecha As Variant, _
        ByVal Movimiento As Variant, _
        ByVal Cuenta As String, _
        ByVal CentroCosto As String, _
        ByVal Tipo As String, _
        ByVal Importe As Double, _
        ByVal CodTercero As String) As Boolean

    ValidarRegistroExcel = True
    


    '------------------------------
    ' Fecha
    '------------------------------
    If Not IsDate(Fecha) Then

        AgregarError _
            Fila, _
            Movimiento, _
            Cuenta, _
            CentroCosto, _
            CodTercero, _
            "La fecha no es válida."

        ValidarRegistroExcel = False

    End If

    '------------------------------
    ' Movimiento
    '------------------------------
    '------------------------------
    ' Movimiento
    '------------------------------
    If Trim(CStr(Movimiento)) = "" Then
    
        AgregarError _
            Fila, _
            Movimiento, _
            Cuenta, _
            CentroCosto, _
            CodTercero, _
            "El número de movimiento está vacío."
    
        ValidarRegistroExcel = False
    
    ElseIf Not IsNumeric(Movimiento) Then
    
        AgregarError _
            Fila, _
            Movimiento, _
            Cuenta, _
            CentroCosto, _
            CodTercero, _
            "El número de movimiento debe ser numérico. Valor recibido: '" & Movimiento & "'."
    
        ValidarRegistroExcel = False
    
    End If

    '------------------------------
    ' Cuenta obligatoria
    '------------------------------
    If Cuenta = "" Then

        AgregarError _
            Fila, _
            Movimiento, _
            Cuenta, _
            CentroCosto, _
            CodTercero, _
            "La cuenta contable está vacía."

        ValidarRegistroExcel = False

    End If

    '------------------------------
    ' Tipo de movimiento
    '------------------------------
    If Tipo <> "03" And Tipo <> "07" Then

        AgregarError _
            Fila, _
            Movimiento, _
            Cuenta, _
            CentroCosto, _
            CodTercero, _
            "Tipo de movimiento inválido (" & Tipo & "). Solo se permite 03 o 07."

        ValidarRegistroExcel = False

    End If

        '------------------------------
        ' Importe
        '------------------------------
        If Trim(CStr(Importe)) = "" Then
        
            AgregarError _
                Fila, _
                Movimiento, _
                Cuenta, _
                CentroCosto, _
                CodTercero, _
                "El importe está vacío."
        
            ValidarRegistroExcel = False
        
        ElseIf Not IsNumeric(Importe) Then
        
            AgregarError _
                Fila, _
                Movimiento, _
                Cuenta, _
                CentroCosto, _
                CodTercero, _
                "El importe debe ser numérico. Valor recibido: '" & Importe & "'."
        
            ValidarRegistroExcel = False
        
        ElseIf CDbl(Importe) <= 0 Then
        
            AgregarError _
                Fila, _
                Movimiento, _
                Cuenta, _
                CentroCosto, _
                CodTercero, _
                "El importe debe ser mayor que cero."
        
            ValidarRegistroExcel = False
        
        End If

    '------------------------------
    ' La cuenta debe existir
    '------------------------------
    If Cuenta <> "" Then

        If Not CuentaExiste(Cuenta) Then

            AgregarError _
                Fila, _
                Movimiento, _
                Cuenta, _
                CentroCosto, _
                CodTercero, _
                "La cuenta contable '" & Cuenta & "' no existe."

            ValidarRegistroExcel = False

        End If

    End If

    '------------------------------
    ' El centro de costo debe existir
    '------------------------------
    If Not CentroCostoExiste(CentroCosto) Then

        AgregarError _
            Fila, _
            Movimiento, _
            Cuenta, _
            CentroCosto, _
            CodTercero, _
            "El centro de costo '" & CentroCosto & "' no existe."

        ValidarRegistroExcel = False

    End If

    '------------------------------
    ' El tercero debe existir
    '------------------------------
    If Not TerceroExiste(CodTercero) Then

        AgregarError _
            Fila, _
            Movimiento, _
            Cuenta, _
            CentroCosto, _
            CodTercero, _
            "El código de tercero '" & CodTercero & "' no existe."

        ValidarRegistroExcel = False

    End If

End Function

Public Function CuentaExiste(ByVal Codigo As String) As Boolean

    CuentaExiste = DicCuentas.Exists(Trim(Codigo))

End Function

Public Function CentroCostoExiste(ByVal Codigo As String) As Boolean

    If Trim(Codigo) = "" Then
        CentroCostoExiste = True
    Else
        CentroCostoExiste = DicCentrosCosto.Exists(Trim(Codigo))
    End If

End Function

Public Function TerceroExiste(ByVal Codigo As String) As Boolean

    If Trim(Codigo) = "" Then
        TerceroExiste = True
    Else
        TerceroExiste = DicTerceros.Exists(Trim(Codigo))
    End If

End Function




Private Sub CmdLeer_Click()

'On Error GoTo TipoErrs

 Dim Directorio As String



    With Me.CommonDialog1
        .CancelError = False
        .Filter = "Archivos de Importación (*.cns;*.xls;*.xlsx)|*.cns;*.xls;*.xlsx|Archivos Texto (*.cns)|*.cns|Archivos Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Todos los archivos (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen
    End With

    Directorio = Me.CommonDialog1.FileName
    
    If Trim(Directorio) = "" Then Exit Sub
    
    
        If LCase$(right$(Directorio, 4)) = ".cns" Then
            LeerTxt Directorio
        Else
            LeerExcel Directorio
       
        End If
        
        





End Sub

Private Sub ADOImporta_Reposition()

If Not AdoImporta.Recordset.EOF Then
' Me.DTPicker1.Value = AdoImporta.Recordset.Fecha
' Me.LblFuente.Caption = Me.AdoImporta.Recordset.Fuente
' Me.LblTransaccion.Caption = Me.AdoImporta.Recordset.NTransaccion
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub


Private Sub DBRegistro_AfterUpdate()
Me.AdoSuma.RecordSource = "SELECT Sum(Registros.ImporteTransaccionDebito) AS TotalDebito, Sum(Registros.ImporteTransaccionCredito) AS TotalCredito From Registros"
Me.AdoSuma.Refresh
TotalCredito = Format(AdoSuma.Recordset("TotalCredito"), "##,##0.00")
Me.TxtCredito.Text = TotalCredito
TotalDebito = Format(AdoSuma.Recordset("TotalDebito"), "##,##0.00")
Me.TxtDebito.Text = TotalDebito
End Sub

Private Sub Form_Load()
Salir = True
MDIPrimero.Skin1.ApplySkin hWnd
 Me.DBRegistro.EvenRowStyle.BackColor = &H80FFFF
 Me.DBRegistro.OddRowStyle.BackColor = &HC0FFFF
 Me.DBRegistro.AlternatingRowStyle = True
 
With Me.AdoCuenta
   .ConnectionString = Conexion
End With

With Me.AdoAnexar
   .ConnectionString = Conexion
End With
 

With Me.AdoConsecutivo
   .ConnectionString = Conexion
   .RecordSource = "NConsecutivos"
   .Refresh
End With

With Me.AdoImporta
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.AdoIndiceTransaccion
   .ConnectionString = Conexion
End With

With Me.AdoPeriodos
   .ConnectionString = Conexion
End With

With Me.AdoTransacciones
   .ConnectionString = Conexion
End With

With Me.AdoRegistros
   .ConnectionString = Conexion
   .RecordSource = "Registros"
   .Refresh
End With

With Me.AdoSuma
   .ConnectionString = Conexion
End With


Consecutivo = 0
Me.AdoImporta.RecordSource = "SELECT Registros.Control, Registros.IdRegistros, Registros.Fecha, Registros.NTransaccion, Registros.Fuente, Registros.CodCuenta, Registros.Descripcion, Registros.CodDepartamento, Registros.CodAcciones, Registros.ClaveProyecto, Registros.FacturaNumero, Registros.TipoMovimiento, Registros.RefCheque, Registros.FechaDescuento, Registros.FechaVencimiento, Registros.ImporteTransaccionDebito, Registros.ImporteTransaccionCredito, Registros.ImporteDescuento, Registros.ValorUnitario, Registros.TipoTransaccion From Registros Where (((Registros.IdRegistros) = " & Consecutivo & ")) "
Me.AdoImporta.Refresh
'Me.DBRegistro.Columns(0).Visible = False
'Me.DBRegistro.Columns(1).Visible = False
'Me.DBRegistro.Columns(2).Visible = False
'Me.DBRegistro.Columns(3).Visible = False
'Me.DBRegistro.Columns(4).Visible = False
'Me.DBRegistro.Columns(5).Caption = "Cuenta"
'Me.DBRegistro.Columns(7).Caption = "Departamento"
'Me.DBRegistro.Columns(8).Visible = False
'Me.DBRegistro.Columns(9).Visible = False
'Me.DBRegistro.Columns(11).Visible = False
'Me.DBRegistro.Columns(13).Visible = False
'Me.DBRegistro.Columns(14).Visible = False
'Me.DBRegistro.Columns(15).Caption = "Debito"
'Me.DBRegistro.Columns(16).Caption = "Credito"
'Me.DBRegistro.Columns(17).Visible = False
'Me.DBRegistro.Columns(18).Visible = False
'Me.DBRegistro.Columns(19).Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Salir = False Then
 Cancel = 1
End If
End Sub

Private Sub TDBGrid1_Click()

End Sub

Private Sub OptPacioli_Click()
  If Me.OptPacioli.Value = True Then
    Me.CmdAnexar.Visible = True
    Me.CmdAnexarZeus.Visible = False
  End If
End Sub

Private Sub OptZeus_Click()
  If Me.OptZeus.Value = True Then
    Me.CmdAnexar.Visible = False
    Me.CmdAnexarZeus.Visible = True
  End If
End Sub
