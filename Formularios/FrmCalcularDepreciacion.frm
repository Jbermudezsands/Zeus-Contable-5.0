VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmCalcularDepreciacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcular Depreciacion"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12120
   Icon            =   "FrmCalcularDepreciacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdProcesar 
      Cancel          =   -1  'True
      Caption         =   "Procesar"
      Height          =   375
      Left            =   10680
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin XtremeSuiteControls.ProgressBar BarCalcular 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   11895
      _Version        =   786432
      _ExtentX        =   20981
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FrmCalcularDepreciacion.frx":57E2
      TabIndex        =   7
      Top             =   240
      Width           =   9495
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton CmdCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   10680
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   10200
      OleObjectBlob   =   "FrmCalcularDepreciacion.frx":599A
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   10440
      OleObjectBlob   =   "FrmCalcularDepreciacion.frx":5A3E
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DCmbCodigo 
      Height          =   315
      Left            =   12240
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   330
      Left            =   10920
      Top             =   7800
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc AdoTasas 
      Height          =   330
      Left            =   10920
      Top             =   7485
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc AdoIndice 
      Height          =   330
      Left            =   10920
      Top             =   7155
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc AdoPeriodos 
      Height          =   330
      Left            =   10920
      Top             =   6840
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc AdoActivoFijo 
      Height          =   330
      Left            =   10920
      Top             =   6525
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoActivoFijo"
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
      Height          =   330
      Left            =   10920
      Top             =   6195
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   10920
      Top             =   5880
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSComCtl2.DTPicker TxtFecha 
      Height          =   285
      Left            =   10560
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   78381057
      CurrentDate     =   38008
   End
   Begin TrueOleDBGrid80.TDBGrid DataGrid2 
      Bindings        =   "FrmCalcularDepreciacion.frx":5ABC
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5741
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
      Splits(0).Caption=   "Listado de Activos"
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
      PictureModifiedRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureModifiedRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
      PictureModifiedRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAP//"
      PictureModifiedRow(2)=   "/////////////////////////////////////////8bHxgAAxsfGAAAAhIaEAP//AP//AP//AP//"
      PictureModifiedRow(3)=   "AP//AP//AP//AP//AP//////xsfGAADGx8YAAACEhoQA//8A//8A//8A//8A//8A//8A//8A//8A"
      PictureModifiedRow(4)=   "///////Gx8YAAMbHxgAAAISGhAD//wD//wD//wD//wD//wD//wD//wD//wD//////8bHxgAAxsfG"
      PictureModifiedRow(5)=   "AAAAhIaEAP//AP//AP//AP//AP//AP//AP//AP//AP//////xsfGAADGx8YAAACEhoQA//8A//8A"
      PictureModifiedRow(6)=   "//8A//8A//8A//8A//8A//8A///////Gx8YAAMbHxgAAAISGhISGhISGhISGhISGhISGhISGhISG"
      PictureModifiedRow(7)=   "hISGhISGhP///8bHxgAAxsfGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxsfG"
      PictureModifiedRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
      PictureModifiedRow.vt=   9
      PictureAddnewRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureAddnewRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
      PictureAddnewRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAAAA"
      PictureAddnewRow(2)=   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMbHxgAAxsfG////hIaEhIaEhIaEhIaEhIaE"
      PictureAddnewRow(3)=   "hIaEhIaEhIaEhIaEhIaEAAAAxsfGAADGx8b///8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP+E"
      PictureAddnewRow(4)=   "hoQAAADGx8YAAMbHxv///wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/4SGhAAAAMbHxgAAxsfG"
      PictureAddnewRow(5)=   "////AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/hIaEAAAAxsfGAADGx8b///8AAP8AAP8AAP8A"
      PictureAddnewRow(6)=   "AP8AAP8AAP8AAP8AAP8AAP+EhoQAAADGx8YAAMbHxv///wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA"
      PictureAddnewRow(7)=   "/wAA/4SGhAAAAMbHxgAAxsfG////////////////////////////////////////////AAAAxsfG"
      PictureAddnewRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
      PictureAddnewRow.vt=   9
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HBFD6DD&,.fgcolor=&H800000&"
      _StyleDefs(24)  =   ":id=22,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=22,.fontname=Lucida Calligraphy"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HBFD6DD&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(27)  =   ":id=14,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin VB.Label LblNombre 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   9735
   End
End
Attribute VB_Name = "FrmCalcularDepreciacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset, rsConexion As New ADODB.Recordset


Private Sub CmdCancelar_Click()
Unload Me
End Sub
Private Sub NumeroMovimiento()
   Me.CmdCalcular.Enabled = True
    'Me.DBGTransacciones.Enabled = True
    mes = Month(Me.TxtFecha.Value)
    Año = Year(Me.TxtFecha.Value)
    FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
    FechaFin = DateSerial(Año, mes + 1, 1 - 1)
    NumFecha1 = FechaIni
    NumFecha2 = FechaFin
    
    Me.DCmbCodigo.Enabled = True
    Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
    Me.AdoConsulta.Refresh
    If Not AdoConsulta.Recordset.EOF Then
     NumeroPeriodo = AdoConsulta.Recordset!NPeriodo
     NumeroTransaccion = AdoConsulta.Recordset!NTransacciones
     EstadoPeriodo = AdoConsulta.Recordset!EstadoPeriodo
     If EstadoPeriodo = "B" Then
      MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
      Me.TxtFecha.SetFocus
      
      Exit Sub
     ElseIf EstadoPeriodo = "C" Then
     MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
     Me.TxtFecha.SetFocus
     TxtFecha.Enabled = True
     
     Exit Sub
     Else
      'Me.DBGTransacciones.Enabled = True
     End If
    Else
      MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
      'Me.DBGTransacciones.Enabled = False
      TxtFecha.Enabled = True
      
      Exit Sub
    End If



End Sub


Private Sub CmdCalcular_Click()
Dim CanRegistros As Integer, i As Integer
Dim ValorOriginal As Double, ValorRescate As Double, VidaEstimada As Double
Dim Depreciacion As Double, TotalDepreciacion As Double, NumFecha As Long
Dim CuentaDepreciacion As String, CuentaGasto As String, Tasas As Double
Dim TipoCuentaGastos As String, TipoCuentaDepreciacion As String
Dim Debito As Double, Credito As Double, CuentaValorOriginal As String
On Error GoTo TipoErrs

 
 Tasas = BuscaTasaCambio(Me.TxtFecha.Value)
 
 If Tasas = 0 Then
     MsgBox "No existe la Tasa de Cambio para la fecha del Movimiento", vbCritical, "sistema Contable"
     Exit Sub
 End If
 
 sqlconsulta = "SELECT CatalogoActivoFijo.* From CatalogoActivoFijo Where (DatoAlta = 1)"
 Me.AdoActivoFijo.RecordSource = sqlconsulta
 Me.AdoActivoFijo.Refresh
 Me.AdoActivoFijo.Recordset.MoveLast
 CanRegistros = Me.AdoActivoFijo.Recordset.RecordCount
 Me.AdoActivoFijo.Recordset.MoveFirst
 MsgBox ("Se Procesarán " & CanRegistros & " Activo Fijos")
 With BarCalcular
 .Min = 0
 .Max = CanRegistros
 .Value = 0
 i = 1

Me.AdoActivoFijo.Refresh
  
     Do While Not Me.AdoActivoFijo.Recordset.EOF

    
         
    
                        If Not IsNull(Me.AdoActivoFijo.Recordset("costogen")) Then
                          ValorOriginal = Me.AdoActivoFijo.Recordset("costogen")
                        End If
                        
                       
                        If Not IsNull(Me.AdoActivoFijo.Recordset("ValorRescate")) Then
                          ValorRescate = Me.AdoActivoFijo.Recordset("ValorRescate")
                        End If
                        
                        If Not IsNull(Me.AdoActivoFijo.Recordset("ValorEstimadoMeses")) Then
                           VidaEstimada = Me.AdoActivoFijo.Recordset("ValorEstimadoMeses")
                        End If
               
                
                
                
                        If Not VidaEstimada = 0 Then
                          Depreciacion = (ValorOriginal - ValorRescate) / VidaEstimada
                        Else
                          Depreciacion = 0
                        End If

             
             Me.AdoActivoFijo.Recordset("DepreciacionAcumulada") = Format(Depreciacion, "##,##0.00")
             Me.AdoActivoFijo.Recordset.Update
    
                
        TotalDepreciacion = Depreciacion + TotalDepreciacion
        Me.AdoActivoFijo.Recordset.MoveNext
      
      
         i = i + 1
        .Value = .Value + 1
      Loop
      

 
 End With
 
cargarcatalogoAF
 
 
 MsgBox "El Proceso ha Finalizado Correctamente", vbInformation, "Sistema Contable"
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub CmdProcesar_Click()
Dim CanRegistros As Integer, i As Integer
Dim ValorOriginal As Double, ValorRescate As Double, VidaEstimada As Double, DescripcionCtaGasto As String, DescripcionCtaDep As String
Dim Depreciacion As Double, TotalDepreciacion As Double, NumFecha As Long
Dim CuentaDepreciacion As String, CuentaGasto As String, Tasas As Double
Dim TipoCuentaGastos As String, TipoCuentaDepreciacion As String
Dim Debito As Double, Credito As Double, CuentaValorOriginal As String
On Error GoTo TipoErrs

Me.CmdProcesar.Enabled = False

NumeroMovimiento
NumeroTransaccion = NumeroTransaccion + 1
 'If Me.DCmbCodigo.Text = "" Then
  'MsgBox "Se necesita la cuenta de Depreciacion", vbCritical, "sistema Contable"
  'Exit Sub
 'End If
 Tasas = BuscaTasaCambio(Me.TxtFecha.Value)
 
 If Tasas = 0 Then
     MsgBox "No existe la Tasa de Cambio para la fecha del Movimiento", vbCritical, "sistema Contable"
     Exit Sub
 End If
 
 
 sqlconsulta = "SELECT CatalogoActivoFijo.* From CatalogoActivoFijo Where (DatoAlta = 1)"
 Me.AdoActivoFijo.RecordSource = sqlconsulta
 Me.AdoActivoFijo.Refresh
 Me.AdoActivoFijo.Recordset.MoveLast
 CanRegistros = Me.AdoActivoFijo.Recordset.RecordCount
 Me.AdoActivoFijo.Recordset.MoveFirst
 MsgBox ("Se Procesarán " & CanRegistros & " Activo Fijos")
 With BarCalcular
         .Min = 0
         .Max = CanRegistros
         .Value = 0
         i = 1
         
                    'Valido que no hayan duplicados de indices de transacciones JP
                    AdoConsulta.RecordSource = "Select * from IndiceTransaccion where FechaTransaccion='" & Format(TxtFecha, "yyyymmdd") & "' and NumeroMovimiento=" & Str(NumeroTransaccion)
                    AdoConsulta.Refresh
                    If Not AdoConsulta.Recordset.EOF Then
                        MsgBox "Ya se ha hecho esta Transacción Anteriormente", vbInformation
                        Exit Sub
                    End If
           'Agrego el indice
                  Me.AdoIndice.Recordset.AddNew
                  Me.AdoIndice.Recordset!FechaTransaccion = Me.TxtFecha.Value
                  Me.AdoIndice.Recordset!NumeroMovimiento = NumeroTransaccion
                  Me.AdoIndice.Recordset!DescripcionMovimiento = "Calculo Automatico Depreciacion"
                  Me.AdoIndice.Recordset!Fuente = "DEPRECIACION"
                  Me.AdoIndice.Recordset!NPeriodo = NumeroPeriodo
                   Me.AdoIndice.Recordset!TipoMoneda = "Córdobas"
                  Me.AdoIndice.Recordset.Update
                  
         Me.AdoActivoFijo.Refresh
         Do While Not Me.AdoActivoFijo.Recordset.EOF
         
                      .Value = i
'                     If Not IsNull(Me.AdoActivoFijo.Recordset("CodCuena")) Then
'                       CuentaValorOriginal = Me.AdoActivoFijo.Recordset("CodCuena")
'                     End If
                     
                     If Not IsNull(Me.AdoActivoFijo.Recordset("CuentaDepreciacion")) Then
                      CuentaDepreciacion = Me.AdoActivoFijo.Recordset("CuentaDepreciacion")
                     Else
                      CuentaDepreciacion = "Nulo"
                     End If
                     
                     If Not IsNull(Me.AdoActivoFijo.Recordset("CuentaGastos")) Then
                       CuentaGasto = Me.AdoActivoFijo.Recordset("CuentaGastos")
                     Else
                       CuentaGasto = "Nulo"
                     End If
                     
                     
                   'Busco si existe la Cuenta
                   DescripcionCtaGasto = "Nulo"
                   Me.AdoConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda From Cuentas WHERE (((Cuentas.CodCuentas)='" & CuentaGasto & "'))"
                   Me.AdoConsulta.Refresh
                    '/////Busco la Cuenta de Gastos////////////////////////////
                      If Not AdoConsulta.Recordset.EOF Then
                        DescripcionCtaGasto = Me.AdoConsulta.Recordset("DescripcionCuentas")
                      End If
                  
                   DescripcionCtaDep = "Nulo"
'                   TipoCuentaGastos = Me.AdoConsulta.Recordset!TipoMoneda
                   Me.AdoConsulta.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.DescripcionCuentas, Cuentas.TipoCuenta, Cuentas.TipoMoneda From Cuentas WHERE (((Cuentas.CodCuentas)='" & CuentaDepreciacion & "'))"
                   Me.AdoConsulta.Refresh
                        '///////////Busco la cuenta de la Depreciacion//////////
                           If Not AdoConsulta.Recordset.EOF Then
                              DescripcionCtaDep = Me.AdoConsulta.Recordset("DescripcionCuentas")
                           
                           End If
                            
                                    ValorOriginal = Me.AdoActivoFijo.Recordset("costogen")
                                    ValorRescate = Me.AdoActivoFijo.Recordset("ValorRescate")
                                    VidaEstimada = Me.AdoActivoFijo.Recordset("ValorEstimadoMeses")
                           

                                    If Not VidaEstimada = 0 Then
                                      Depreciacion = (ValorOriginal - ValorRescate) / VidaEstimada
                                    Else
                                      Depreciacion = 0
                                    End If
                                    
                                     Me.AdoTransacciones.Recordset.AddNew
                                     Me.AdoTransacciones.Recordset!CodCuentas = CuentaGasto
                                     Me.AdoTransacciones.Recordset!FechaTransaccion = Me.TxtFecha.Value
                                     Me.AdoTransacciones.Recordset!NPeriodo = NumeroPeriodo
                                     Me.AdoTransacciones.Recordset!NumeroMovimiento = NumeroTransaccion
                                     Me.AdoTransacciones.Recordset!NombreCuenta = DescripcionCtaGasto
                                     Me.AdoTransacciones.Recordset!DescripcionMovimiento = "Movimiento de Depreciacion"
                                     Me.AdoTransacciones.Recordset!Clave = "Debito"
                                     Me.AdoTransacciones.Recordset!Debito = Depreciacion
                                     Me.AdoTransacciones.Recordset!Fuente = "DEPRECIACION"
                                     Me.AdoTransacciones.Recordset!TCambio = 1
                                    Me.AdoTransacciones.Recordset.Update
                                    
                                    
                                     
                                     Me.AdoTransacciones.Recordset.AddNew
                                     Me.AdoTransacciones.Recordset!CodCuentas = CuentaDepreciacion
                                     Me.AdoTransacciones.Recordset!FechaTransaccion = Me.TxtFecha.Value
                                     Me.AdoTransacciones.Recordset!NPeriodo = NumeroPeriodo
                                     Me.AdoTransacciones.Recordset!NumeroMovimiento = NumeroTransaccion
                                     Me.AdoTransacciones.Recordset!NombreCuenta = DescripcionCtaDep
                                     Me.AdoTransacciones.Recordset!DescripcionMovimiento = "Movimiento de Depreciacion"
                                     Me.AdoTransacciones.Recordset!Clave = "Credito"
                                     Me.AdoTransacciones.Recordset!Credito = Depreciacion
                                     Me.AdoTransacciones.Recordset!Fuente = "DEPRECIACION"
                                     Me.AdoTransacciones.Recordset!TCambio = 1

                                    Me.AdoTransacciones.Recordset.Update
                          
                          

                          
                            
                            TotalDepreciacion = Depreciacion + TotalDepreciacion
                           

                    Me.AdoActivoFijo.Recordset.MoveNext
                  
                  
                  i = i + 1
          Loop
          
        '   Me.AdoTransacciones.Recordset.AddNew
        '   Me.AdoTransacciones.Recordset.CodCuentas = Me.DCmbCodigo.Text
        '   Me.AdoTransacciones.Recordset("FechaTransaccion") = Me.TxtFecha.Value
        '   Me.AdoTransacciones.Recordset.NPeriodo = NumeroPeriodo
        '   Me.AdoTransacciones.Recordset("NumeroMovimiento") = NumeroTransaccion
        '   Me.AdoTransacciones.Recordset.NombreCuenta = "Calculo Automatico Depreciacion"
        '   Me.AdoTransacciones.Recordset.DescripcionMovimiento = "Movimiento de Depreciacion"
        '   Me.AdoTransacciones.Recordset.Clave = "Debito"
        '   Me.AdoTransacciones.Recordset.Debito = TotalDepreciacion
        '   Me.AdoTransacciones.Recordset("Fuente") = "DEPRECIACION"
        '   Me.AdoTransacciones.Recordset.Tcambio = 1
        '  Me.AdoTransacciones.Recordset.Update
         
         'edito periodos
         mes = Month(Me.TxtFecha.Value)
         Año = Year(Me.TxtFecha.Value)
         FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
         FechaFin = DateSerial(Año, mes + 1, 1 - 1)
         NumFecha1 = FechaIni
         NumFecha2 = FechaFin
         
         Me.DCmbCodigo.Enabled = True
         Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
         Me.AdoConsulta.Refresh
         If Not AdoConsulta.Recordset.EOF Then
          'AdoConsulta.Recordset.Edit
           AdoConsulta.Recordset!NTransacciones = NumeroTransaccion
          AdoConsulta.Recordset.Update
         End If
 
 End With
 
 MsgBox "El Proceso ha Finalizado Correctamente", vbInformation, "Sistema Contable"
 
 Me.CmdProcesar.Enabled = True
Exit Sub
TipoErrs:
MsgBox err.Description
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DCmbCodigo_Change()
On Error GoTo TipoErrs
Criterio = "CodCuentas='" & Me.DCmbCodigo.Text & "'"
Me.AdoCuentas.Recordset.Find (Criterio)
If Not AdoCuentas.Recordset.EOF Then
 If Not Me.AdoCuentas.Recordset.EOF Then
   Me.LblNombre.Caption = Me.AdoCuentas.Recordset!DescripcionCuentas
   Me.CmdCalcular.Enabled = True
 End If
End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub Form_Activate()
'Me.TxtFecha.Value = Format(FechaSistema, "dd/mm/yyyy")

Me.AdoCuentas.RecordSource = "SELECT Cuentas.CodCuentas, Cuentas.TipoCuenta, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.TipoCuenta) = 'Gastos')) ORDER BY Cuentas.CodCuentas"
Me.AdoCuentas.Refresh
'Me.DCmbCodigo.ListField = "CodCuentas"
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
MDIPrimero.Skin1.ApplySkin hWnd
With Me.AdoCuentas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Cuentas"
   .Refresh
End With

With Me.AdoActivoFijo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from ActivoFijo"
   .Refresh
End With

With Me.AdoConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.AdoIndice
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from IndiceTransaccion"
   .Refresh
End With

With Me.AdoPeriodos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Periodos"
   .Refresh
End With

With Me.AdoTasas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Tasas"
   .Refresh
End With

With Me.AdoTransacciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Select * from Transacciones"
   .Refresh
End With



Me.TxtFecha.Value = Format(Now, "dd/mm/yyyy")
AdoCuentas.ConnectionString = Conexion
Me.AdoCuentas.RecordSource = "SELECT top 10 Cuentas.CodCuentas, Cuentas.TipoCuenta, Cuentas.DescripcionCuentas From Cuentas Where (((Cuentas.TipoCuenta) = 'Cuentas de Gastos')) ORDER BY Cuentas.CodCuentas"
Me.AdoCuentas.Refresh
LlenarDataCombos AdoCuentas, DCmbCodigo, "CodCuentas", "CodCuentas"
'Me.DCmbCodigo.ListField = "CodCuentas"

 cargarcatalogoAF

' Me.DataGrid2.DataSource = rs
 Me.DataGrid2.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DataGrid2.OddRowStyle.BackColor = &H80000005
 Me.DataGrid2.AlternatingRowStyle = True
 Me.BackColor = RGB(216, 228, 248)


Exit Sub
TipoErrs:
ControlErrores
End Sub
Private Sub cargarcatalogoAF()
  Dim sqlconsulta As String

    
     
    


'        Adodc3.RecordSource = "Select idReg as No, unidad as Unidad, DescripcionAF AS Activo_Fijo, Serie from dbo.CatalogoActivoFijo WHERE(DatoAlta = 1)"
    sqlconsulta = "SELECT idReg AS No, Unidad, DescripcionAF AS Activo_Fijo, Serie, CuentaGastos, CuentaDepreciacion, DepreciacionAcumulada AS Depreciacion From CatalogoActivoFijo Where (DatoAlta = 1)"

    Me.AdoActivoFijo.ConnectionString = Conexion
    Me.AdoActivoFijo.RecordSource = sqlconsulta
    Me.AdoActivoFijo.Refresh
    
    Me.DataGrid2.DataSource = Me.AdoActivoFijo
    Me.DataGrid2.Columns(2).Caption = "Voucher/Dpto"
    Me.DataGrid2.Columns(2).Width = 1100
'    Adodc3.Refresh
End Sub



Private Sub TxtFecha_Change()
On Error GoTo TipoErrs
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumeroMovimiento


Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFecha_GotFocus()
On Error GoTo TipoErrs
' 'Me.DBGTransacciones.Enabled = True
' mes = Month(Me.TxtFecha.Value)
' Año = Year(Me.TxtFecha.Value)
' FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
' FechaFin = DateSerial(Año, mes + 1, 1 - 1)
' NumFecha1 = FechaIni
' NumFecha2 = FechaFin
'
' Me.DCmbCodigo.Enabled = True
' Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
' Me.AdoConsulta.Refresh
' If Not AdoConsulta.Recordset.EOF Then
'  NumeroPeriodo = AdoConsulta.Recordset!NPeriodo
'  NumeroTransaccion = AdoConsulta.Recordset!NTransacciones
'  EstadoPeriodo = AdoConsulta.Recordset!EstadoPeriodo
'  If EstadoPeriodo = "B" Then
'   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
'   Me.TxtFecha.SetFocus
'
'   Exit Sub
'  ElseIf EstadoPeriodo = "C" Then
'  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
'  Me.TxtFecha.SetFocus
'  TxtFecha.Enabled = True
'
'  Exit Sub
'  Else
'   'Me.DBGTransacciones.Enabled = True
'  End If
' Else
'   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
'   'Me.DBGTransacciones.Enabled = False
'   TxtFecha.Enabled = True
'
'   Exit Sub
' End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumeroMovimiento

Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TxtFecha_LostFocus()
On Error GoTo TipoErrs
' 'Me.DBGTransacciones.Enabled = True
' mes = Month(Me.TxtFecha.Value)
' Año = Year(Me.TxtFecha.Value)
' FechaIni = CDate("1/" & Month(Me.TxtFecha.Value) & "/" & Year(Me.TxtFecha.Value))
' FechaFin = DateSerial(Año, mes + 1, 1 - 1)
' NumFecha1 = FechaIni
' NumFecha2 = FechaFin
'
' Me.DCmbCodigo.Enabled = True
' Me.AdoConsulta.RecordSource = "SELECT Periodos.NPeriodo, Periodos.NumeroTabla, Periodos.FechaPeriodo, Periodos.EstadoPeriodo, Periodos.NTransacciones, Periodos.Periodo From Periodos WHERE (((Periodos.FechaPeriodo) Between " & NumFecha1 & " And " & NumFecha2 & "))"
' Me.AdoConsulta.Refresh
' If Not AdoConsulta.Recordset.EOF Then
'  NumeroPeriodo = AdoConsulta.Recordset!NPeriodo
'  NumeroTransaccion = AdoConsulta.Recordset!NTransacciones
'  EstadoPeriodo = AdoConsulta.Recordset!EstadoPeriodo
'  If EstadoPeriodo = "B" Then
'   MsgBox "El Periodo Esta Bloqueado", vbCritical, "Sistema Contable"
'   Me.TxtFecha.SetFocus
'
'   Exit Sub
'  ElseIf EstadoPeriodo = "C" Then
'  MsgBox "El Periodo esta Cerrado", vbCritical, "Sistema Contable"
'  Me.TxtFecha.SetFocus
'  TxtFecha.Enabled = True
'
'  Exit Sub
'  Else
'   'Me.DBGTransacciones.Enabled = True
'  End If
' Else
'   MsgBox "La Fecha esta fuera del Rango de Periodos", vbCritical, "Sistema Contable"
'   'Me.DBGTransacciones.Enabled = False
'   TxtFecha.Enabled = True
'
'   Exit Sub
' End If
 
 '///////Verifico si esta registrada la fecha de la tasa//////

NumeroMovimiento

Exit Sub
TipoErrs:
 ControlErrores
End Sub
