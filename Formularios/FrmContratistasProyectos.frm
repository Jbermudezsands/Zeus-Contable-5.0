VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmContratistasProyectos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Proyectos"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoProyectos 
      Height          =   375
      Left            =   840
      Top             =   5880
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
      Caption         =   "AdoProyectos"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desactivar"
      Height          =   495
      Left            =   10560
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   10560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin TrueOleDBGrid80.TDBGrid DBGCuentas 
      Bindings        =   "FrmContratistasProyectos.frx":0000
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8070
      _LayoutType     =   1
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IdProyecto"
      Columns(0).DataField=   "IdProyecto"
      Columns(0).DataWidth=   11
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NombreProyecto"
      Columns(1).DataField=   "NombreProyecto"
      Columns(1).DataWidth=   50
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "CodigoContratista"
      Columns(2).DataField=   "CodigoContratista"
      Columns(2).DataWidth=   50
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "FechaContrato"
      Columns(3).DataField=   "FechaContrato"
      Columns(3).DataWidth=   19
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "FechaFinalizacion"
      Columns(4).DataField=   "FechaFinalizacion"
      Columns(4).DataWidth=   19
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Descripcion_Trabajos"
      Columns(5).DataField=   "Descripcion_Trabajos"
      Columns(5).DataWidth=   250
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "MontoContratado"
      Columns(6).DataField=   "MontoContratado"
      Columns(6).DataWidth=   23
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "PagoAnterioresManual"
      Columns(7).DataField=   "PagoAnterioresManual"
      Columns(7).DataWidth=   23
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Observaciones"
      Columns(8).DataField=   "Observaciones"
      Columns(8).DataWidth=   250
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Moneda"
      Columns(9).DataField=   "Moneda"
      Columns(9).DataWidth=   20
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Activo"
      Columns(10).DataField=   "Activo"
      Columns(10).DataWidth=   10
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Listado de Proyectos"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=131588"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=3254"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=131588"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=3096"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3016"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=131588"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(3)._AlignLeft=0"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=3096"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3016"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=131588"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(4)._AlignLeft=0"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=3493"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=3413"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=131588"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=3254"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=3175"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=131588"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(6)._AlignLeft=0"
      Splits(0)._ColumnProps(40)=   "Column(7).Width=3598"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=3519"
      Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=131588"
      Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(45)=   "Column(7)._AlignLeft=0"
      Splits(0)._ColumnProps(46)=   "Column(8).Width=3254"
      Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(49)=   "Column(8)._ColStyle=131588"
      Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(51)=   "Column(9).Width=3254"
      Splits(0)._ColumnProps(52)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(9)._WidthInPix=3175"
      Splits(0)._ColumnProps(54)=   "Column(9)._ColStyle=131588"
      Splits(0)._ColumnProps(55)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(56)=   "Column(10).Width=1667"
      Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=1588"
      Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=131588"
      Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(61)=   "Column(10)._AlignLeft=0"
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
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&HFFAEFF&"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFFA8A8&,.fgcolor=&H800080&"
      _StyleDefs(20)  =   ":id=22,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(21)  =   ":id=22,.fontname=Lucida Calligraphy"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&HECB877&"
      _StyleDefs(23)  =   ":id=14,.fgcolor=&H800000&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(24)  =   ":id=14,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=2,.bgcolor=&HFF0000&"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&H80000005&"
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
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
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
      _StyleDefs(75)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(76)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(79)  =   "Named:id=33:Normal"
      _StyleDefs(80)  =   ":id=33,.parent=0"
      _StyleDefs(81)  =   "Named:id=34:Heading"
      _StyleDefs(82)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   ":id=34,.wraptext=-1"
      _StyleDefs(84)  =   "Named:id=35:Footing"
      _StyleDefs(85)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   "Named:id=36:Selected"
      _StyleDefs(87)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(88)  =   "Named:id=37:Caption"
      _StyleDefs(89)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(90)  =   "Named:id=38:HighlightRow"
      _StyleDefs(91)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(92)  =   "Named:id=39:EvenRow"
      _StyleDefs(93)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(94)  =   "Named:id=40:OddRow"
      _StyleDefs(95)  =   ":id=40,.parent=33"
      _StyleDefs(96)  =   "Named:id=41:RecordSelector"
      _StyleDefs(97)  =   ":id=41,.parent=34"
      _StyleDefs(98)  =   "Named:id=42:FilterBar"
      _StyleDefs(99)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FrmContratistasProyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoContratista As String

Private Sub CmdNuevo_Click()
FrmContratistaNProyectos.CodigoContratista = FrmContratistasProyectos.CodigoContratista
FrmContratistaNProyectos.Show 1
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd

    With Me.AdoProyectos
       .ConnectionString = Conexion
       .RecordSource = "SELECT  ProyectosContratistas.* From ProyectosContratistas WHERE (CodigoContratista = '" & CodigoContratista & "')"
       .Refresh
    End With

End Sub
