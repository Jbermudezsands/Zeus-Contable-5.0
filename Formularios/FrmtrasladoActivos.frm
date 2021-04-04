VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmtrasladoActivos 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslado de Bienes"
   ClientHeight    =   7410
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "FrmtrasladoActivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   9075
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8895
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   2880
         OleObjectBlob   =   "FrmtrasladoActivos.frx":038A
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   6720
      Width           =   8775
      Begin VB.TextBox txtfiltrorapido 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   16
         ToolTipText     =   "Filtrar por Codigo Activo, localizacion o Nombre del Bien"
         Top             =   150
         Width           =   3825
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   5520
         TabIndex        =   17
         Top             =   120
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Guardar"
         ForeColor       =   0
         BackColor       =   -2147483633
         Appearance      =   6
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   375
         Left            =   6960
         TabIndex        =   18
         Top             =   120
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Salir"
         ForeColor       =   0
         BackColor       =   -2147483633
         Appearance      =   6
         ImageAlignment  =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmtrasladoActivos.frx":040E
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Rechazo 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   8685
      Begin MSAdodcLib.Adodc AdoHist 
         Height          =   330
         Left            =   60
         Top             =   5640
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
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
         Appearance      =   0
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
         Caption         =   "Registro 0 de 0"
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
      Begin MSAdodcLib.Adodc Adoreg 
         Height          =   330
         Left            =   0
         Top             =   5880
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
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
         Enabled         =   0
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
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
      Begin MSAdodcLib.Adodc adoactivos 
         Height          =   330
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
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
         Appearance      =   0
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
         Caption         =   "Registro 0 de 0"
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
      Begin TrueOleDBGrid80.TDBGrid DataGrid7 
         Bindings        =   "FrmtrasladoActivos.frx":0486
         Height          =   2415
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4260
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
         Splits(0).Caption=   "Detalles de  Movimientos"
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
         AllowAddNew     =   -1  'True
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
   End
   Begin VB.TextBox txtobserva 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1800
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   3465
   End
   Begin VB.TextBox txtfecha 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6240
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      MaxLength       =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2985
   End
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   840
      Width           =   615
      _Version        =   786432
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "?"
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSDataListLib.DataCombo cmdgrupo2 
      Bindings        =   "FrmtrasladoActivos.frx":049F
      DataField       =   "Idreg"
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "Descripcion"
      BoundColumn     =   "Idreg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   1200
      Width           =   375
      _Version        =   786432
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc ofic 
      Height          =   330
      Left            =   3720
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin MSDataListLib.DataCombo cmdgrupo3 
      Bindings        =   "FrmtrasladoActivos.frx":04B2
      DataField       =   "Idreg"
      Height          =   360
      Left            =   6240
      TabIndex        =   6
      Top             =   1200
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "Descripcion"
      BoundColumn     =   "Idreg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   1200
      Width           =   375
      _Version        =   786432
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc ofic3 
      Height          =   330
      Left            =   7680
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin MSDataListLib.DataCombo dtrespo 
      Bindings        =   "FrmtrasladoActivos.frx":04C5
      DataField       =   "IdReg"
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "NombreResponsable"
      BoundColumn     =   "IdReg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtrespo2 
      Bindings        =   "FrmtrasladoActivos.frx":04DC
      DataField       =   "IdReg"
      Height          =   360
      Left            =   5160
      TabIndex        =   10
      Top             =   5400
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "NombreResponsable"
      BoundColumn     =   "IdReg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnreci 
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   5400
      Width           =   375
      _Version        =   786432
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnentre 
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   5400
      Width           =   375
      _Version        =   786432
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSDataListLib.DataCombo dtrespo3 
      Bindings        =   "FrmtrasladoActivos.frx":04F3
      DataField       =   "IdReg"
      Height          =   360
      Left            =   2880
      TabIndex        =   13
      Top             =   6000
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "NombreResponsable"
      BoundColumn     =   "IdReg"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      ToolTipText     =   "Filtrar Oficinas "
      Top             =   6000
      Width           =   375
      _Version        =   786432
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   6
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc adorespo 
      Height          =   330
      Left            =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmtrasladoActivos.frx":050A
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmtrasladoActivos.frx":058A
      TabIndex        =   23
      Top             =   1200
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmtrasladoActivos.frx":0606
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "FrmtrasladoActivos.frx":067E
      TabIndex        =   25
      Top             =   840
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "FrmtrasladoActivos.frx":06E6
      TabIndex        =   26
      Top             =   1200
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "FrmtrasladoActivos.frx":0762
      TabIndex        =   28
      Top             =   5760
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   6480
      OleObjectBlob   =   "FrmtrasladoActivos.frx":07D0
      TabIndex        =   29
      Top             =   5760
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   4080
      OleObjectBlob   =   "FrmtrasladoActivos.frx":0840
      TabIndex        =   30
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5280
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "FrmtrasladoActivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idactivo As String
Private Sub btnentre_Click()
FrmResponsablesAreas.Show vbModal
End Sub

Private Sub btnreci_Click()
FrmResponsablesAreas.Show vbModal
End Sub

Private Sub DataGrid7_Click()
If adoactivos.Recordset.RecordCount = 0 Then
Else
    idactivo = adoactivos.Recordset!idactivo
    quienrecibeyentrego
    yasetraslado
End If
End Sub
Private Sub yasetraslado()
Set rsa = Nothing
SQL = "select IdUserAutoriza from TrasladoBienes where IdActivoTraslada='" & idactivo & "'"
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
If rsa.EOF Then
    dtrespo3.Locked = False
'    Label11.Visible = False
    Exit Sub
Else
    dtrespo3.BoundText = rsa!IdUserAutoriza
    dtrespo3.Locked = True
    Label11.Visible = True
    Label11.Caption = "Este Activo ya se ha dado de alta y se ha trasladado"
    
End If
End Sub
Private Sub quienrecibeyentrego()
Set rsa = Nothing
SQL = "select IdUserRecibe, IdUserEntrega from AltadeBienes where IdActivoAlta='" & idactivo & "'"
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
dtrespo.BoundText = rsa!IdUserRecibe
dtrespo2.BoundText = rsa!IdUserEntrega
 Label11.Visible = True
Label11.Caption = "Este Activo unicamente se ha dado de alta, pero no trasladado ni dado de baja"
End Sub

Private Sub Form_Activate()
cargaoficinas
generareferencia
cargaresponsables
End Sub
Private Sub cargaresponsables()
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo.Name, "Trim", Conexion, Me, "order by Idreg"
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo2.Name, "Trim", Conexion, Me, "order by Idreg"
CargaADODCConta "ResponsablesAreas", adorespo, "1", dtrespo3.Name, "Trim", Conexion, Me, "order by Idreg"
End Sub
Private Sub generareferencia()
Set rsa = Nothing
SQL = "select max (idreg) as idreg from TrasladoBienes "
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
If IsNull(rsa!idreg) Then
    Text1.Text = "000000" & 1
Else
    Text1.Text = "000000" & rsa!idreg + 1
End If
End Sub

Private Sub Form_Load()
cargaoficinas
ActivosDisponiblesAlta 1 'filtra todos los activos disponibles para ser dados de alta.
                        'Se registra en el catalogo, luego deben darse de alta
MDIPrimero.Skin1.ApplySkin hWnd
 Me.DataGrid7.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DataGrid7.OddRowStyle.BackColor = &H80000005
 Me.DataGrid7.AlternatingRowStyle = True
End Sub
Private Sub ActivosDisponiblesAlta(opcionFiltro As Integer)
    adoactivos.ConnectionString = Conexion
    adoactivos.CommandTimeout = 0
    If opcionFiltro = 1 Then 'filtra todos los activos que aun no se le han dado de alta
                             'luego de ser registrados en el catalogo
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo where (datoalta=1) and (dadobaja=0) "
    End If
    If opcionFiltro = 2 Then 'filtrado rapido, busca el activo por nombre o codigo escrito
        adoactivos.RecordSource = "select idreg as IdActivo, DescripcionAF as Descripcion, Serie as Marbete,fcompragen as Fecha_Adquisicion, Marca from dbo.CatalogoActivoFijo  where (dadoalta=1 or dadoalta='True') and (dadobaja is null or dadobaja='False' or dadobaja=0) and  (descripcionactivo LIKE '" & Trim(txtfiltrorapido.Text) & "%' or codcuenta LIKE '" & Trim(txtfiltrorapido.Text) & "%' or localizacion LIKE '" & Trim(txtfiltrorapido.Text) & "%' ) "
    End If
    adoactivos.Refresh
End Sub

Private Sub cargaoficinas()
 With Me.ofic
    .ConnectionString = Conexion
 End With
    CargaADODCConta "Oficinas", ofic, "1", cmdgrupo2.Name, "Trim", Conexion, Me, "order by Idreg"
    CargaADODCConta "Oficinas", ofic3, "1", cmdgrupo3.Name, "Trim", Conexion, Me, "order by Idreg"
End Sub

Private Sub Label9_Click()

End Sub

Private Sub PushButton1_Click()
FrmOficinas.Show
End Sub

Private Sub PushButton2_Click()
FrmResponsablesAreas.Show vbModal
End Sub

Private Sub PushButton3_Click()
FrmOficinas.Show
End Sub

Private Sub PushButton4_Click()
If txtfecha.Text = "" Or txtobserva.Text = "" Or dtrespo.Text = "" Or dtrespo2.Text = "" Or idactivo = "" Then
    MsgBox ("Informacion incompleta, por favor verifique"), vbInformation
Else
   guardatras
   actualizaestadoactivo
   ActivosDisponiblesAlta 1
End If
End Sub
Private Sub guardatras()
Set rsa = Nothing
SQL = "select * from TrasladoBienes "
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
rsa.AddNew
rsa!IdReferencia = Text1.Text
rsa!FechaGraba = Format(CDate(txtfecha.Text), "YYYY/MM/DD")
rsa!IdOfiOrigen = DepartamentoID(cmdgrupo2.BoundText)
rsa!DescriOficina = cmdgrupo2.Text
rsa!IdOfiDestino = DepartamentoID(cmdgrupo3.Text)
rsa!DescriOficinaDest = cmdgrupo3.Text
rsa!Observaciones = txtobserva.Text
rsa!IdUserRecibe = ResponsableID(dtrespo.Text)
rsa!NombreRecibe = dtrespo.Text
rsa!IdUserEntrega = ResponsableID(dtrespo2.Text)
rsa!NombreEntrega = dtrespo2.Text
rsa!IdUserAutoriza = ResponsableID(dtrespo3.Text)
rsa!NombreAutoriza = dtrespo3.Text
rsa!IdActivoTraslada = idactivo
rsa.Update
PushButton4.Enabled = False
End Sub
Private Sub actualizaestadoactivo()
Set rsa = Nothing
SQL = "update CatalogoActivoFijo set Trasladado=1, Fechatraslado='" & Format(Now, "dd/MM/yyyy") & "'  where Idreg='" & Trim(idactivo) & "'"
rsa.Open SQL, Conexion, adOpenForwardOnly, adLockOptimistic
End Sub

Private Sub PushButton5_Click()
     wfecha = IIf(Len(Trim(txtfecha.Text)) = 0 Or Not IsDate(txtfecha.Text), Date, txtfecha.Text)
    Set wforma = Me
    wtextc = txtfecha.Name
    whabfe = True
    On Local Error Resume Next
    Load fcalendario
    On Local Error GoTo 0
    If fcalendario.Visible = False Then fcalendario.Show vbModal
End Sub

Private Sub PushButton6_Click()
Unload Me
End Sub

Private Sub txtfiltrorapido_Change()
If Not IsNumeric(txtfiltrorapido.Text) Then 'Significa que esta escribiendo el nombre del activo
                                        'o el numero de codigo del mismo
    ActivosDisponiblesAlta 2 'busca el filtro del activo
Else
    If txtfiltrorapido.Text = "" Then
        ActivosDisponiblesAlta 1
    End If
End If
End Sub
