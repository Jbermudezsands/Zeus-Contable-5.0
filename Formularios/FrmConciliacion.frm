VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmConciliacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliacion Bancaria"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   Icon            =   "FrmConciliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoConciliacion2 
      Height          =   375
      Left            =   7440
      Top             =   9360
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
      Caption         =   "AdoConciliacion2"
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Marc Todos"
      Height          =   375
      Left            =   4080
      TabIndex        =   37
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9480
      TabIndex        =   24
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   7680
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5160
      OleObjectBlob   =   "FrmConciliacion.frx":57E2
      TabIndex        =   14
      Top             =   1440
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblSaldoCuenta 
      Height          =   255
      Left            =   6720
      OleObjectBlob   =   "FrmConciliacion.frx":5866
      TabIndex        =   13
      Top             =   1440
      Width           =   3615
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
      Bindings        =   "FrmConciliacion.frx":58C4
      Height          =   4575
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8070
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "FechaTransaccion"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripcion"
      Columns(1).DataField=   "DescripcionMovimiento"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Voucher No"
      Columns(2).DataField=   "VoucherNo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Cheque No"
      Columns(3).DataField=   "ChequeNo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tasa Cambio"
      Columns(4).DataField=   "TCambio"
      Columns(4).NumberFormat=   "Edit Mask"
      Columns(4).EditMask=   "##,##0.0000"
      Columns(4).EditMaskUpdate=   -1  'True
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Debito"
      Columns(5).DataField=   "Debito"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   80
      Columns(6)._MaxComboItems=   5
      Columns(6).ValueItems(0)._DefaultItem=   0
      Columns(6).ValueItems(0).Value=   "0"
      Columns(6).ValueItems(0).Value.vt=   8
      Columns(6).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(6).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
      Columns(6).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
      Columns(6).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
      Columns(6).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
      Columns(6).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
      Columns(6).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
      Columns(6).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
      Columns(6).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(6).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(6).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
      Columns(6).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
      Columns(6).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(6).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
      Columns(6).ValueItems(0).DisplayValue.vt=   9
      Columns(6).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(6).ValueItems(1)._DefaultItem=   0
      Columns(6).ValueItems(1).Value=   "1"
      Columns(6).ValueItems(1).Value.vt=   8
      Columns(6).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(6).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
      Columns(6).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
      Columns(6).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
      Columns(6).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(6).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
      Columns(6).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(6).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
      Columns(6).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
      Columns(6).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
      Columns(6).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
      Columns(6).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
      Columns(6).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(6).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
      Columns(6).ValueItems(1).DisplayValue.vt=   9
      Columns(6).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(6).ValueItems.Count=   2
      Columns(6).Caption=   "Marca"
      Columns(6).DataField=   "Conciliada"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Credito"
      Columns(7).DataField=   "Credito"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Movimientos Conciliacion"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
   Begin MSAdodcLib.Adodc DtaConciliacion 
      Height          =   375
      Left            =   4440
      Top             =   9240
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
      Caption         =   "DtaConciliacion"
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
      Left            =   360
      Top             =   9360
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
   Begin MSAdodcLib.Adodc DtaConciliacion2 
      Height          =   375
      Left            =   240
      Top             =   9360
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
      Caption         =   "DtaConciliacion2"
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
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   10575
      Begin VB.CommandButton Command3 
         Caption         =   "&Desc Todos"
         Height          =   375
         Left            =   5160
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtDebitosDesmarcados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TxtCreditosDesmarcados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   3480
         TabIndex        =   34
         Top             =   160
         Width           =   855
      End
      Begin VB.TextBox TxtVariacion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin MSMask.MaskEdBox TxtSaldoEstado 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "##,##0.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   79036417
         CurrentDate     =   38428
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmConciliacion.frx":58E2
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmConciliacion.frx":596A
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   8760
         OleObjectBlob   =   "FrmConciliacion.frx":59F4
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   4680
         OleObjectBlob   =   "FrmConciliacion.frx":5A78
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   4680
         OleObjectBlob   =   "FrmConciliacion.frx":5AEC
         TabIndex        =   26
         Top             =   120
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   6600
         OleObjectBlob   =   "FrmConciliacion.frx":5B6C
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   6600
         OleObjectBlob   =   "FrmConciliacion.frx":5BE0
         TabIndex        =   28
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox TxtVariaciones 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TxtCredito 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TxtDebito 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmConciliacion.frx":5C62
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmConciliacion.frx":5CD6
         TabIndex        =   30
         Top             =   120
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   6480
         OleObjectBlob   =   "FrmConciliacion.frx":5D56
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   6480
         OleObjectBlob   =   "FrmConciliacion.frx":5DCA
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   8640
         OleObjectBlob   =   "FrmConciliacion.frx":5E4C
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales de la Cuenta"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.TextBox TxtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         Top             =   720
         Width           =   8655
      End
      Begin VB.ComboBox CmbTipoMoneda 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmConciliacion.frx":5ECC
         Left            =   7320
         List            =   "FrmConciliacion.frx":5ED9
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin MSDBCtls.DBCombo DBCliente 
         Bindings        =   "FrmConciliacion.frx":5EF7
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "CodCuentas"
         Text            =   ""
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmConciliacion.frx":5F10
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmConciliacion.frx":5F84
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmConciliacion.frx":5FF8
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Dim FechaConciliacion As Date
Dim CodCuenta As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

'Me.DtaConciliacion.Refresh
'Do While Not Me.DtaConciliacion.Recordset.EOF
'' Me.DtaConciliacion.Recordset.Edit
' Me.DtaConciliacion.Recordset("Conciliada") = 0
' Me.DtaConciliacion.Recordset.Update
' Me.DtaConciliacion.Recordset.MoveNext
'Loop

        FechaConciliacion = Me.dtfecha.Value
        
        CodCuenta = FrmAuxiliarCuentas.DBCliente.Text
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////////ACTUALIZO TODAS LAS OPCIONES DE CONCILIACION /////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
        rs.Open "UPDATE [Transacciones] SET [Conciliada] = 0 WHERE (((Transacciones.Conciliada)=1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102))", Conexion

Unload Me
End Sub

Private Sub CmdImprimir_Click()
ArepConciliacion.LblFechaConcilia.Caption = Me.dtfecha.Value
ArepConciliacion.LblFechaReporte.Caption = Format(Now, "dd/mm/yyyy")
ArepConciliacion.LblDebito.Caption = Me.TxtDebito.Text
ArepConciliacion.LblCredito.Caption = Me.TxtCredito.Text
ArepConciliacion.LblDescripcion.Caption = Me.TxtDescripcion.Text
ArepConciliacion.LblCuenta.Caption = Me.DBCliente.Text
ArepConciliacion.LblTotal.Caption = Format(SaldoLibros, "##,##0.00")
ArepConciliacion.LblEstadoCuenta.Caption = Format(Me.TxtSaldoEstado.Text, "##,##0.00")
ArepConciliacion.LblTotal.Caption = Format(SaldoLibros, "##,##0.00")
ArepConciliacion.LblVaracion.Caption = Me.TxtVariacion.Text
If Me.CmbTipoMoneda.Text = "Córdobas" Then
  ArepConciliacion.DataControl1.Source = "SELECT FechaTransaccion, DescripcionMovimiento, ChequeNo, VoucherNo, TCambio, Debito * TCambio AS Debito, Conciliada, TCambio * Credito AS Credito From Transacciones  " & _
                                         "WHERE (CodCuentas = '" & Me.DBCliente.Text & "') AND (NombreCuenta <> '**********CANCELADO*************') AND (Conciliada <> 1) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) ORDER BY FechaTransaccion, NumeroMovimiento"
  
Else
  ArepConciliacion.DataControl1.Source = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada, ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.MontoCordobas FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                         "WHERE (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') AND (Transacciones.Conciliada <> 1) AND (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
End If
ArepConciliacion.Show 1
End Sub

Private Sub CmdNuevo_Click()
Dim CodCuenta As String

FrmFecha.Show 1
  TipoCuenta = FrmAuxiliarCuentas.DtaCuentas.Recordset("TipoCuenta")
  Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento, Transacciones.DescripcionMovimiento, Transacciones.TCambio,Transacciones.Debito, Transacciones.Credito, Debito*TCambio AS MDebito, TCambio*Credito AS MCredito From Transacciones Where (((Transacciones.CodCuentas) =  '" & Me.DBCliente.Text & "')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
  Me.DtaConsulta.Refresh
  
  Do While Not Me.DtaConsulta.Recordset.EOF
   Me.CmbTipoMoneda.Enabled = False
   If TipoCuenta = "Activo Fijo" Or TipoCuenta = "Otros Activos" Or TipoCuenta = "Caja" Or TipoCuenta = "Cuentas x Cobrar" Or TipoCuenta = "Bancos" Or TipoCuenta = "Costos" Or TipoCuenta = "Gastos" Or TipoCuenta = "Papeleria - Utiles" Or TipoCuenta = "Inventario" Then
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
     Debito = Me.DtaConsulta.Recordset("MDebito")
    End If
    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
     Credito = Me.DtaConsulta.Recordset("MCredito")
    End If
    Total1 = Debito - Credito + Total1
    Debito = 0
    Credito = 0
   Else
    If Not IsNull(Me.DtaConsulta.Recordset("MDebito")) Then
     Debito = Me.DtaConsulta.Recordset("MDebito")
    End If
    If Not IsNull(Me.DtaConsulta.Recordset("MCredito")) Then
     Credito = Me.DtaConsulta.Recordset("MCredito")
    End If
    Total1 = Credito - Debito + Total1
    Debito = 0
    Credito = 0
   End If
   
   Me.DtaConsulta.Recordset.MoveNext
  Loop
   SaldoLibros = Total1
   Me.LblSaldoCuenta.Caption = Format(Total1, "##,##0.00")
   FrmAuxiliarCuentas.LblSaldoCuenta.Caption = Format(Total1, "##,##0.00")
   





FechaConciliacion = Me.dtfecha.Value

CodCuenta = FrmAuxiliarCuentas.DBCliente.Text

Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.ConciliacionProcesada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento "
Me.DtaConciliacion.Refresh

Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
Me.DtaConciliacion2.Refresh

  TotalDebito = 0
  TotalCredito = 0
  Debito = 0
  Credito = 0

Do While Not Me.DtaConciliacion2.Recordset.EOF
 Credito = Me.DtaConciliacion2.Recordset("Credito")
 Debito = Me.DtaConciliacion2.Recordset("Debito")
 TotalDebito = TotalDebito + Debito
 TotalCredito = TotalCredito + Credito
 Me.DtaConciliacion2.Recordset.MoveNext
Loop
Me.TxtCreditosDesmarcados.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebitosDesmarcados.Text = Format(TotalDebito, "##,##0.00")

SaldoEstadoCta = Me.TxtSaldoEstado.Text
Me.TxtVariacion.Text = Format((TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros), "##,##0.00")
'Me.TxtVariacion.Text = Format((TotalCredito - TotalDebito) + (SaldoLibros - SaldoEstadoCta), "##,##0.00")
Me.TxtVariaciones.Text = Format(TotalDebito - TotalCredito, "##,##0.00")

  




End Sub

Private Sub CmdProcesar_Click()
Dim Conciliada As Integer
If Val(Me.TxtVariacion.Text) = 0 Then
   Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.ConciliacionProcesada,Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.ConciliacionProcesada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
   Me.DtaConciliacion2.Refresh

   Do While Not Me.DtaConciliacion2.Recordset.EOF
      Conciliada = DtaConciliacion2.Recordset("Conciliada")
   If Conciliada = 1 Then
      'Me.DtaConciliacion2.Recordset.Edit
      Me.DtaConciliacion2.Recordset("ConciliacionProcesada") = 1
      Me.DtaConciliacion2.Recordset.Update

    End If
      Me.DtaConciliacion2.Recordset.MoveNext
   Loop
   Unload Me
Else
   MsgBox "La Conciliacion esta descuadrada", vbCritical, "Sistema Contable"
   Exit Sub
End If
End Sub

Private Sub Command1_Click()
Dim FechaConciliacion As Date
Dim CodCuenta As String, Variacion As Double


FechaConciliacion = Me.dtfecha.Value

CodCuenta = FrmAuxiliarCuentas.DBCliente.Text

If Me.CmbTipoMoneda.Text = "Córdobas" Then
    Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.ConciliacionProcesada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento "
    Me.DtaConciliacion.Refresh
    
    Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConciliacion2.Refresh
Else
    Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConciliacion.Refresh
    
    Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConciliacion2.Refresh
End If




Credito = 0
Debito = 0
TotalDebito = 0
TotalCredito = 0


Do While Not Me.DtaConciliacion.Recordset.EOF
 Credito = Me.DtaConciliacion.Recordset("Credito")
 Debito = Me.DtaConciliacion.Recordset("Debito")
 TotalDebito = TotalDebito + Debito
 TotalCredito = TotalCredito + Credito
 Me.DtaConciliacion.Recordset.MoveNext
Loop

Me.TxtCreditosDesmarcados.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebitosDesmarcados.Text = Format(TotalDebito, "##,##0.00")
DoEvents



If Me.CmbTipoMoneda.Text = "Córdobas" Then
        Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
                                      "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY MAX(NumeroMovimiento)"
        Me.DtaConsulta.Refresh
Else
'        Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio,SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito,SUM(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas) AS MDebito,SUM(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas) AS MCredito,SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)- Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
'                                      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
        Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 2)) AS MCredito, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 2) - ROUND(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas), 2)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
        Me.DtaConsulta.Refresh
End If



If Not Me.DtaConsulta.Recordset.EOF Then
  SaldoLibros = Me.DtaConsulta.Recordset("Saldo")
  Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
 
Else
  SaldoLibros = 0
    TotalDebito = 0
  TotalCredito = 0
  Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
End If

'-----------------------------------------------------------------------------------------------------------
'----------------------------GRABO EL ESTADO DE CUENTA DEL BANCO--------------------------------------------
'-----------------------------------------------------------------------------------------------------------
Me.AdoConciliacion2.RecordSource = "SELECT  * From Conciliacion WHERE  (FechaConciliacion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) AND (CodCuenta = '" & CodCuenta & "') AND (Activo = 1)"
Me.AdoConciliacion2.Refresh
If Not Me.AdoConciliacion2.Recordset.EOF Then
'    Me.DTFecha.Value = Me.AdoConciliacion2.Recordset("FechaConciliacion")
    Me.TxtSaldoEstado.Text = Me.AdoConciliacion2.Recordset("SaldoEstadoCuenta")
End If


If Not Me.TxtSaldoEstado.Text = "" Then
 SaldoEstadoCta = Me.TxtSaldoEstado.Text
Else
 SaldoEstadoCta = 0
End If

Variacion = (TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros)
Me.TxtVariacion.Text = Format((TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros), "##,##0.00")




Me.TDBGrid1.Columns(4).NumberFormat = "##,##0.0000"
Me.TDBGrid1.Columns(5).NumberFormat = "##,##0.00"
Me.TDBGrid1.Columns(7).NumberFormat = "##,##0.00"
Me.TDBGrid1.Columns(0).Locked = True
Me.TDBGrid1.Columns(1).Locked = True
Me.TDBGrid1.Columns(2).Locked = True
Me.TDBGrid1.Columns(3).Locked = True
Me.TDBGrid1.Columns(4).Locked = True
Me.TDBGrid1.Columns(5).Locked = True
Me.TDBGrid1.Columns(7).Locked = True
Me.TDBGrid1.Columns(0).Width = 1000
Me.TDBGrid1.Columns(1).Width = 2700
Me.TDBGrid1.Columns(2).Width = 1000
Me.TDBGrid1.Columns(3).Width = 1000
Me.TDBGrid1.Columns(4).Width = 1000
Me.TDBGrid1.Columns(5).Width = 1400
Me.TDBGrid1.Columns(6).Width = 550
Me.TDBGrid1.Columns(7).Width = 1400
End Sub

Private Sub Command2_Click()
Dim FechaConciliacion As Date
Dim CodCuenta As String, Variacion As Double
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command


        FechaConciliacion = Me.dtfecha.Value
        
        CodCuenta = FrmAuxiliarCuentas.DBCliente.Text
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////////ACTUALIZO TODAS LAS OPCIONES DE CONCILIACION /////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
        rs.Open "UPDATE [Transacciones] SET [Conciliada] = 1 WHERE (((Transacciones.Conciliada)=0) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102))", Conexion
        
        
        
        If Me.CmbTipoMoneda.Text = "Córdobas" Then
            Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.ConciliacionProcesada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento "
            Me.DtaConciliacion.Refresh
            
            Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
            Me.DtaConciliacion2.Refresh
        Else
            Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
            Me.DtaConciliacion.Refresh
            
            Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
            Me.DtaConciliacion2.Refresh
        End If
        
        
        
        
        Credito = 0
        Debito = 0
        TotalDebito = 0
        TotalCredito = 0
        
        
        Do While Not Me.DtaConciliacion.Recordset.EOF
          If Me.DtaConciliacion.Recordset("Conciliada") <> 1 Then
            If Not IsNull(Me.DtaConciliacion.Recordset("Credito")) Then
             Credito = Me.DtaConciliacion.Recordset("Credito")
            Else
             Credito = 0
            End If
            If Not IsNull(Me.DtaConciliacion.Recordset("Debito")) Then
             Debito = Me.DtaConciliacion.Recordset("Debito")
            Else
             Debito = 0
            End If
          End If
         TotalDebito = TotalDebito + Debito
         TotalCredito = TotalCredito + Credito
         Me.DtaConciliacion.Recordset.MoveNext
        Loop
        
        Me.TxtCreditosDesmarcados.Text = Format(TotalCredito, "##,##0.00")
        Me.TxtDebitosDesmarcados.Text = Format(TotalDebito, "##,##0.00")
        DoEvents
        
        
        
        If Me.CmbTipoMoneda.Text = "Córdobas" Then
                Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
                                              "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY MAX(NumeroMovimiento)"
                Me.DtaConsulta.Refresh
        Else
'                Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio,SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito,SUM(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas) AS MDebito,SUM(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas) AS MCredito,SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)- Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
'                                              "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
'
'                Me.DtaConsulta.Refresh
            Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 2)) AS MCredito, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 2) - ROUND(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas), 2)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                          "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
            Me.DtaConsulta.Refresh
        End If
        
        'Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
        '                              "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & CodCuenta & "') ORDER BY MAX(NumeroMovimiento)"
        'Me.DtaConsulta.Refresh
        
        If Not Me.DtaConsulta.Recordset.EOF Then
          SaldoLibros = Me.DtaConsulta.Recordset("Saldo")
          Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
         
        Else
          SaldoLibros = 0
            TotalDebito = 0
          TotalCredito = 0
          Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
        End If
        
        
        If Not Me.TxtSaldoEstado.Text = "" Then
         SaldoEstadoCta = Me.TxtSaldoEstado.Text
        Else
         SaldoEstadoCta = 0
        End If
        
        Variacion = (TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros)
        Me.TxtVariacion.Text = Format((TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros), "##,##0.00")
        
        Me.TDBGrid1.Columns(4).NumberFormat = "##,##0.0000"
        Me.TDBGrid1.Columns(5).NumberFormat = "##,##0.00"
        Me.TDBGrid1.Columns(7).NumberFormat = "##,##0.00"
        Me.TDBGrid1.Columns(0).Locked = True
        Me.TDBGrid1.Columns(1).Locked = True
        Me.TDBGrid1.Columns(2).Locked = True
        Me.TDBGrid1.Columns(3).Locked = True
        Me.TDBGrid1.Columns(4).Locked = True
        Me.TDBGrid1.Columns(5).Locked = True
        Me.TDBGrid1.Columns(7).Locked = True
        Me.TDBGrid1.Columns(0).Width = 1000
        Me.TDBGrid1.Columns(1).Width = 2700
        Me.TDBGrid1.Columns(2).Width = 1000
        Me.TDBGrid1.Columns(3).Width = 1000
        Me.TDBGrid1.Columns(4).Width = 1000
        Me.TDBGrid1.Columns(5).Width = 1400
        Me.TDBGrid1.Columns(6).Width = 550
        Me.TDBGrid1.Columns(7).Width = 1400
End Sub

Private Sub Command3_Click()
Dim FechaConciliacion As Date
Dim CodCuenta As String, Variacion As Double
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command


        FechaConciliacion = Me.dtfecha.Value
        
        CodCuenta = FrmAuxiliarCuentas.DBCliente.Text
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////////ACTUALIZO TODAS LAS OPCIONES DE CONCILIACION /////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
        rs.Open "UPDATE [Transacciones] SET [Conciliada] = 0 WHERE (((Transacciones.Conciliada)=1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102))", Conexion
        
        
        
        If Me.CmbTipoMoneda.Text = "Córdobas" Then
            Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.ConciliacionProcesada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento "
            Me.DtaConciliacion.Refresh
            
            Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
            Me.DtaConciliacion2.Refresh
        Else
            Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
            Me.DtaConciliacion.Refresh
            
            Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
            Me.DtaConciliacion2.Refresh
        End If
        
        
        
        
        Credito = 0
        Debito = 0
        TotalDebito = 0
        TotalCredito = 0
        
        
        Do While Not Me.DtaConciliacion.Recordset.EOF
          If Me.DtaConciliacion.Recordset("Conciliada") <> 1 Then
            If Not IsNull(Me.DtaConciliacion.Recordset("Credito")) Then
             Credito = Me.DtaConciliacion.Recordset("Credito")
            Else
             Credito = 0
            End If
            If Not IsNull(Me.DtaConciliacion.Recordset("Debito")) Then
             Debito = Me.DtaConciliacion.Recordset("Debito")
            Else
             Debito = 0
            End If
          End If
         TotalDebito = TotalDebito + Debito
         TotalCredito = TotalCredito + Credito
         Me.DtaConciliacion.Recordset.MoveNext
        Loop
        
        Me.TxtCreditosDesmarcados.Text = Format(TotalCredito, "##,##0.00")
        Me.TxtDebitosDesmarcados.Text = Format(TotalDebito, "##,##0.00")
        DoEvents
        
        
        
        If Me.CmbTipoMoneda.Text = "Córdobas" Then
                Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
                                              "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY MAX(NumeroMovimiento)"
                Me.DtaConsulta.Refresh
        Else
'                Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio,SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito,SUM(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas) AS MDebito,SUM(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas) AS MCredito,SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)- Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
'                                              "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
'
'                Me.DtaConsulta.Refresh
                Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 2)) AS MCredito, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 2) - ROUND(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas), 2)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                              "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
                Me.DtaConsulta.Refresh
        End If
        
        'Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
        '                              "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & CodCuenta & "') ORDER BY MAX(NumeroMovimiento)"
        'Me.DtaConsulta.Refresh
        
        If Not Me.DtaConsulta.Recordset.EOF Then
          SaldoLibros = Me.DtaConsulta.Recordset("Saldo")
          Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
         
        Else
          SaldoLibros = 0
            TotalDebito = 0
          TotalCredito = 0
          Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
        End If
        
        
        If Not Me.TxtSaldoEstado.Text = "" Then
         SaldoEstadoCta = Me.TxtSaldoEstado.Text
        Else
         SaldoEstadoCta = 0
        End If
        
        Variacion = (TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros)
        Me.TxtVariacion.Text = Format((TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros), "##,##0.00")
        
        Me.TDBGrid1.Columns(4).NumberFormat = "##,##0.0000"
        Me.TDBGrid1.Columns(5).NumberFormat = "##,##0.00"
        Me.TDBGrid1.Columns(7).NumberFormat = "##,##0.00"
        Me.TDBGrid1.Columns(0).Locked = True
        Me.TDBGrid1.Columns(1).Locked = True
        Me.TDBGrid1.Columns(2).Locked = True
        Me.TDBGrid1.Columns(3).Locked = True
        Me.TDBGrid1.Columns(4).Locked = True
        Me.TDBGrid1.Columns(5).Locked = True
        Me.TDBGrid1.Columns(7).Locked = True
        Me.TDBGrid1.Columns(0).Width = 1000
        Me.TDBGrid1.Columns(1).Width = 2700
        Me.TDBGrid1.Columns(2).Width = 1000
        Me.TDBGrid1.Columns(3).Width = 1000
        Me.TDBGrid1.Columns(4).Width = 1000
        Me.TDBGrid1.Columns(5).Width = 1400
        Me.TDBGrid1.Columns(6).Width = 550
        Me.TDBGrid1.Columns(7).Width = 1400
End Sub

Private Sub Form_Load()
Dim FechaConciliacion As Date, CodCuenta As String

MDIPrimero.Skin1.ApplySkin hWnd
Me.dtfecha.Value = Format(Now, "dd/mm/yyyy")
Debito = 0
Credito = 0
TotalDebito = 0
TotalCredito = 0
Me.DBCliente.Text = FrmAuxiliarCuentas.DBCliente.Text
Me.TxtDescripcion.Text = FrmAuxiliarCuentas.TxtDescripcion.Text
Me.CmbTipoMoneda.Text = FrmAuxiliarCuentas.CmbTipoMoneda.Text
Me.LblSaldoCuenta.Caption = FrmAuxiliarCuentas.LblSaldoCuenta.Caption


With Me.AdoConciliacion2
   .ConnectionString = Conexion
End With

With Me.DtaConciliacion
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With
With Me.DtaConciliacion2
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With
With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

'If Not FrmAuxiliarCuentas.LblSaldoCuenta.Caption = "" Then
'SaldoLibros = FrmAuxiliarCuentas.LblSaldoCuenta.Caption
'Else
' SaldoLibros = 0
'End If

If Me.CmbTipoMoneda.Text = "Córdobas" Then
        Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
                                      "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY MAX(NumeroMovimiento)"
        Me.DtaConsulta.Refresh
Else
        Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio,SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito,SUM(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas) AS MDebito,SUM(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas) AS MCredito,SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)- Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"

        Me.DtaConsulta.Refresh
End If



If Not Me.DtaConsulta.Recordset.EOF Then
  SaldoLibros = Me.DtaConsulta.Recordset("Saldo")
  Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")

Else
  SaldoLibros = 0
  TotalDebito = 0
  TotalCredito = 0
  Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
End If


If TipoCuenta = "Bancos" Then
  Frame2.Visible = True
  Frame3.Visible = False
Else
  Frame2.Visible = False
  Frame3.Visible = True
End If


'Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.ConciliacionProcesada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento "
'Me.DtaConciliacion.Refresh
'
'Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
'Me.DtaConciliacion2.Refresh

FechaConciliacion = Me.dtfecha.Value

CodCuenta = FrmAuxiliarCuentas.DBCliente.Text

If Me.CmbTipoMoneda.Text = "Córdobas" Then
    Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.ConciliacionProcesada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento "
    Me.DtaConciliacion.Refresh

    Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConciliacion2.Refresh
Else
    Me.DtaConciliacion.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConciliacion.Refresh

    Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo,Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada,ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.ConciliacionProcesada <> 1) AND  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConciliacion2.Refresh
End If


Credito = 0
Debito = 0
TotalDebito = 0
TotalCredito = 0

Do While Not Me.DtaConciliacion.Recordset.EOF
  If Me.DtaConciliacion.Recordset("Conciliada") <> 1 Then
    If Not IsNull(Me.DtaConciliacion.Recordset("Credito")) Then
      Credito = Me.DtaConciliacion.Recordset("Credito")
    Else
      Credito = 0
    End If
    If Not IsNull(Me.DtaConciliacion.Recordset("Debito")) Then
      Debito = Me.DtaConciliacion.Recordset("Debito")
    Else
      Debito = 0
    End If
  End If
  
 TotalDebito = TotalDebito + Debito
 TotalCredito = TotalCredito + Credito
 Me.DtaConciliacion.Recordset.MoveNext
Loop
Me.TxtCreditosDesmarcados.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebitosDesmarcados.Text = Format(TotalDebito, "##,##0.00")
''Me.TxtSaldoEstado.Text = 0

'-----------------------------------------------------------------------------------------------------------
'----------------------------GRABO EL ESTADO DE CUENTA DEL BANCO--------------------------------------------
'-----------------------------------------------------------------------------------------------------------
Me.AdoConciliacion2.RecordSource = "SELECT  * From Conciliacion WHERE (CodCuenta = '" & CodCuenta & "') AND (Activo = 1)"
Me.AdoConciliacion2.Refresh
If Not Me.AdoConciliacion2.Recordset.EOF Then
    Me.dtfecha.Value = Me.AdoConciliacion2.Recordset("FechaConciliacion")
    Me.TxtSaldoEstado.Text = Me.AdoConciliacion2.Recordset("SaldoEstadoCuenta")
    SaldoEstadoCta = Me.AdoConciliacion2.Recordset("SaldoEstadoCuenta")
End If

Me.TxtVariacion.Text = Format((TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros), "##,##0.00")

Me.TDBGrid1.Columns(4).NumberFormat = "##,##0.0000"
Me.TDBGrid1.Columns(5).NumberFormat = "##,##0.00"
Me.TDBGrid1.Columns(7).NumberFormat = "##,##0.00"
Me.TDBGrid1.Columns(0).Locked = True
Me.TDBGrid1.Columns(1).Locked = True
Me.TDBGrid1.Columns(2).Locked = True
Me.TDBGrid1.Columns(3).Locked = True
Me.TDBGrid1.Columns(4).Locked = True
Me.TDBGrid1.Columns(5).Locked = True
Me.TDBGrid1.Columns(7).Locked = True
Me.TDBGrid1.Columns(0).Width = 1000
Me.TDBGrid1.Columns(1).Width = 2700
Me.TDBGrid1.Columns(2).Width = 1000
Me.TDBGrid1.Columns(3).Width = 1000
Me.TDBGrid1.Columns(4).Width = 1000
Me.TDBGrid1.Columns(5).Width = 1400
Me.TDBGrid1.Columns(6).Width = 550
Me.TDBGrid1.Columns(7).Width = 1400
End Sub


Private Sub Label3_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label13_Click()
End Sub

Private Sub TDBGrid1_AfterUpdate()
Dim FechaConciliacion As Date
Dim CodCuenta As String

FechaConciliacion = Me.dtfecha.Value

CodCuenta = FrmAuxiliarCuentas.DBCliente.Text

If Me.CmbTipoMoneda.Text = "Córdobas" Then
   
    Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) AND (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
    Me.DtaConciliacion2.Refresh
Else

     Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Debito, Transacciones.Conciliada, ROUND(Transacciones.Credito * Transacciones.TCambio / Tasas.MontoCordobas, 2) AS Credito, Tasas.FechaTasas FROM Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas " & _
                                        "WHERE (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "') AND (Transacciones.NombreCuenta <> '**********CANCELADO*************') AND (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) AND (Transacciones.Conciliada <> 1) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
     Me.DtaConciliacion2.Refresh
End If

'Me.DtaConciliacion2.RecordSource = "SELECT Transacciones.FechaTransaccion, Transacciones.DescripcionMovimiento, Transacciones.ChequeNo, Transacciones.VoucherNo, Transacciones.TCambio, Debito*TCambio AS Debito, Transacciones.Conciliada, TCambio*Credito AS Credito From Transacciones WHERE (((Transacciones.Conciliada)<>1) AND ((Transacciones.CodCuentas)='" & Me.DBCliente.Text & "') AND ((Transacciones.NombreCuenta)<>'**********CANCELADO*************')) ORDER BY Transacciones.FechaTransaccion, Transacciones.NumeroMovimiento"
'Me.DtaConciliacion2.Refresh

  TotalDebito = 0
  TotalCredito = 0
  Credito = 0
  Debito = 0

Do While Not Me.DtaConciliacion2.Recordset.EOF
 If Not IsNull(Me.DtaConciliacion2.Recordset("Credito")) Then
  Credito = Me.DtaConciliacion2.Recordset("Credito")
 Else
  Credito = 0
 End If
 If Not IsNull(Me.DtaConciliacion2.Recordset("Debito")) Then
   Debito = Me.DtaConciliacion2.Recordset("Debito")
 Else
   Debito = 0
 End If
 TotalDebito = TotalDebito + Debito
 TotalCredito = TotalCredito + Credito
 Me.DtaConciliacion2.Recordset.MoveNext
Loop

Me.TxtCreditosDesmarcados.Text = Format(TotalCredito, "##,##0.00")
Me.TxtDebitosDesmarcados.Text = Format(TotalDebito, "##,##0.00")


FechaConciliacion = Me.dtfecha.Value
'
'Me.DtaConsulta.RecordSource = "SELECT  CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones WHERE  (FechaTransaccion <= CONVERT(DATETIME, '" & Format(FechaConciliacion, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING (CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY MAX(NumeroMovimiento)"
'Me.DtaConsulta.Refresh

If Me.CmbTipoMoneda.Text = "Córdobas" Then
        Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
                                      "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY MAX(NumeroMovimiento)"
        Me.DtaConsulta.Refresh
Else
'        Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio,SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito,SUM(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas) AS MDebito,SUM(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas) AS MCredito,SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)- Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
'                                      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
                                    
        Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 2)) AS MCredito, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 2) - ROUND(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas), 2)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
        Me.DtaConsulta.Refresh
End If

If Not Me.DtaConsulta.Recordset.EOF Then
  SaldoLibros = Me.DtaConsulta.Recordset("Saldo")
Else
  SaldoLibros = 0
End If


If Not Me.TxtSaldoEstado.Text = "" Then
 SaldoEstadoCta = Me.TxtSaldoEstado.Text
Else
 SaldoEstadoCta = 0
End If
Me.TxtVariacion.Text = Format((TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros), "##,##0.00")
'Me.TxtVariacion.Text = Format((TotalCredito - TotalDebito) + (SaldoLibros - SaldoEstadoCta), "##,##0.00")
'Me.TxtVariaciones.Text = Format(TotalDebito - TotalCredito, "##,##0.00")
End Sub

Private Sub TxtSaldoEstado_Change()

Dim FechaConciliacion As Date
Dim CodCuenta As String, Variacion As Double

FechaConciliacion = Me.dtfecha.Value
CodCuenta = FrmAuxiliarCuentas.DBCliente.Text

If Me.CmbTipoMoneda.Text = "Córdobas" Then
        Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
                                      "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & Me.DBCliente.Text & "') ORDER BY MAX(NumeroMovimiento)"
        Me.DtaConsulta.Refresh
Else
'        Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio,SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito,SUM(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas) AS MDebito,SUM(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas) AS MCredito,SUM(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas)- Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas)) AS Saldo FROM  Transacciones INNER JOIN Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
'                                      "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING  (Transacciones.CodCuentas = '" & Me.DBCliente.Text & "')"
         Me.DtaConsulta.RecordSource = "SELECT Transacciones.CodCuentas, MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.TCambio) AS TCambio, SUM(ROUND(Transacciones.Debito, 2)) AS Debito, SUM(ROUND(Transacciones.Credito, 2)) AS Credito, SUM(ROUND(Transacciones.Debito * Transacciones.TCambio / Tasas.MontoCordobas, 2)) AS MDebito, SUM(ROUND(Transacciones.TCambio * Transacciones.Credito / Tasas.MontoCordobas, 2)) AS MCredito, SUM(ROUND(Transacciones.Debito * (Transacciones.TCambio / Tasas.MontoCordobas), 2) - ROUND(Transacciones.Credito * (Transacciones.TCambio / Tasas.MontoCordobas), 2)) AS Saldo FROM Transacciones INNER JOIN  Tasas ON Transacciones.FechaTasas = Tasas.FechaTasas  " & _
                                       "WHERE (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas HAVING (Transacciones.CodCuentas =  '" & Me.DBCliente.Text & "')"
        Me.DtaConsulta.Refresh
End If

'Me.DtaConsulta.RecordSource = "SELECT CodCuentas, MAX(FechaTransaccion) AS FechaTransaccion, MAX(NumeroMovimiento) AS NumeroMovimiento, MAX(DescripcionMovimiento) AS DescripcionMovimiento, MAX(TCambio) AS TCambio, SUM(Debito) AS Debito, SUM(Credito) AS Credito, SUM(Debito * TCambio) AS MDebito, SUM(TCambio * Credito) AS MCredito, SUM(Debito * TCambio - TCambio * Credito) AS Saldo From Transacciones " & _
'                              "WHERE (FechaTransaccion <= CONVERT(DATETIME, '" & Format(Me.DTFecha.Value, "yyyy-mm-dd") & "', 102)) GROUP BY CodCuentas HAVING  (CodCuentas = '" & CodCuenta & "') ORDER BY MAX(NumeroMovimiento)"
'Me.DtaConsulta.Refresh

If Not Me.DtaConsulta.Recordset.EOF Then
  SaldoLibros = Me.DtaConsulta.Recordset("Saldo")
  Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
 
Else
  SaldoLibros = 0
  Me.LblSaldoCuenta.Caption = Format(SaldoLibros, "##,##0.00")
End If

If Not IsNumeric(Me.TxtSaldoEstado.Text) Then
  MsgBox "El monto Digitado no es numerico", vbCritical, "Zeus Facturacion"
  Exit Sub
End If


If Not Me.TxtSaldoEstado.Text = "" Then
 SaldoEstadoCta = Me.TxtSaldoEstado.Text
Else
 SaldoEstadoCta = 0
End If

Variacion = (TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros)
Me.TxtVariacion.Text = Format((TotalDebito - TotalCredito) + (SaldoEstadoCta - SaldoLibros), "##,##0.00")

'-----------------------------------------------------------------------------------------------------------
'----------------------------GRABO EL ESTADO DE CUENTA DEL BANCO--------------------------------------------
'-----------------------------------------------------------------------------------------------------------
Me.AdoConciliacion2.RecordSource = "SELECT  * From Conciliacion WHERE  (FechaConciliacion <= CONVERT(DATETIME, '" & Format(Me.dtfecha.Value, "yyyy-mm-dd") & "', 102)) AND (CodCuenta = '" & CodCuenta & "') AND (Activo = 1)"
Me.AdoConciliacion2.Refresh
If Me.AdoConciliacion2.Recordset.EOF Then
  Me.AdoConciliacion2.Recordset.AddNew
    Me.AdoConciliacion2.Recordset("FechaConciliacion") = Me.dtfecha.Value
    Me.AdoConciliacion2.Recordset("CodCuenta") = CodCuenta
    Me.AdoConciliacion2.Recordset("SaldoEstadoCuenta") = SaldoEstadoCta
    Me.AdoConciliacion2.Recordset("Activo") = True
  Me.AdoConciliacion2.Recordset.Update
Else
'    Me.AdoConciliacion2.Recordset("FechaConciliacion") = Me.DTFecha.Value
    Me.AdoConciliacion2.Recordset("SaldoEstadoCuenta") = SaldoEstadoCta
  Me.AdoConciliacion2.Recordset.Update
End If


End Sub

